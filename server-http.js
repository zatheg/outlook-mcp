#!/usr/bin/env node
/**
 * Cloud Run HTTP entry point for M365 Assistant MCP Server
 * Uses StreamableHTTP transport instead of stdio
 * Tokens are persisted via GCP Secret Manager
 * OAuth 2.1 Authorization Code + PKCE for Claude.ai/Cowork authentication
 */

// Must set HOME before any module loads config.js (which resolves token path at load time)
if (!process.env.HOME || process.env.HOME === '/root') {
  process.env.HOME = '/tmp';
}

const CLOUD_RUN_URL = 'https://outlook-mcp-465123644836.northamerica-northeast1.run.app';
if (!process.env.MS_REDIRECT_URI) {
  process.env.MS_REDIRECT_URI = `${CLOUD_RUN_URL}/auth/callback`;
}
// Ensure OUTLOOK_CLIENT_ID is set from MS_CLIENT_ID (config.js only checks OUTLOOK_*)
if (!process.env.OUTLOOK_CLIENT_ID && process.env.MS_CLIENT_ID) {
  process.env.OUTLOOK_CLIENT_ID = process.env.MS_CLIENT_ID;
}
if (!process.env.OUTLOOK_CLIENT_SECRET && process.env.MS_CLIENT_SECRET) {
  process.env.OUTLOOK_CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
}

const crypto = require('crypto');
const express = require('express');
const fs = require('fs');
const path = require('path');
const { Server } = require("@modelcontextprotocol/sdk/server/index.js");
const { StreamableHTTPServerTransport } = require("@modelcontextprotocol/sdk/server/streamableHttp.js");
const config = require('./config');

const { authTools } = require('./auth');
const { calendarTools } = require('./calendar');
const { emailTools } = require('./email');
const { folderTools } = require('./folder');
const { rulesTools } = require('./rules');
const { onedriveTools } = require('./onedrive');
const { powerAutomateTools } = require('./power-automate');
const { sharepointTools } = require('./sharepoint');

const TOOLS = [
  ...authTools, ...calendarTools, ...emailTools, ...folderTools,
  ...rulesTools, ...onedriveTools, ...powerAutomateTools, ...sharepointTools,
];

// ── OAuth token lifetime ─────────────────────────────────────────────────────
const OAUTH_TOKEN_EXPIRES_IN = 180 * 24 * 60 * 60; // 180 days in seconds

// ── OAuth stores (persisted to Secret Manager on Cloud Run) ─────────────────
const registeredClients = new Map(); // client_id -> { client_id, client_name, redirect_uris }
const authCodes = new Map();         // code -> { client_id, redirect_uri, code_challenge, expires_at }
const activeTokens = new Map();      // token -> { client_id, expires_at }
const refreshTokens = new Map();     // refresh_token -> { client_id }

// ── Secret Manager persistence for OAuth tokens ─────────────────────────────
let _smClient = null;
let _oauthSecretParent = null;

async function _getSecretManagerClient() {
  if (_smClient) return _smClient;
  const projectId = process.env.GCP_PROJECT_ID;
  if (!projectId) return null;
  try {
    const { SecretManagerServiceClient } = require('@google-cloud/secret-manager');
    _smClient = new SecretManagerServiceClient();
    _oauthSecretParent = `projects/${projectId}/secrets/${process.env.OAUTH_SECRET_NAME || 'outlook-mcp-oauth-tokens'}`;
    return _smClient;
  } catch (e) {
    console.error('Secret Manager client unavailable for OAuth tokens:', e.message);
    return null;
  }
}

function _serializeOAuthState() {
  return JSON.stringify({
    activeTokens: Array.from(activeTokens.entries()),
    refreshTokens: Array.from(refreshTokens.entries()),
    registeredClients: Array.from(registeredClients.entries()),
  });
}

async function loadOAuthTokens() {
  const client = await _getSecretManagerClient();
  if (!client) return;
  try {
    const [version] = await client.accessSecretVersion({ name: `${_oauthSecretParent}/versions/latest` });
    const data = JSON.parse(version.payload.data.toString('utf8'));
    const now = Date.now();
    if (data.activeTokens) {
      for (const [k, v] of data.activeTokens) {
        if (now < v.expires_at) activeTokens.set(k, v);
      }
    }
    if (data.refreshTokens) {
      for (const [k, v] of data.refreshTokens) refreshTokens.set(k, v);
    }
    if (data.registeredClients) {
      for (const [k, v] of data.registeredClients) registeredClients.set(k, v);
    }
    console.error(`OAuth tokens loaded from Secret Manager (${activeTokens.size} active, ${refreshTokens.size} refresh)`);
  } catch (e) {
    if (e.code === 5) console.error('OAuth token secret not found — first auth will create it');
    else console.error('Could not load OAuth tokens from Secret Manager:', e.message);
  }
}

// Persist OAuth state to Secret Manager. Returns a promise that resolves once
// THIS call's write reaches Secret Manager, so callers can `await` before
// responding (critical: Cloud Run may SIGTERM the instance before any deferred
// write fires, and clients holding a never-persisted refresh token would then
// fail to refresh against the next instance, forcing full re-auth).
//
// Writes are serialised via a promise chain so concurrent callers don't race
// to publish a stale snapshot. Each link captures `_serializeOAuthState()` at
// the moment it runs, after the previous write completes.
let _persistChain = Promise.resolve();
function persistOAuthState() {
  const next = _persistChain.then(async () => {
    const client = await _getSecretManagerClient();
    if (!client) return;
    try {
      await client.addSecretVersion({
        parent: _oauthSecretParent,
        payload: { data: Buffer.from(_serializeOAuthState()) },
      });
    } catch (e) {
      // If secret doesn't exist yet, create it
      if (e.code === 5) {
        try {
          await client.createSecret({
            parent: _oauthSecretParent.replace(/\/secrets\/.*/, ''),
            secretId: process.env.OAUTH_SECRET_NAME || 'outlook-mcp-oauth-tokens',
            secret: { replication: { automatic: {} } },
          });
          await client.addSecretVersion({
            parent: _oauthSecretParent,
            payload: { data: Buffer.from(_serializeOAuthState()) },
          });
          console.error('OAuth token secret created and persisted');
        } catch (createErr) {
          console.error('Failed to create OAuth token secret:', createErr.message);
        }
      } else {
        console.error('Failed to persist OAuth tokens:', e.message);
      }
    }
  });
  // Swallow errors on the chain itself so one failure doesn't poison later writes
  _persistChain = next.catch(() => {});
  return next;
}

// Allowed redirect URIs for Claude.ai and Claude Code
const ALLOWED_REDIRECT_PATTERNS = [
  'https://claude.ai/',
  'https://claude.com/',
  'http://localhost:',
  'http://127.0.0.1:',
];

function isAllowedRedirectUri(uri) {
  return ALLOWED_REDIRECT_PATTERNS.some((pattern) => uri.startsWith(pattern));
}

function cleanupExpired() {
  const now = Date.now();
  for (const [k, v] of authCodes) { if (now > v.expires_at) authCodes.delete(k); }
  for (const [k, v] of activeTokens) { if (now > v.expires_at) activeTokens.delete(k); }
}

// ── MCP Server factory ──────────────────────────────────────────────────────
function createMcpServer() {
  const server = new Server(
    { name: config.SERVER_NAME, version: config.SERVER_VERSION },
    { capabilities: { tools: {} } }
  );

  server.fallbackRequestHandler = async (request) => {
    try {
      const { method, params, id } = request;
      console.error(`REQUEST: ${method} [${id}]`);

      if (method === "initialize") {
        return {
          protocolVersion: "2025-11-25",
          capabilities: { tools: {} },
          serverInfo: { name: config.SERVER_NAME, version: config.SERVER_VERSION },
        };
      }
      if (method === "tools/list") {
        return { tools: TOOLS.map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema })) };
      }
      if (method === "resources/list") return { resources: [] };
      if (method === "prompts/list") return { prompts: [] };
      if (method === "tools/call") {
        const { name, arguments: args = {} } = params || {};
        console.error(`TOOL CALL: ${name}`);
        const tool = TOOLS.find((t) => t.name === name);
        if (tool && tool.handler) return await tool.handler(args);
        return { error: { code: -32601, message: `Tool not found: ${name}` } };
      }
      return { error: { code: -32601, message: `Method not found: ${method}` } };
    } catch (error) {
      console.error(`Error in fallbackRequestHandler:`, error);
      return { error: { code: -32603, message: `Error: ${error.message}` } };
    }
  };

  return server;
}

// ── GCP Secret Manager token sync ────────────────────────────────────────────
async function initGcpTokens() {
  const projectId = process.env.GCP_PROJECT_ID;
  if (!projectId) return;

  const secretName = process.env.TOKEN_SECRET_NAME || 'outlook-mcp-tokens';
  const tokenPath = path.join('/tmp', '.outlook-mcp-tokens.json');

  let client;
  try {
    const { SecretManagerServiceClient } = require('@google-cloud/secret-manager');
    client = new SecretManagerServiceClient();
  } catch (e) {
    console.error('Secret Manager client unavailable:', e.message);
    return;
  }

  try {
    const name = `projects/${projectId}/secrets/${secretName}/versions/latest`;
    const [version] = await client.accessSecretVersion({ name });
    const tokenJson = version.payload.data.toString('utf8');
    fs.writeFileSync(tokenPath, tokenJson, { mode: 0o600 });
    console.error('Tokens loaded from Secret Manager');
  } catch (e) {
    if (e.code === 5) console.error(`Secret '${secretName}' not found — OAuth authentication required first`);
    else console.error('Could not load tokens from Secret Manager:', e.message);
    // Don't return — still apply fs patches so OAuth callback tokens get synced
  }

  // Patch writes to sync back to Secret Manager (always, even if initial load failed)
  const origWriteFileAsync = fs.promises.writeFile;
  fs.promises.writeFile = async (filePath, data, options) => {
    await origWriteFileAsync(filePath, data, options);
    if (filePath === tokenPath) {
      try {
        await client.addSecretVersion({
          parent: `projects/${projectId}/secrets/${secretName}`,
          payload: { data: Buffer.from(typeof data === 'string' ? data : JSON.stringify(data)) },
        });
        console.error('Tokens synced to Secret Manager');
      } catch (e) { console.error('Failed to sync tokens:', e.message); }
    }
  };

  const origWriteFileSync = fs.writeFileSync;
  fs.writeFileSync = (filePath, data, options) => {
    origWriteFileSync(filePath, data, options);
    if (filePath === tokenPath) {
      client.addSecretVersion({
        parent: `projects/${projectId}/secrets/${secretName}`,
        payload: { data: Buffer.from(typeof data === 'string' ? data : JSON.stringify(data)) },
      }).catch((e) => console.error('Failed to sync tokens (sync):', e.message));
    }
  };
}

// ── Bearer token middleware ──────────────────────────────────────────────────
// resource_metadata in WWW-Authenticate points MCP clients (Claude Code, Claude.ai)
// at our PRM so they can discover scopes_supported (notably offline_access) and
// silently refresh access tokens instead of forcing a fresh browser flow.
const PRM_URL = `${CLOUD_RUN_URL}/.well-known/oauth-protected-resource`;
function requireBearerToken(req, res, next) {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.set('WWW-Authenticate', `Bearer resource_metadata="${PRM_URL}"`);
    return res.status(401).json({ error: 'unauthorized', error_description: 'Bearer token required' });
  }
  const token = authHeader.slice(7);
  const tokenData = activeTokens.get(token);
  if (!tokenData || Date.now() > tokenData.expires_at) {
    activeTokens.delete(token);
    res.set('WWW-Authenticate', `Bearer error="invalid_token", resource_metadata="${PRM_URL}"`);
    return res.status(401).json({ error: 'invalid_token', error_description: 'Token expired or invalid' });
  }
  next();
}

// ── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  await initGcpTokens();
  await loadOAuthTokens();

  const TokenStorage = require('./auth/token-storage');
  const { setupOAuthRoutes, createAuthConfig } = require('./auth/oauth-server');
  const tokenStorage = new TokenStorage();
  const msAuthConfig = createAuthConfig('MS_');

  const app = express();
  app.use(express.json());
  app.use(express.urlencoded({ extended: true }));

  // ── Health ──
  app.get('/', (req, res) => res.send('OK'));
  app.get('/health', (req, res) => res.send('OK'));

  // ── OAuth 2.1 Authorization Server Metadata (RFC 8414) ──
  // scopes_supported MUST include offline_access — Claude Code only appends
  // it to the authorization request when it sees it advertised here, which is
  // what unlocks silent refresh.
  app.get('/.well-known/oauth-authorization-server', (req, res) => {
    res.json({
      issuer: CLOUD_RUN_URL,
      authorization_endpoint: `${CLOUD_RUN_URL}/authorize`,
      token_endpoint: `${CLOUD_RUN_URL}/token`,
      registration_endpoint: `${CLOUD_RUN_URL}/register`,
      scopes_supported: ['openid', 'offline_access'],
      response_types_supported: ['code'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      token_endpoint_auth_methods_supported: ['none'],
      code_challenge_methods_supported: ['S256'],
    });
  });

  // ── OAuth 2.0 Protected Resource Metadata (RFC 9728) ──
  // Claude Code prefers PRM over OASM for scope discovery. Serve at both the
  // bare path and the /mcp-suffixed path so strict and lenient clients both
  // resolve to the same metadata.
  const prmPayload = {
    resource: `${CLOUD_RUN_URL}/mcp`,
    authorization_servers: [CLOUD_RUN_URL],
    scopes_supported: ['openid', 'offline_access'],
    bearer_methods_supported: ['header'],
  };
  app.get('/.well-known/oauth-protected-resource', (req, res) => res.json(prmPayload));
  app.get('/.well-known/oauth-protected-resource/mcp', (req, res) => res.json(prmPayload));

  // ── Dynamic Client Registration (RFC 7591) ──
  app.post('/register', async (req, res) => {
    const { client_name, redirect_uris } = req.body;

    // Validate redirect URIs
    if (redirect_uris && Array.isArray(redirect_uris)) {
      for (const uri of redirect_uris) {
        if (!isAllowedRedirectUri(uri)) {
          return res.status(400).json({ error: 'invalid_redirect_uri', error_description: `Redirect URI not allowed: ${uri}` });
        }
      }
    }

    const clientId = crypto.randomUUID();
    const clientData = {
      client_id: clientId,
      client_name: client_name || 'Unknown',
      redirect_uris: redirect_uris || [],
      client_id_issued_at: Math.floor(Date.now() / 1000),
    };
    registeredClients.set(clientId, clientData);
    await persistOAuthState();
    console.error(`Registered OAuth client: ${client_name} (${clientId})`);

    res.status(201).json(clientData);
  });

  // ── Authorization Endpoint ──
  app.get('/authorize', (req, res) => {
    const { response_type, client_id, redirect_uri, state, code_challenge, code_challenge_method, scope } = req.query;

    if (response_type !== 'code') {
      return res.status(400).send('Unsupported response_type. Must be "code".');
    }

    if (!registeredClients.has(client_id)) {
      return res.status(400).send('Unknown client_id. Register first via /register.');
    }

    if (!redirect_uri || !isAllowedRedirectUri(redirect_uri)) {
      return res.status(400).send('Invalid redirect_uri.');
    }

    if (!code_challenge || code_challenge_method !== 'S256') {
      return res.status(400).send('PKCE required. Provide code_challenge with S256 method.');
    }

    // Show consent page
    const escapedState = (state || '').replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const escapedRedirect = redirect_uri.replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const escapedChallenge = code_challenge.replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const escapedClientId = client_id.replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const clientName = (registeredClients.get(client_id).client_name || 'Unknown').replace(/</g, '&lt;');

    res.send(`<!DOCTYPE html>
<html><head><title>Authorize - M365 MCP</title></head>
<body style="font-family:system-ui;max-width:500px;margin:80px auto;text-align:center">
  <h1>Authorize Access</h1>
  <p><strong>${clientName}</strong> wants to access your M365 Assistant (Outlook, Calendar, OneDrive).</p>
  <form method="POST" action="/authorize">
    <input type="hidden" name="client_id" value="${escapedClientId}">
    <input type="hidden" name="redirect_uri" value="${escapedRedirect}">
    <input type="hidden" name="state" value="${escapedState}">
    <input type="hidden" name="code_challenge" value="${escapedChallenge}">
    <button type="submit" style="padding:12px 32px;font-size:16px;background:#2ecc71;color:white;border:none;border-radius:6px;cursor:pointer">
      Authorize
    </button>
  </form>
</body></html>`);
  });

  // ── Authorization POST (user approves) ──
  app.post('/authorize', (req, res) => {
    const { client_id, redirect_uri, state, code_challenge } = req.body;

    if (!registeredClients.has(client_id) || !redirect_uri || !isAllowedRedirectUri(redirect_uri)) {
      return res.status(400).send('Invalid request.');
    }

    // Generate authorization code
    const code = crypto.randomBytes(32).toString('hex');
    authCodes.set(code, {
      client_id,
      redirect_uri,
      code_challenge,
      expires_at: Date.now() + 10 * 60 * 1000, // 10 min
    });

    cleanupExpired();

    const url = new URL(redirect_uri);
    url.searchParams.set('code', code);
    if (state) url.searchParams.set('state', state);

    console.error(`Authorization code issued for client ${client_id}`);
    res.redirect(url.toString());
  });

  // ── Token Endpoint ──
  // Persist BEFORE responding so the new (rotated) refresh token reaches Secret
  // Manager before the client could use it against a different Cloud Run instance.
  app.post('/token', async (req, res) => {
    const { grant_type, code, redirect_uri, code_verifier, client_id, refresh_token } = req.body;

    // ── Refresh token grant ──
    if (grant_type === 'refresh_token') {
      if (!refresh_token || !refreshTokens.has(refresh_token)) {
        return res.status(400).json({ error: 'invalid_grant', error_description: 'Invalid refresh token' });
      }
      const rtData = refreshTokens.get(refresh_token);
      refreshTokens.delete(refresh_token);

      const accessToken = crypto.randomBytes(32).toString('hex');
      const newRefreshToken = crypto.randomBytes(32).toString('hex');
      const expiresIn = OAUTH_TOKEN_EXPIRES_IN;
      activeTokens.set(accessToken, { client_id: rtData.client_id, expires_at: Date.now() + expiresIn * 1000 });
      refreshTokens.set(newRefreshToken, { client_id: rtData.client_id });
      await persistOAuthState();

      return res.json({ access_token: accessToken, token_type: 'Bearer', expires_in: expiresIn, refresh_token: newRefreshToken });
    }

    // ── Authorization code grant ──
    if (grant_type !== 'authorization_code') {
      return res.status(400).json({ error: 'unsupported_grant_type' });
    }

    if (!code || !authCodes.has(code)) {
      return res.status(400).json({ error: 'invalid_grant', error_description: 'Invalid or expired authorization code' });
    }

    const codeData = authCodes.get(code);
    authCodes.delete(code);

    if (Date.now() > codeData.expires_at) {
      return res.status(400).json({ error: 'invalid_grant', error_description: 'Authorization code expired' });
    }

    if (codeData.redirect_uri !== redirect_uri) {
      return res.status(400).json({ error: 'invalid_grant', error_description: 'redirect_uri mismatch' });
    }

    // Validate PKCE
    if (!code_verifier) {
      return res.status(400).json({ error: 'invalid_request', error_description: 'code_verifier required' });
    }
    const expectedChallenge = crypto.createHash('sha256').update(code_verifier).digest('base64url');
    if (expectedChallenge !== codeData.code_challenge) {
      return res.status(400).json({ error: 'invalid_grant', error_description: 'PKCE verification failed' });
    }

    // Issue tokens
    const accessToken = crypto.randomBytes(32).toString('hex');
    const newRefreshToken = crypto.randomBytes(32).toString('hex');
    const expiresIn = OAUTH_TOKEN_EXPIRES_IN;

    activeTokens.set(accessToken, { client_id: codeData.client_id, expires_at: Date.now() + expiresIn * 1000 });
    refreshTokens.set(newRefreshToken, { client_id: codeData.client_id });
    await persistOAuthState();

    console.error(`Access token issued for client ${codeData.client_id}`);
    res.json({ access_token: accessToken, token_type: 'Bearer', expires_in: expiresIn, refresh_token: newRefreshToken });
  });

  // ── Microsoft OAuth routes (for initial MS login): /auth and /auth/callback ──
  setupOAuthRoutes(app, tokenStorage, msAuthConfig);

  // ── MCP endpoint (protected) ──
  app.all('/mcp', requireBearerToken, async (req, res) => {
    try {
      const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
      const server = createMcpServer();
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (e) {
      console.error('Error handling MCP request:', e);
      if (!res.headersSent) {
        res.status(500).json({ error: e.message });
      }
    }
  });

  const PORT = parseInt(process.env.PORT || '8080', 10);
  const server = app.listen(PORT, '0.0.0.0', () => {
    console.error(`${config.SERVER_NAME} HTTP server listening on port ${PORT}`);
    console.error(`MS OAuth login: ${CLOUD_RUN_URL}/auth`);
    console.error(`MCP OAuth metadata: ${CLOUD_RUN_URL}/.well-known/oauth-authorization-server`);
  });

  // Drain in-flight persists on SIGTERM (Cloud Run gives ~10s grace before SIGKILL).
  // Without this, a refresh token that just rotated could be lost if the instance
  // shuts down mid-write, and the client's next refresh would fail.
  const shutdown = async (signal) => {
    console.error(`${signal} received — flushing OAuth state and closing server`);
    try {
      await _persistChain;
    } catch (e) {
      console.error('Error draining persist chain:', e.message);
    }
    server.close(() => process.exit(0));
    // Hard timeout as a safety net
    setTimeout(() => process.exit(0), 8000).unref();
  };
  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
}

main().catch((e) => {
  console.error('Fatal error:', e);
  process.exit(1);
});
