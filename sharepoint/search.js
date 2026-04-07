/**
 * SharePoint search operations using Microsoft Search API
 * POST /search/query — semantic/KQL search across SharePoint content
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Search across all SharePoint content using Microsoft Search API
 */
async function handleSearch(args) {
  const query = args.query;
  const entityTypes = args.entityTypes || ['driveItem'];
  const count = args.count || 25;
  const siteId = args.siteId;
  const contentSources = args.contentSources;

  if (!query) {
    return {
      content: [{ type: "text", text: "query is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const searchRequest = {
      entityTypes: entityTypes,
      query: {
        queryString: query
      },
      from: 0,
      size: Math.min(50, count),
      fields: [
        'id', 'name', 'title', 'webUrl', 'lastModifiedDateTime',
        'lastModifiedBy', 'createdDateTime', 'createdBy',
        'size', 'parentReference', 'siteId', 'listId'
      ]
    };

    // Scope to specific site if provided
    if (siteId) {
      searchRequest.query.queryString = `${query} siteId:${siteId}`;
    }

    // External content sources (for connectors)
    if (contentSources && Array.isArray(contentSources) && contentSources.length > 0) {
      searchRequest.contentSources = contentSources;
    }

    // Enable result trimming and hit highlights
    searchRequest.enableTopResults = true;

    const body = { requests: [searchRequest] };

    const response = await callGraphAPI(accessToken, 'POST', 'search/query', body);

    if (!response.value || response.value.length === 0 || !response.value[0].hitsContainers) {
      return {
        content: [{ type: "text", text: `No results found for "${query}".` }]
      };
    }

    const hitsContainer = response.value[0].hitsContainers[0];

    if (!hitsContainer.hits || hitsContainer.hits.length === 0) {
      return {
        content: [{ type: "text", text: `No results found for "${query}".` }]
      };
    }

    const total = hitsContainer.total || hitsContainer.hits.length;
    const moreAvailable = hitsContainer.moreResultsAvailable || false;

    const results = hitsContainer.hits.map((hit, index) => {
      const resource = hit.resource || {};
      const summary = hit.summary || '';
      const rank = hit.rank || index + 1;

      const lines = [`${index + 1}. ${resource.name || resource.title || 'Untitled'}`];

      if (resource.webUrl) lines.push(`   URL: ${resource.webUrl}`);
      if (resource.size) lines.push(`   Size: ${formatSize(resource.size)}`);
      if (resource.lastModifiedDateTime) lines.push(`   Modified: ${new Date(resource.lastModifiedDateTime).toLocaleString()}`);
      if (resource.lastModifiedBy && resource.lastModifiedBy.user) {
        lines.push(`   Modified by: ${resource.lastModifiedBy.user.displayName}`);
      }
      if (resource.parentReference) {
        if (resource.parentReference.siteId) lines.push(`   Site ID: ${resource.parentReference.siteId}`);
        if (resource.parentReference.driveId) lines.push(`   Drive ID: ${resource.parentReference.driveId}`);
      }
      if (resource.id) lines.push(`   ID: ${resource.id}`);
      if (summary) lines.push(`   Summary: ${summary.replace(/<[^>]*>/g, '')}`);

      return lines.join('\n');
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${total} result(s) for "${query}"${moreAvailable ? ' (more available)' : ''}:\n\n${results}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Error searching: ${error.message}` }]
    };
  }
}

/**
 * Search specifically for list items across SharePoint
 */
async function handleSearchListItems(args) {
  const query = args.query;
  const count = args.count || 25;
  const siteId = args.siteId;

  if (!query) {
    return {
      content: [{ type: "text", text: "query is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const searchRequest = {
      entityTypes: ['listItem'],
      query: {
        queryString: siteId ? `${query} siteId:${siteId}` : query
      },
      from: 0,
      size: Math.min(50, count),
      enableTopResults: true
    };

    const body = { requests: [searchRequest] };
    const response = await callGraphAPI(accessToken, 'POST', 'search/query', body);

    if (!response.value || response.value.length === 0 || !response.value[0].hitsContainers) {
      return {
        content: [{ type: "text", text: `No list items found for "${query}".` }]
      };
    }

    const hitsContainer = response.value[0].hitsContainers[0];

    if (!hitsContainer.hits || hitsContainer.hits.length === 0) {
      return {
        content: [{ type: "text", text: `No list items found for "${query}".` }]
      };
    }

    const total = hitsContainer.total || hitsContainer.hits.length;

    const results = hitsContainer.hits.map((hit, index) => {
      const resource = hit.resource || {};
      const fields = resource.fields || {};
      const summary = hit.summary || '';

      const lines = [`${index + 1}. ${fields.title || resource.name || 'Untitled'}`];

      if (resource.webUrl) lines.push(`   URL: ${resource.webUrl}`);
      if (resource.lastModifiedDateTime) lines.push(`   Modified: ${new Date(resource.lastModifiedDateTime).toLocaleString()}`);
      if (resource.parentReference) {
        if (resource.parentReference.siteId) lines.push(`   Site ID: ${resource.parentReference.siteId}`);
        if (resource.parentReference.listId) lines.push(`   List ID: ${resource.parentReference.listId}`);
      }
      if (resource.id) lines.push(`   ID: ${resource.id}`);

      // Show field values
      const fieldEntries = Object.entries(fields)
        .filter(([key]) => !key.startsWith('@') && !key.startsWith('_') && key !== 'title')
        .slice(0, 8)
        .map(([key, value]) => `   ${key}: ${typeof value === 'object' ? JSON.stringify(value) : value}`);
      if (fieldEntries.length > 0) lines.push(...fieldEntries);

      if (summary) lines.push(`   Summary: ${summary.replace(/<[^>]*>/g, '')}`);

      return lines.join('\n');
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${total} list item(s) for "${query}":\n\n${results}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Error searching list items: ${error.message}` }]
    };
  }
}

/**
 * Search across multiple entity types at once (files, list items, sites)
 */
async function handleSearchAll(args) {
  const query = args.query;
  const count = args.count || 10;

  if (!query) {
    return {
      content: [{ type: "text", text: "query is required." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = {
      requests: [
        {
          entityTypes: ['driveItem'],
          query: { queryString: query },
          from: 0,
          size: Math.min(25, count),
          enableTopResults: true
        },
        {
          entityTypes: ['listItem'],
          query: { queryString: query },
          from: 0,
          size: Math.min(25, count),
          enableTopResults: true
        },
        {
          entityTypes: ['site'],
          query: { queryString: query },
          from: 0,
          size: Math.min(10, count),
          enableTopResults: true
        }
      ]
    };

    const response = await callGraphAPI(accessToken, 'POST', 'search/query', body);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: `No results found for "${query}".` }]
      };
    }

    const sections = [];

    const entityLabels = ['Files', 'List Items', 'Sites'];

    response.value.forEach((result, idx) => {
      const container = result.hitsContainers && result.hitsContainers[0];
      if (!container || !container.hits || container.hits.length === 0) return;

      const label = entityLabels[idx] || `Results ${idx + 1}`;
      const total = container.total || container.hits.length;

      const items = container.hits.map((hit, i) => {
        const r = hit.resource || {};
        const summary = hit.summary ? ` — ${hit.summary.replace(/<[^>]*>/g, '')}` : '';
        return `  ${i + 1}. ${r.name || r.displayName || r.title || 'Untitled'}${summary}\n     URL: ${r.webUrl || 'N/A'}\n     ID: ${r.id || 'N/A'}`;
      }).join('\n');

      sections.push(`── ${label} (${total} found) ──\n${items}`);
    });

    if (sections.length === 0) {
      return {
        content: [{ type: "text", text: `No results found for "${query}".` }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Search results for "${query}":\n\n${sections.join('\n\n')}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }
    return {
      content: [{ type: "text", text: `Error searching: ${error.message}` }]
    };
  }
}

function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = {
  handleSearch,
  handleSearchListItems,
  handleSearchAll
};
