/**
 * SharePoint sites operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List all sites the user has access to
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListSites(args) {
  const count = args.count || 25;

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      search: '*',
      $top: Math.min(50, count)
    };

    const response = await callGraphAPI(accessToken, 'GET', 'sites', null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No SharePoint sites found."
        }]
      };
    }

    const siteList = response.value.map((site, index) => {
      return `${index + 1}. ${site.displayName || site.name}\n   URL: ${site.webUrl}\n   ID: ${site.id}\n   Description: ${site.description || 'N/A'}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} site(s):\n\n${siteList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing sites: ${error.message}`
      }]
    };
  }
}

/**
 * Get a specific site by ID or URL
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetSite(args) {
  const siteId = args.siteId;
  const hostname = args.hostname;
  const sitePath = args.sitePath;

  if (!siteId && !hostname) {
    return {
      content: [{
        type: "text",
        text: "Either siteId or hostname (with optional sitePath) is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    let endpoint;
    if (siteId) {
      endpoint = `sites/${siteId}`;
    } else if (hostname && sitePath) {
      const normalizedPath = sitePath.replace(/^\/+|\/+$/g, '');
      endpoint = `sites/${hostname}:/${normalizedPath}`;
    } else {
      endpoint = `sites/${hostname}`;
    }

    const response = await callGraphAPI(accessToken, 'GET', endpoint);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Site not found."
        }]
      };
    }

    const details = [
      `Name: ${response.displayName || response.name}`,
      `ID: ${response.id}`,
      `URL: ${response.webUrl}`,
      `Description: ${response.description || 'N/A'}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Last Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`
    ].join('\n');

    return {
      content: [{
        type: "text",
        text: `Site details:\n\n${details}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error getting site: ${error.message}`
      }]
    };
  }
}

/**
 * Search for sites by keyword
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchSites(args) {
  const query = args.query;
  const count = args.count || 25;

  if (!query) {
    return {
      content: [{
        type: "text",
        text: "Search query is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      search: query,
      $top: Math.min(50, count)
    };

    const response = await callGraphAPI(accessToken, 'GET', 'sites', null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No sites found matching "${query}".`
        }]
      };
    }

    const siteList = response.value.map((site, index) => {
      return `${index + 1}. ${site.displayName || site.name}\n   URL: ${site.webUrl}\n   ID: ${site.id}\n   Description: ${site.description || 'N/A'}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} site(s) matching "${query}":\n\n${siteList}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error searching sites: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleListSites,
  handleGetSite,
  handleSearchSites
};
