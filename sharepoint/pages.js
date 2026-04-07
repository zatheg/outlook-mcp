/**
 * SharePoint pages operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List site pages
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListPages(args) {
  const siteId = args.siteId;
  const count = args.count || 25;

  if (!siteId) {
    return {
      content: [{
        type: "text",
        text: "siteId is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      $top: Math.min(50, count)
    };

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/pages`, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No pages found in this site."
        }]
      };
    }

    const pageList = response.value.map((page, index) => {
      return `${index + 1}. ${page.title || page.name}\n   ID: ${page.id}\n   URL: ${page.webUrl}\n   Created: ${page.createdDateTime ? new Date(page.createdDateTime).toLocaleString() : 'N/A'}\n   Modified: ${page.lastModifiedDateTime ? new Date(page.lastModifiedDateTime).toLocaleString() : 'N/A'}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} page(s):\n\n${pageList}`
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
        text: `Error listing pages: ${error.message}`
      }]
    };
  }
}

/**
 * Get a specific page
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetPage(args) {
  const siteId = args.siteId;
  const pageId = args.pageId;

  if (!siteId || !pageId) {
    return {
      content: [{
        type: "text",
        text: "Both siteId and pageId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/pages/${pageId}`);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Page not found."
        }]
      };
    }

    const details = [
      `Title: ${response.title || response.name}`,
      `ID: ${response.id}`,
      `URL: ${response.webUrl}`,
      `Description: ${response.description || 'N/A'}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`,
      `Created By: ${response.createdBy && response.createdBy.user ? response.createdBy.user.displayName : 'N/A'}`,
      `Modified By: ${response.lastModifiedBy && response.lastModifiedBy.user ? response.lastModifiedBy.user.displayName : 'N/A'}`,
      `Publishing State: ${response.publishingState ? response.publishingState.level : 'N/A'}`
    ].join('\n');

    return {
      content: [{
        type: "text",
        text: `Page details:\n\n${details}`
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
        text: `Error getting page: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleListPages,
  handleGetPage
};
