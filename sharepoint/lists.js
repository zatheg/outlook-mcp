/**
 * SharePoint lists operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List all lists in a site
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListLists(args) {
  const siteId = args.siteId;

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

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/lists`);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No lists found in this site."
        }]
      };
    }

    const listList = response.value.map((list, index) => {
      return `${index + 1}. ${list.displayName}\n   ID: ${list.id}\n   Template: ${list.list ? list.list.template : 'N/A'}\n   URL: ${list.webUrl}\n   Created: ${list.createdDateTime ? new Date(list.createdDateTime).toLocaleString() : 'N/A'}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} list(s):\n\n${listList}`
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
        text: `Error listing lists: ${error.message}`
      }]
    };
  }
}

/**
 * Get a specific list
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetList(args) {
  const siteId = args.siteId;
  const listId = args.listId;

  if (!siteId || !listId) {
    return {
      content: [{
        type: "text",
        text: "Both siteId and listId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/lists/${listId}`);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "List not found."
        }]
      };
    }

    const details = [
      `Name: ${response.displayName}`,
      `ID: ${response.id}`,
      `Description: ${response.description || 'N/A'}`,
      `URL: ${response.webUrl}`,
      `Template: ${response.list ? response.list.template : 'N/A'}`,
      `Hidden: ${response.list ? response.list.hidden : 'N/A'}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Last Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`
    ].join('\n');

    return {
      content: [{
        type: "text",
        text: `List details:\n\n${details}`
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
        text: `Error getting list: ${error.message}`
      }]
    };
  }
}

/**
 * Get items from a list
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListListItems(args) {
  const siteId = args.siteId;
  const listId = args.listId;
  const count = args.count || 25;
  const filter = args.filter;

  if (!siteId || !listId) {
    return {
      content: [{
        type: "text",
        text: "Both siteId and listId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      $expand: 'fields',
      $top: Math.min(50, count)
    };

    if (filter) {
      queryParams.$filter = filter;
    }

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/lists/${listId}/items`, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No items found in this list."
        }]
      };
    }

    const itemList = response.value.map((item, index) => {
      const fields = item.fields || {};
      const fieldEntries = Object.entries(fields)
        .filter(([key]) => !key.startsWith('@') && !key.startsWith('_'))
        .map(([key, value]) => `   ${key}: ${value}`)
        .join('\n');

      return `${index + 1}. Item ID: ${item.id}\n   Created: ${item.createdDateTime ? new Date(item.createdDateTime).toLocaleString() : 'N/A'}\n   Modified: ${item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : 'N/A'}\n${fieldEntries}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s):\n\n${itemList}`
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
        text: `Error listing list items: ${error.message}`
      }]
    };
  }
}

/**
 * Get a specific list item
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetListItem(args) {
  const siteId = args.siteId;
  const listId = args.listId;
  const itemId = args.itemId;

  if (!siteId || !listId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "siteId, listId, and itemId are all required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      $expand: 'fields'
    };

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/lists/${listId}/items/${itemId}`, null, queryParams);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "List item not found."
        }]
      };
    }

    const fields = response.fields || {};
    const fieldEntries = Object.entries(fields)
      .filter(([key]) => !key.startsWith('@') && !key.startsWith('_'))
      .map(([key, value]) => `${key}: ${typeof value === 'object' ? JSON.stringify(value) : value}`)
      .join('\n');

    const details = [
      `Item ID: ${response.id}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`,
      `URL: ${response.webUrl || 'N/A'}`,
      '',
      'Fields:',
      fieldEntries
    ].join('\n');

    return {
      content: [{
        type: "text",
        text: `List item details:\n\n${details}`
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
        text: `Error getting list item: ${error.message}`
      }]
    };
  }
}

/**
 * Create a new list item
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateListItem(args) {
  const siteId = args.siteId;
  const listId = args.listId;
  const fields = args.fields;

  if (!siteId || !listId) {
    return {
      content: [{
        type: "text",
        text: "Both siteId and listId are required."
      }]
    };
  }

  if (!fields || typeof fields !== 'object') {
    return {
      content: [{
        type: "text",
        text: "fields is required and must be an object with field name/value pairs."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = {
      fields: fields
    };

    const response = await callGraphAPI(accessToken, 'POST', `sites/${siteId}/lists/${listId}/items`, body);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to create list item."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully created list item.\n\nItem ID: ${response.id}\nCreated: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`
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
        text: `Error creating list item: ${error.message}`
      }]
    };
  }
}

/**
 * Update a list item
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpdateListItem(args) {
  const siteId = args.siteId;
  const listId = args.listId;
  const itemId = args.itemId;
  const fields = args.fields;

  if (!siteId || !listId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "siteId, listId, and itemId are all required."
      }]
    };
  }

  if (!fields || typeof fields !== 'object') {
    return {
      content: [{
        type: "text",
        text: "fields is required and must be an object with field name/value pairs."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'PATCH', `sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);

    if (!response) {
      return {
        content: [{
          type: "text",
          text: "Failed to update list item."
        }]
      };
    }

    const updatedFields = Object.entries(response)
      .filter(([key]) => !key.startsWith('@') && !key.startsWith('_'))
      .map(([key, value]) => `${key}: ${typeof value === 'object' ? JSON.stringify(value) : value}`)
      .join('\n');

    return {
      content: [{
        type: "text",
        text: `Successfully updated list item ${itemId}.\n\nUpdated fields:\n${updatedFields}`
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
        text: `Error updating list item: ${error.message}`
      }]
    };
  }
}

/**
 * Delete a list item
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteListItem(args) {
  const siteId = args.siteId;
  const listId = args.listId;
  const itemId = args.itemId;

  if (!siteId || !listId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "siteId, listId, and itemId are all required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    await callGraphAPI(accessToken, 'DELETE', `sites/${siteId}/lists/${listId}/items/${itemId}`);

    return {
      content: [{
        type: "text",
        text: `Successfully deleted list item ${itemId}.`
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
        text: `Error deleting list item: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleListLists,
  handleGetList,
  handleListListItems,
  handleGetListItem,
  handleCreateListItem,
  handleUpdateListItem,
  handleDeleteListItem
};
