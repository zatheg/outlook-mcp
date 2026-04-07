/**
 * SharePoint files and folders operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List files/folders in a drive or folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListItems(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;
  const count = args.count || 25;

  if (!driveId) {
    return {
      content: [{
        type: "text",
        text: "driveId is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    let endpoint;
    if (itemId) {
      endpoint = `drives/${driveId}/items/${itemId}/children`;
    } else {
      endpoint = `drives/${driveId}/root/children`;
    }

    const queryParams = {
      $top: Math.min(50, count),
      $orderby: 'name'
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No files or folders found."
        }]
      };
    }

    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : 'N/A';

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Modified: ${modified}\n   ID: ${item.id}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s):\n\n${fileList}`
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
        text: `Error listing items: ${error.message}`
      }]
    };
  }
}

/**
 * Get metadata for a file or folder
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetItem(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;

  if (!driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Both driveId and itemId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const response = await callGraphAPI(accessToken, 'GET', `drives/${driveId}/items/${itemId}`);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Item not found."
        }]
      };
    }

    const isFolder = !!response.folder;
    const details = [
      `Name: ${response.name}`,
      `Type: ${isFolder ? 'Folder' : 'File'}`,
      `ID: ${response.id}`,
      `Size: ${response.size ? formatSize(response.size) : 'N/A'}`,
      `URL: ${response.webUrl}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`,
      `Created By: ${response.createdBy && response.createdBy.user ? response.createdBy.user.displayName : 'N/A'}`,
      `Modified By: ${response.lastModifiedBy && response.lastModifiedBy.user ? response.lastModifiedBy.user.displayName : 'N/A'}`
    ];

    if (isFolder && response.folder) {
      details.push(`Child Count: ${response.folder.childCount || 0}`);
    }

    if (response.file) {
      details.push(`MIME Type: ${response.file.mimeType || 'N/A'}`);
    }

    return {
      content: [{
        type: "text",
        text: `Item details:\n\n${details.join('\n')}`
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
        text: `Error getting item: ${error.message}`
      }]
    };
  }
}

/**
 * Get download URL for a file
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDownload(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;

  if (!driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Both driveId and itemId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get item metadata first (which includes the download URL)
    const response = await callGraphAPI(accessToken, 'GET', `drives/${driveId}/items/${itemId}`);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Item not found."
        }]
      };
    }

    const downloadUrl = response['@microsoft.graph.downloadUrl'];

    if (!downloadUrl) {
      return {
        content: [{
          type: "text",
          text: `Item "${response.name}" does not have a downloadable content (it may be a folder).`
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Download URL for "${response.name}":\n\n${downloadUrl}\n\nSize: ${response.size ? formatSize(response.size) : 'N/A'}\nMIME Type: ${response.file ? response.file.mimeType : 'N/A'}\n\nNote: This URL is temporary and will expire.`
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
        text: `Error getting download URL: ${error.message}`
      }]
    };
  }
}

/**
 * Upload a file to SharePoint
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleUpload(args) {
  const driveId = args.driveId;
  const parentId = args.parentId;
  const parentPath = args.parentPath;
  const filename = args.filename;
  const content = args.content;

  if (!driveId) {
    return {
      content: [{
        type: "text",
        text: "driveId is required."
      }]
    };
  }

  if (!filename) {
    return {
      content: [{
        type: "text",
        text: "filename is required."
      }]
    };
  }

  if (content === undefined || content === null) {
    return {
      content: [{
        type: "text",
        text: "content is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    let endpoint;
    if (parentId) {
      endpoint = `drives/${driveId}/items/${parentId}:/${filename}:/content`;
    } else if (parentPath) {
      const normalizedPath = parentPath.replace(/^\/+|\/+$/g, '');
      endpoint = `drives/${driveId}/root:/${normalizedPath}/${filename}:/content`;
    } else {
      endpoint = `drives/${driveId}/root:/${filename}:/content`;
    }

    const response = await callGraphAPI(accessToken, 'PUT', endpoint, content);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to upload file."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully uploaded "${response.name}"\n\nID: ${response.id}\nSize: ${response.size ? formatSize(response.size) : 'N/A'}\nURL: ${response.webUrl}`
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
        text: `Error uploading file: ${error.message}`
      }]
    };
  }
}

/**
 * Create a folder in SharePoint
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCreateFolder(args) {
  const driveId = args.driveId;
  const parentId = args.parentId;
  const name = args.name;

  if (!driveId) {
    return {
      content: [{
        type: "text",
        text: "driveId is required."
      }]
    };
  }

  if (!name) {
    return {
      content: [{
        type: "text",
        text: "Folder name is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    let endpoint;
    if (parentId) {
      endpoint = `drives/${driveId}/items/${parentId}/children`;
    } else {
      endpoint = `drives/${driveId}/root/children`;
    }

    const body = {
      name: name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    };

    const response = await callGraphAPI(accessToken, 'POST', endpoint, body);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to create folder."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Successfully created folder "${response.name}"\n\nID: ${response.id}\nURL: ${response.webUrl}`
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
        text: `Error creating folder: ${error.message}`
      }]
    };
  }
}

/**
 * Delete a file or folder from SharePoint
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleDeleteItem(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;

  if (!driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Both driveId and itemId are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    // Get item info first
    const itemInfo = await callGraphAPI(accessToken, 'GET', `drives/${driveId}/items/${itemId}`);

    if (!itemInfo || !itemInfo.id) {
      return {
        content: [{
          type: "text",
          text: "Item not found."
        }]
      };
    }

    const itemName = itemInfo.name;
    const isFolder = !!itemInfo.folder;

    // Delete the item
    await callGraphAPI(accessToken, 'DELETE', `drives/${driveId}/items/${itemId}`);

    return {
      content: [{
        type: "text",
        text: `Successfully deleted ${isFolder ? 'folder' : 'file'} "${itemName}".`
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
        text: `Error deleting item: ${error.message}`
      }]
    };
  }
}

/**
 * Move or rename a file/folder in SharePoint
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveItem(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;
  const newName = args.newName;
  const destinationParentId = args.destinationParentId;
  const destinationDriveId = args.destinationDriveId;

  if (!driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Both driveId and itemId are required."
      }]
    };
  }

  if (!newName && !destinationParentId) {
    return {
      content: [{
        type: "text",
        text: "At least one of newName or destinationParentId is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = {};

    if (newName) {
      body.name = newName;
    }

    if (destinationParentId) {
      body.parentReference = {
        id: destinationParentId
      };
      if (destinationDriveId) {
        body.parentReference.driveId = destinationDriveId;
      }
    }

    const response = await callGraphAPI(accessToken, 'PATCH', `drives/${driveId}/items/${itemId}`, body);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Failed to move/rename item."
        }]
      };
    }

    const actions = [];
    if (newName) actions.push(`renamed to "${response.name}"`);
    if (destinationParentId) actions.push('moved to new location');

    return {
      content: [{
        type: "text",
        text: `Successfully ${actions.join(' and ')}.\n\nID: ${response.id}\nName: ${response.name}\nURL: ${response.webUrl}`
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
        text: `Error moving/renaming item: ${error.message}`
      }]
    };
  }
}

/**
 * Copy a file in SharePoint
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCopyItem(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;
  const destinationParentId = args.destinationParentId;
  const destinationDriveId = args.destinationDriveId;
  const newName = args.newName;

  if (!driveId || !itemId) {
    return {
      content: [{
        type: "text",
        text: "Both driveId and itemId are required."
      }]
    };
  }

  if (!destinationParentId) {
    return {
      content: [{
        type: "text",
        text: "destinationParentId is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const body = {
      parentReference: {
        id: destinationParentId
      }
    };

    if (destinationDriveId) {
      body.parentReference.driveId = destinationDriveId;
    }

    if (newName) {
      body.name = newName;
    }

    await callGraphAPI(accessToken, 'POST', `drives/${driveId}/items/${itemId}/copy`, body);

    return {
      content: [{
        type: "text",
        text: `Copy operation initiated successfully.${newName ? ` New name: "${newName}"` : ''}\n\nNote: The copy operation is asynchronous. The file will appear in the destination shortly.`
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
        text: `Error copying item: ${error.message}`
      }]
    };
  }
}

/**
 * Search files across a site
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleSearchFiles(args) {
  const siteId = args.siteId;
  const query = args.query;
  const count = args.count || 25;

  if (!siteId || !query) {
    return {
      content: [{
        type: "text",
        text: "Both siteId and query are required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const queryParams = {
      $top: Math.min(50, count)
    };

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/drive/root/search(q='${query}')`, null, queryParams);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No files found matching "${query}".`
        }]
      };
    }

    const fileList = response.value.map((item, index) => {
      const isFolder = item.folder ? '[FOLDER]' : '[FILE]';
      const size = item.size ? formatSize(item.size) : '';
      const modified = item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).toLocaleString() : 'N/A';

      return `${index + 1}. ${isFolder} ${item.name}${size ? ` (${size})` : ''}\n   Modified: ${modified}\n   ID: ${item.id}\n   Drive ID: ${item.parentReference ? item.parentReference.driveId : 'N/A'}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} item(s) matching "${query}":\n\n${fileList}`
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
        text: `Error searching files: ${error.message}`
      }]
    };
  }
}

/**
 * Format file size to human-readable string
 */
function formatSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

module.exports = {
  handleListItems,
  handleGetItem,
  handleDownload,
  handleUpload,
  handleCreateFolder,
  handleDeleteItem,
  handleMoveItem,
  handleCopyItem,
  handleSearchFiles
};
