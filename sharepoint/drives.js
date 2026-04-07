/**
 * SharePoint document libraries (drives) operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List document libraries in a site
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListDrives(args) {
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

    const response = await callGraphAPI(accessToken, 'GET', `sites/${siteId}/drives`);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No document libraries found in this site."
        }]
      };
    }

    const driveList = response.value.map((drive, index) => {
      const usedBytes = drive.quota && drive.quota.used ? formatSize(drive.quota.used) : 'N/A';
      const totalBytes = drive.quota && drive.quota.total ? formatSize(drive.quota.total) : 'N/A';
      return `${index + 1}. ${drive.name}\n   ID: ${drive.id}\n   URL: ${drive.webUrl}\n   Type: ${drive.driveType || 'N/A'}\n   Usage: ${usedBytes} / ${totalBytes}`;
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} document library/libraries:\n\n${driveList}`
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
        text: `Error listing document libraries: ${error.message}`
      }]
    };
  }
}

/**
 * Get a specific drive/document library
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleGetDrive(args) {
  const driveId = args.driveId;

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

    const response = await callGraphAPI(accessToken, 'GET', `drives/${driveId}`);

    if (!response || !response.id) {
      return {
        content: [{
          type: "text",
          text: "Drive not found."
        }]
      };
    }

    const usedBytes = response.quota && response.quota.used ? formatSize(response.quota.used) : 'N/A';
    const totalBytes = response.quota && response.quota.total ? formatSize(response.quota.total) : 'N/A';
    const remainingBytes = response.quota && response.quota.remaining ? formatSize(response.quota.remaining) : 'N/A';

    const details = [
      `Name: ${response.name}`,
      `ID: ${response.id}`,
      `URL: ${response.webUrl}`,
      `Type: ${response.driveType || 'N/A'}`,
      `Owner: ${response.owner && response.owner.user ? response.owner.user.displayName : (response.owner && response.owner.group ? response.owner.group.displayName : 'N/A')}`,
      `Usage: ${usedBytes} / ${totalBytes}`,
      `Remaining: ${remainingBytes}`,
      `Created: ${response.createdDateTime ? new Date(response.createdDateTime).toLocaleString() : 'N/A'}`,
      `Last Modified: ${response.lastModifiedDateTime ? new Date(response.lastModifiedDateTime).toLocaleString() : 'N/A'}`
    ].join('\n');

    return {
      content: [{
        type: "text",
        text: `Drive details:\n\n${details}`
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
        text: `Error getting drive: ${error.message}`
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
  handleListDrives,
  handleGetDrive
};
