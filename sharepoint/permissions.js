/**
 * SharePoint permissions and sharing operations
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Share a file or folder by creating a sharing link
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleShareItem(args) {
  const driveId = args.driveId;
  const itemId = args.itemId;
  const type = args.type || 'view';
  const scope = args.scope || 'organization';

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

    const body = {
      type: type,
      scope: scope
    };

    const response = await callGraphAPI(accessToken, 'POST', `drives/${driveId}/items/${itemId}/createLink`, body);

    if (!response || !response.link) {
      return {
        content: [{
          type: "text",
          text: "Failed to create sharing link."
        }]
      };
    }

    const linkInfo = response.link;

    return {
      content: [{
        type: "text",
        text: `Sharing link created:\n\nLink: ${linkInfo.webUrl}\nType: ${type}\nScope: ${scope}\n\nNote: ${scope === 'anonymous' ? 'Anyone with this link can access the item.' : 'Only people in your organization can access.'}`
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
        text: `Error creating sharing link: ${error.message}`
      }]
    };
  }
}

/**
 * List permissions on an item
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListPermissions(args) {
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

    const response = await callGraphAPI(accessToken, 'GET', `drives/${driveId}/items/${itemId}/permissions`);

    if (!response.value || response.value.length === 0) {
      return {
        content: [{
          type: "text",
          text: "No permissions found for this item."
        }]
      };
    }

    const permList = response.value.map((perm, index) => {
      const parts = [`${index + 1}. Permission ID: ${perm.id}`];

      if (perm.roles) {
        parts.push(`   Roles: ${perm.roles.join(', ')}`);
      }

      if (perm.grantedTo && perm.grantedTo.user) {
        parts.push(`   Granted To: ${perm.grantedTo.user.displayName} (${perm.grantedTo.user.email || 'N/A'})`);
      }

      if (perm.grantedToIdentities) {
        const identities = perm.grantedToIdentities.map(identity => {
          if (identity.user) return identity.user.displayName || identity.user.email;
          return 'Unknown';
        }).join(', ');
        parts.push(`   Granted To: ${identities}`);
      }

      if (perm.link) {
        parts.push(`   Link Type: ${perm.link.type}`);
        parts.push(`   Link Scope: ${perm.link.scope}`);
        if (perm.link.webUrl) {
          parts.push(`   Link URL: ${perm.link.webUrl}`);
        }
      }

      if (perm.inheritedFrom) {
        parts.push(`   Inherited: Yes`);
      }

      return parts.join('\n');
    }).join("\n\n");

    return {
      content: [{
        type: "text",
        text: `Found ${response.value.length} permission(s):\n\n${permList}`
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
        text: `Error listing permissions: ${error.message}`
      }]
    };
  }
}

module.exports = {
  handleShareItem,
  handleListPermissions
};
