/**
 * SharePoint module for Outlook MCP server
 */
const { handleListSites, handleGetSite, handleSearchSites } = require('./sites');
const { handleListDrives, handleGetDrive } = require('./drives');
const { handleListItems, handleGetItem, handleDownload, handleUpload, handleCreateFolder, handleDeleteItem, handleMoveItem, handleCopyItem, handleSearchFiles } = require('./files');
const { handleListLists, handleGetList, handleListListItems, handleGetListItem, handleCreateListItem, handleUpdateListItem, handleDeleteListItem } = require('./lists');
const { handleListPages, handleGetPage } = require('./pages');
const { handleShareItem, handleListPermissions } = require('./permissions');
const { handleSearch, handleSearchListItems, handleSearchAll } = require('./search');

// SharePoint tool definitions
const sharepointTools = [
  // ===== Sites =====
  {
    name: "sharepoint-list-sites",
    description: "List all SharePoint sites the user has access to",
    inputSchema: {
      type: "object",
      properties: {
        count: {
          type: "number",
          description: "Number of sites to retrieve (default: 25, max: 50)"
        }
      },
      required: []
    },
    handler: handleListSites
  },
  {
    name: "sharepoint-get-site",
    description: "Get details of a specific SharePoint site by ID or by hostname and path",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The unique ID of the site (e.g., 'contoso.sharepoint.com,2C712604-1370-44E7-A1F5-426573FDA80A,2D2244C3-251A-49EA-93A8-39E1C3A060FE')"
        },
        hostname: {
          type: "string",
          description: "The hostname of the SharePoint site (e.g., 'contoso.sharepoint.com'). Use with sitePath for path-based lookup."
        },
        sitePath: {
          type: "string",
          description: "The server-relative path of the site (e.g., '/sites/marketing'). Used together with hostname."
        }
      },
      required: []
    },
    handler: handleGetSite
  },
  {
    name: "sharepoint-search-sites",
    description: "Search for SharePoint sites by keyword",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search keyword to find sites"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 25, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchSites
  },

  // ===== Document Libraries (Drives) =====
  {
    name: "sharepoint-list-drives",
    description: "List all document libraries (drives) in a SharePoint site",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        }
      },
      required: ["siteId"]
    },
    handler: handleListDrives
  },
  {
    name: "sharepoint-get-drive",
    description: "Get details of a specific document library (drive) by its ID",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        }
      },
      required: ["driveId"]
    },
    handler: handleGetDrive
  },

  // ===== Files & Folders =====
  {
    name: "sharepoint-list-items",
    description: "List files and folders in a SharePoint document library or folder. Provide driveId to list root, or driveId + itemId to list a specific folder's contents.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the folder to list contents of. If omitted, lists the root of the drive."
        },
        count: {
          type: "number",
          description: "Number of items to retrieve (default: 25, max: 50)"
        }
      },
      required: ["driveId"]
    },
    handler: handleListItems
  },
  {
    name: "sharepoint-get-item",
    description: "Get metadata for a specific file or folder in a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder"
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleGetItem
  },
  {
    name: "sharepoint-download",
    description: "Get a temporary download URL for a file in a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the file to download"
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleDownload
  },
  {
    name: "sharepoint-upload",
    description: "Upload a small file (< 4MB) to a SharePoint document library. Specify the destination using parentId or parentPath, or omit both to upload to root.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        parentId: {
          type: "string",
          description: "The ID of the parent folder to upload into. If omitted, uploads to root or parentPath."
        },
        parentPath: {
          type: "string",
          description: "The path of the parent folder (e.g., '/General/Reports'). Alternative to parentId."
        },
        filename: {
          type: "string",
          description: "The name of the file to create (e.g., 'report.txt')"
        },
        content: {
          type: "string",
          description: "The file content to upload"
        }
      },
      required: ["driveId", "filename", "content"]
    },
    handler: handleUpload
  },
  {
    name: "sharepoint-create-folder",
    description: "Create a new folder in a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        parentId: {
          type: "string",
          description: "The ID of the parent folder. If omitted, creates in the root of the drive."
        },
        name: {
          type: "string",
          description: "The name of the new folder"
        }
      },
      required: ["driveId", "name"]
    },
    handler: handleCreateFolder
  },
  {
    name: "sharepoint-delete-item",
    description: "Delete a file or folder from a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to delete"
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleDeleteItem
  },
  {
    name: "sharepoint-move-item",
    description: "Move or rename a file/folder in a SharePoint document library. Provide newName to rename, destinationParentId to move, or both.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library containing the item"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to move/rename"
        },
        newName: {
          type: "string",
          description: "The new name for the item (for renaming)"
        },
        destinationParentId: {
          type: "string",
          description: "The ID of the destination folder (for moving)"
        },
        destinationDriveId: {
          type: "string",
          description: "The ID of the destination drive, if moving across drives"
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleMoveItem
  },
  {
    name: "sharepoint-copy-item",
    description: "Copy a file or folder to a new location in SharePoint. The copy operation is asynchronous.",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library containing the item"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to copy"
        },
        destinationParentId: {
          type: "string",
          description: "The ID of the destination folder"
        },
        destinationDriveId: {
          type: "string",
          description: "The ID of the destination drive, if copying across drives"
        },
        newName: {
          type: "string",
          description: "Optional new name for the copied item"
        }
      },
      required: ["driveId", "itemId", "destinationParentId"]
    },
    handler: handleCopyItem
  },
  {
    name: "sharepoint-search-files",
    description: "Search for files across a SharePoint site's default document library",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site to search in"
        },
        query: {
          type: "string",
          description: "Search query to find files by name or content"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 25, max: 50)"
        }
      },
      required: ["siteId", "query"]
    },
    handler: handleSearchFiles
  },

  // ===== Lists =====
  {
    name: "sharepoint-list-lists",
    description: "List all SharePoint lists in a site",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        }
      },
      required: ["siteId"]
    },
    handler: handleListLists
  },
  {
    name: "sharepoint-get-list",
    description: "Get details of a specific SharePoint list",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        }
      },
      required: ["siteId", "listId"]
    },
    handler: handleGetList
  },
  {
    name: "sharepoint-list-list-items",
    description: "Get items from a SharePoint list with their field values",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        },
        count: {
          type: "number",
          description: "Number of items to retrieve (default: 25, max: 50)"
        },
        filter: {
          type: "string",
          description: "OData filter expression to filter items (e.g., \"fields/Status eq 'Active'\")"
        }
      },
      required: ["siteId", "listId"]
    },
    handler: handleListListItems
  },
  {
    name: "sharepoint-get-list-item",
    description: "Get a specific item from a SharePoint list with all its field values",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        },
        itemId: {
          type: "string",
          description: "The ID of the list item"
        }
      },
      required: ["siteId", "listId", "itemId"]
    },
    handler: handleGetListItem
  },
  {
    name: "sharepoint-create-list-item",
    description: "Create a new item in a SharePoint list",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        },
        fields: {
          type: "object",
          description: "An object containing field name/value pairs for the new item (e.g., {\"Title\": \"My Item\", \"Status\": \"Active\"})"
        }
      },
      required: ["siteId", "listId", "fields"]
    },
    handler: handleCreateListItem
  },
  {
    name: "sharepoint-update-list-item",
    description: "Update an existing item in a SharePoint list",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        },
        itemId: {
          type: "string",
          description: "The ID of the list item to update"
        },
        fields: {
          type: "object",
          description: "An object containing field name/value pairs to update (e.g., {\"Status\": \"Completed\"})"
        }
      },
      required: ["siteId", "listId", "itemId", "fields"]
    },
    handler: handleUpdateListItem
  },
  {
    name: "sharepoint-delete-list-item",
    description: "Delete an item from a SharePoint list",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        listId: {
          type: "string",
          description: "The ID or name of the list"
        },
        itemId: {
          type: "string",
          description: "The ID of the list item to delete"
        }
      },
      required: ["siteId", "listId", "itemId"]
    },
    handler: handleDeleteListItem
  },

  // ===== Pages =====
  {
    name: "sharepoint-list-pages",
    description: "List all pages in a SharePoint site",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        count: {
          type: "number",
          description: "Number of pages to retrieve (default: 25, max: 50)"
        }
      },
      required: ["siteId"]
    },
    handler: handleListPages
  },
  {
    name: "sharepoint-get-page",
    description: "Get details of a specific page in a SharePoint site",
    inputSchema: {
      type: "object",
      properties: {
        siteId: {
          type: "string",
          description: "The ID of the SharePoint site"
        },
        pageId: {
          type: "string",
          description: "The ID of the page"
        }
      },
      required: ["siteId", "pageId"]
    },
    handler: handleGetPage
  },

  // ===== Permissions =====
  {
    name: "sharepoint-share-item",
    description: "Create a sharing link for a file or folder in a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder to share"
        },
        type: {
          type: "string",
          description: "Link type: 'view' (read-only, default), 'edit' (read-write), or 'embed' (embeddable)",
          enum: ["view", "edit", "embed"]
        },
        scope: {
          type: "string",
          description: "Link scope: 'organization' (default, only org members) or 'anonymous' (anyone with the link)",
          enum: ["anonymous", "organization"]
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleShareItem
  },
  // ===== Search (Microsoft Search API) =====
  {
    name: "sharepoint-search",
    description: "Semantic search across SharePoint files using Microsoft Search API. Supports KQL (Keyword Query Language) for advanced queries. Much more powerful than basic file listing — finds content inside documents, matches by relevance, and returns highlighted summaries.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query. Supports KQL syntax: 'budget report', 'author:John', 'filetype:pdf', 'title:quarterly', 'created>2024-01-01', etc."
        },
        entityTypes: {
          type: "array",
          items: { type: "string" },
          description: "Types of content to search. Options: 'driveItem' (files, default), 'listItem' (list entries), 'site' (sites), 'list' (lists). Can combine multiple."
        },
        siteId: {
          type: "string",
          description: "Optional: scope search to a specific site by its ID"
        },
        count: {
          type: "number",
          description: "Number of results (default: 25, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearch
  },
  {
    name: "sharepoint-search-list-items",
    description: "Search specifically for SharePoint list items across all lists. Useful for finding records, tasks, issues, or any structured data stored in SharePoint lists.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query to find list items by content or field values"
        },
        siteId: {
          type: "string",
          description: "Optional: scope search to a specific site"
        },
        count: {
          type: "number",
          description: "Number of results (default: 25, max: 50)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchListItems
  },
  {
    name: "sharepoint-search-all",
    description: "Search across ALL SharePoint content types at once — files, list items, and sites — in a single query. Returns results grouped by type. Best for broad discovery queries like 'what do we have about X'.",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query. Supports KQL syntax for advanced filtering."
        },
        count: {
          type: "number",
          description: "Number of results per category (default: 10)"
        }
      },
      required: ["query"]
    },
    handler: handleSearchAll
  },

  {
    name: "sharepoint-list-permissions",
    description: "List all permissions on a file or folder in a SharePoint document library",
    inputSchema: {
      type: "object",
      properties: {
        driveId: {
          type: "string",
          description: "The ID of the drive/document library"
        },
        itemId: {
          type: "string",
          description: "The ID of the file or folder"
        }
      },
      required: ["driveId", "itemId"]
    },
    handler: handleListPermissions
  }
];

module.exports = {
  sharepointTools,
  handleListSites,
  handleGetSite,
  handleSearchSites,
  handleListDrives,
  handleGetDrive,
  handleListItems,
  handleGetItem,
  handleDownload,
  handleUpload,
  handleCreateFolder,
  handleDeleteItem,
  handleMoveItem,
  handleCopyItem,
  handleSearchFiles,
  handleListLists,
  handleGetList,
  handleListListItems,
  handleGetListItem,
  handleCreateListItem,
  handleUpdateListItem,
  handleDeleteListItem,
  handleListPages,
  handleGetPage,
  handleShareItem,
  handleListPermissions,
  handleSearch,
  handleSearchListItems,
  handleSearchAll
};
