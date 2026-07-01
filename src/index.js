#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { getFolders, createFolder, deleteFolder, getFolderTree } from "./services/folderService.js";
import { getDocuments, getDocumentContent, uploadDocument, updateDocumentContent, deleteDocument, searchDocumentsByKeywords } from "./services/fileService.js";
import { getLists, createList, getListColumns, addColumn, getListItems, addListItem, updateListItem, deleteListItem, createView } from "./services/listService.js";

// Crea il server
const server = new Server(
  {
    name: "mcp-sharepoint",
    version: "1.1.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Validatore per autenticazione
function validateAuth(args) {
  // Retro-compatibilita': le installazioni esistenti (pre-1.1.0) non inviano
  // authMode e si aspettano il comportamento storico app-only. Default
  // "application" cosi' la release puo' essere promossa a `latest` senza
  // rompere le istanze gia' configurate.
  const authMode = args.authMode || "application";
  args.authMode = authMode;
  if (authMode === "application") {
    if (!args.tenantId?.trim() || !args.clientId?.trim() || !args.clientSecret?.trim()) {
      throw new Error("Per l'autenticazione applicativa sono richiesti tenantId, clientId e clientSecret");
    }
  } else if (authMode === "delegated") {
    if (!args.accessToken?.trim()) {
      throw new Error("Per l'autenticazione delegata è richiesto l'accessToken");
    }
  } else {
    throw new Error("authMode deve essere 'application' o 'delegated'");
  }
}

// Lista dei tool disponibili
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "getFolders",
        description: "Retrieve a list of folders from the specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            path: { type: "string", description: "The path in SharePoint to retrieve folders from" },
          },
          required: ["siteId", "driveId", "path"],
        },
      },
      {
        name: "createFolder",
        description: "Create a new folder in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            path: { type: "string", description: "The parent path where the folder will be created" },
            folderName: { type: "string", description: "The name of the new folder to create" },
          },
          required: ["siteId", "driveId", "path", "folderName"],
        },
      },
      {
        name: "deleteFolder",
        description: "Delete an empty folder in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            path: { type: "string", description: "The path of the folder to delete" },
          },
          required: ["siteId", "driveId", "path"],
        },
      },
      {
        name: "getFolderTree",
        description: "Get a tree view of the folder structure in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            path: { type: "string", description: "The starting path (default: 'root')" },
            maxDepth: { type: "number", description: "Maximum depth to traverse (default: 3)" },
          },
          required: ["siteId", "driveId"],
        },
      },
      {
        name: "getDocuments",
        description: "List all documents and their metadata in a specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            path: { type: "string", description: "The path in SharePoint to retrieve documents from" },
          },
          required: ["siteId", "driveId", "path"],
        },
      },
      {
        name: "getDocumentContent",
        description: "Get the content of a document in SharePoint, supporting multiple formats (PDF, Word, Excel)",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            filePath: { type: "string", description: "The path to the file (e.g., 'Cartella_1/file.docx')" },
          },
          required: ["siteId", "driveId", "filePath"],
        },
      },
      {
        name: "uploadDocument",
        description: "Upload a document to a specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            filePath: { type: "string", description: "The path where the file will be uploaded (e.g., 'Cartella_1')" },
            content: { type: "string", description: "The string or base64-encoded content of the file to upload" },
            contentType: { type: "string", description: "The MIME type of the content (e.g., 'application/pdf')" },
            overwrite: { type: "boolean", description: "Whether to overwrite existing files" },
          },
          required: ["siteId", "driveId", "filePath", "content"],
        },
      },
      {
        name: "updateDocumentContent",
        description: "Update the content of an existing document in SharePoint, Replaces the entire content.",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            filePath: { type: "string", description: "The path to the existing file to update (e.g., 'Cartella_1/file.docx')" },
            content: { type: "string", description: "The new string or base64-encoded content of the file" },
            contentType: { type: "string", description: "The MIME type of the content (e.g., 'application/pdf')" },
          },
          required: ["siteId", "driveId", "filePath", "content"],
        },
      },
      {
        name: "deleteDocument",
        description: "Delete a document in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            driveId: {type: "string",description: "The ID of the drive within the SharePoint site",},
            filePath: { type: "string", description: "The path to the file to delete (e.g., 'Cartella_1/file.docx')" },
          },
          required: ["siteId", "driveId", "filePath"],
        },
      },
      {
        name: "searchDocumentsByKeywords",
        description: "Search for documents in SharePoint containing specific keywords in the given attribute",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type:"string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: {type: "string", description: "The directory (tenant) ID (required if authMode is 'application')",},
            clientId: {type: "string", description: "The application (client) ID (required if authMode is 'application')",},
            clientSecret: {type: "string", description: "The client secret (required if authMode is 'application')",},
            siteId: {type: "string",description: "The ID of the SharePoint site",},
            listId: { type: "string", description: "The ID of the SharePoint list to search in" },
            keywords: { type: "array", items: { type: "string" }, description: "Array of keywords to search for" },
            attributeName: { type: "string", description: "The document attribute to search in (e.g., 'name', 'content')" },
          },
          required: ["siteId", "listId", "keywords", "attributeName"],
        },
      },
      {
        name: "getLists",
        description: "List all SharePoint lists in a site",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
          },
          required: ["siteId"],
        },
      },
      {
        name: "createList",
        description: "Create a new SharePoint list, optionally with inline column definitions (Microsoft Graph columnDefinition objects)",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            displayName: { type: "string", description: "The display name of the new list" },
            description: { type: "string", description: "Optional description for the list" },
            columns: { type: "array", items: { type: "object" }, description: "Optional array of Microsoft Graph columnDefinition objects to create with the list" },
            template: { type: "string", description: "List template (default 'genericList')" },
          },
          required: ["siteId", "displayName"],
        },
      },
      {
        name: "getListColumns",
        description: "Retrieve the columns of a SharePoint list",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID (or name) of the SharePoint list" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "addColumn",
        description: "Add a column to an existing list. columnDefinition is a Microsoft Graph columnDefinition object supporting text, number, boolean, dateTime, currency, choice, personOrGroup, lookup and calculated column types",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID of the SharePoint list" },
            columnDefinition: { type: "object", description: "A Microsoft Graph columnDefinition object, e.g. { name: 'WeightedValue', calculated: { formula: '=[EstValue]*[Probability]', outputType: 'number' } }" },
          },
          required: ["siteId", "listId", "columnDefinition"],
        },
      },
      {
        name: "getListItems",
        description: "Retrieve items from a list with their field values expanded, with optional OData filter, orderby and top",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID of the SharePoint list" },
            filter: { type: "string", description: "Optional OData $filter on fields, e.g. \"fields/Reviewed eq false\"" },
            top: { type: "number", description: "Optional page size ($top)" },
            orderby: { type: "string", description: "Optional OData $orderby, e.g. \"fields/Published desc\"" },
          },
          required: ["siteId", "listId"],
        },
      },
      {
        name: "addListItem",
        description: "Create an item in a list. fields is a map of internal column name to value",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID of the SharePoint list" },
            fields: { type: "object", description: "A map of internal column name to value, e.g. { Title: 'CBY circular', Category: 'CBY-Aden Circular' }" },
          },
          required: ["siteId", "listId", "fields"],
        },
      },
      {
        name: "updateListItem",
        description: "Update the field values of an existing list item",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID of the SharePoint list" },
            itemId: { type: "string", description: "The ID of the list item to update" },
            fields: { type: "object", description: "A map of internal column name to new value" },
          },
          required: ["siteId", "listId", "itemId", "fields"],
        },
      },
      {
        name: "deleteListItem",
        description: "Delete an item from a list",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "The user access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The ID of the SharePoint list" },
            itemId: { type: "string", description: "The ID of the list item to delete" },
          },
          required: ["siteId", "listId", "itemId"],
        },
      },
      {
        name: "createView",
        description: "Create a saved view on a list (via SharePoint REST, since Graph cannot create views). In application mode a SharePoint-scoped token is acquired automatically; in delegated mode accessToken must be SharePoint-scoped",
        inputSchema: {
          type: "object",
          properties: {
            authMode: { type: "string", enum: ["application", "delegated"], description: "Authentication mode, application with secret or delegated with token" },
            accessToken: { type: "string", description: "A SharePoint-scoped access token for delegated authentication (required if authMode is 'delegated')" },
            tenantId: { type: "string", description: "The directory (tenant) ID (required if authMode is 'application')" },
            clientId: { type: "string", description: "The application (client) ID (required if authMode is 'application')" },
            clientSecret: { type: "string", description: "The client secret (required if authMode is 'application')" },
            siteId: { type: "string", description: "The ID of the SharePoint site" },
            listId: { type: "string", description: "The GUID of the SharePoint list" },
            viewTitle: { type: "string", description: "The title of the new view" },
            viewFields: { type: "array", items: { type: "string" }, description: "Ordered internal names of the fields to display" },
            viewQuery: { type: "string", description: "Optional CAML query body (filters/sorts) without the wrapping <Query> tag" },
            rowLimit: { type: "number", description: "Optional row limit (default 30)" },
          },
          required: ["siteId", "listId", "viewTitle"],
        },
      }
    ],
  };
});

// Gestione delle chiamate ai tool
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  validateAuth(args);

  switch (name) {
    case "getFolders":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getFolders(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.path)),
          },
        ],
      };

    case "createFolder":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await createFolder(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.path, args.folderName)),
          },
        ],
      };

    case "deleteFolder":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await deleteFolder(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.path)),
          },
        ],
      };

    case "getFolderTree":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getFolderTree(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.path, args.maxDepth)),
          },
        ],
      };

    case "getDocuments":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getDocuments(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.path)),
          },
        ],
      };

    case "getDocumentContent":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getDocumentContent(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.filePath)),
          },
        ],
      };

    case "uploadDocument":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await uploadDocument(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.filePath, args.content, args.contentType, args.overwrite)),
          },
        ],
      };

    case "updateDocumentContent":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await updateDocumentContent(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.filePath, args.content, args.contentType)),
          },
        ],
      };

    case "deleteDocument":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await deleteDocument(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.driveId, args.filePath)),
          },
        ],
      };

    case "searchDocumentsByKeywords":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await searchDocumentsByKeywords(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.keywords, args.attributeName)),
          },
        ],
      };

    case "getLists":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getLists(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId)),
          },
        ],
      };

    case "createList":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await createList(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.displayName, args.description, args.columns, args.template)),
          },
        ],
      };

    case "getListColumns":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getListColumns(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId)),
          },
        ],
      };

    case "addColumn":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await addColumn(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.columnDefinition)),
          },
        ],
      };

    case "getListItems":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getListItems(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.filter, args.top, args.orderby)),
          },
        ],
      };

    case "addListItem":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await addListItem(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.fields)),
          },
        ],
      };

    case "updateListItem":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await updateListItem(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.itemId, args.fields)),
          },
        ],
      };

    case "deleteListItem":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await deleteListItem(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.itemId)),
          },
        ],
      };

    case "createView":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await createView(args.authMode, args.accessToken, args.tenantId, args.clientId, args.clientSecret, args.siteId, args.listId, args.viewTitle, args.viewFields, args.viewQuery, args.rowLimit)),
          },
        ],
      };

    default:
      throw new Error(`Tool sconosciuto: ${name}`);
  }
});

// Avvia il server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Server MCP avviato");
}

main().catch((error) => {
  console.error("Errore:", error);
  process.exit(1);
});