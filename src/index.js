#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { getFolders, createFolder, deleteFolder, getFolderTree } from "./services/folderService.js";
import { getDocuments, getDocumentContent, uploadDocument, updateDocumentContent, deleteDocument, searchDocumentsByKeywords } from "./services/fileService.js";

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
  const authMode = args.authMode;
  if (!authMode) {
    throw new Error("authMode è obbligatorio");
  }
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
          required: ["authMode", "siteId", "driveId", "path"],
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
          required: ["authMode", "siteId", "driveId", "path", "folderName"],
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
          required: ["authMode", "siteId", "driveId", "path"],
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
          required: ["authMode", "siteId", "driveId"],
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
          required: ["authMode", "siteId", "driveId", "path"],
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
          required: ["authMode", "siteId", "driveId", "filePath"],
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
          required: ["authMode", "siteId", "driveId", "filePath", "content"],
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
          required: ["authMode", "siteId", "driveId", "filePath", "content"],
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
          required: ["authMode", "siteId", "driveId", "filePath"],
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
          required: ["authMode", "siteId", "listId", "keywords", "attributeName"],
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