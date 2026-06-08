import { getAccessToken } from "./authService.js";
import axios from 'axios';

/**
* Retrieve a list of folders from the specified path in SharePoint.
* @param {string} authMode - Authentication mode ("application" or "delegated").
* @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
* @param {string} tenantId - tenant ID (required if authMode is "application")
* @param {string} clientId - application (client) ID (required if authMode is "application")
* @param {string} clientSecret - application secret (required if authMode is "application")
* @param {string} siteId - SharePoint site ID
* @param {string} driveId - SharePoint drive ID
* @param {string} path - The path in SharePoint to retrieve folders from.
* @returns {Promise<Array>} - A promise that resolves to an array of folder objects.
*/
export async function getFolders(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, path) {
  console.error("--- DEBUG getFolders ---");
  console.error("authMode", authMode);
  console.error("accessToken", accessToken ? "****" : null);
  console.error("tenantId", tenantId);
  console.error("clientId", clientId);
  console.error("clientSecret", clientSecret ? "****" : null);
  console.error("siteId", siteId);
  console.error("driveId", driveId);
  console.error("path", path);
  console.error("------------------------");

  try {
    let tokenToUse = "";
    let url = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret); // Ottieni il token di accesso
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    if (!path  || path === "/" || path.toLowerCase() === "root") {
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children?$filter=folder ne null`;
    } else {
      // Rimuovi eventuali slash iniziali o finali dal percorso
      const cleanPath = path.replace(/^\/|\/$/g, '');
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/children?$filter=folder ne null`;
    }
    const response = await axios.get(url, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
        "Content-type": 'application/json',
      },
    });
    const folders = response.data.value || [];
    return folders;
  } catch (error) {
    console.error("Errore durante il recupero delle cartelle:", error);
    throw error;
  }
}

/**
 * Create a new folder in SharePoint at the specified path.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} path - The parent path where the folder will be created.
 * @param {string} folderName - The name of the new folder to create.
 * @returns {Promise<Object>} - A promise that resolves to the created folder object.
 */
export async function createFolder(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, path, folderName) {
  try {
    let tokenToUse = "";
    let url = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret); // Ottieni il token di accesso
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    if (!path || path === "/" || path.toLowerCase() === "root") {
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`;
    } else {
      const cleanPath = path.replace(/^\/|\/$/g, '');
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/children`;
    }

    const response = await axios.post(
      url,
      {
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail"
      },
      {
        headers: {
          "Authorization": `Bearer ${tokenToUse}`,
          "Content-Type": "application/json",
        },
      }
    );
    
    return response.data;
  } catch (error) {
    console.error("Errore durante la creazione della cartella:", error);
    throw error;
  }
}

/**
 * Delete an empty folder in SharePoint at the specified path.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} path - The path of the folder to delete.
 * @returns {Promise<Object>} - A promise that resolves when the folder is deleted.
 */
export async function deleteFolder(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, path) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }

    if (!path || path === "/" || path.toLowerCase() === "root") {
      throw new Error("Cannot delete root folder");
    }
    
    const cleanPath = path.replace(/^\/|\/$/g, '');
    
    // Verifica che la cartella sia vuota
    const checkUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/children`;
    const checkResponse = await axios.get(checkUrl, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
      },
    });
    
    if (checkResponse.data.value && checkResponse.data.value.length > 0) {
      throw new Error("Folder is not empty. Cannot delete.");
    }
    
    // Elimina la cartella
    const deleteUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}`;
    await axios.delete(deleteUrl, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
      },
    });
    
    return { success: true, message: "Folder deleted successfully" };
  } catch (error) {
    console.error("Errore durante l'eliminazione della cartella:", error);
    throw error;
  }
}

/**
 * Get a tree view of the folder structure in SharePoint.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} path - The starting path (default: "root").
 * @param {number} maxDepth - Maximum depth to traverse (default: 3).
 * @returns {Promise<Object>} - A promise that resolves to a tree structure.
 */
export async function getFolderTree(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, path = "root", maxDepth = 3) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }

    async function buildTree(currentPath, depth = 0) {
      if (depth >= maxDepth) {
        return null;
      }
      
      let url = "";
      const normalizedPath = currentPath ? currentPath.replace(/^\/|\/$/g, '') : "";
      if (!normalizedPath || normalizedPath === "/" || normalizedPath.toLowerCase() === "root") {
        url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children?$filter=folder ne null`;
      } else {
        url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${normalizedPath}:/children?$filter=folder ne null`;
      }
      
      const response = await axios.get(url, {
        headers: {
          "Authorization": `Bearer ${tokenToUse}`,
          "Content-Type": "application/json",
        },
      });
      
      const folders = response.data.value || [];
      const tree = [];
      
      for (const folder of folders) {
        const folderPath = normalizedPath.toLowerCase() === "root" || !normalizedPath 
          ? folder.name 
          : `${normalizedPath}/${folder.name}`;
        
        const children = await buildTree(folderPath, depth + 1);
        
        tree.push({
          name: folder.name,
          path: folderPath,
          id: folder.id,
          childCount: folder.folder?.childCount || 0,
          children: children || []
        });
      }
      
      return tree;
    }
    
    const tree = await buildTree(path);
    return tree;
  } catch (error) {
    console.error("Errore durante il recupero dell'albero delle cartelle:", error);
    throw error;
  }
}