import { getAccessToken } from "./authService.js";
import axios from 'axios';
import { PDFParse } from 'pdf-parse';
import * as XLSX from 'xlsx';

// TODO: metadata is not yet implemented

/**
 * List all documents and their metadata in a specified path in SharePoint.
* @param {string} authMode - Authentication mode ("application" or "delegated").
* @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
* @param {string} tenantId - tenant ID (required if authMode is "application")
* @param {string} clientId - application (client) ID (required if authMode is "application")
* @param {string} clientSecret - application secret (required if authMode is "application")
* @param {string} siteId - SharePoint site ID
* @param {string} driveId - SharePoint drive ID
 * @param {string} path - The path in SharePoint to retrieve documents from.
 * @returns {Promise<Array>} - A promise that resolves to an array of document objects with metadata.
 */
export async function getDocuments(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, path) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    let url = "";
    if (!path || path === "/" || path.toLowerCase() === "root") {
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`;
    } else {
      const cleanPath = path.replace(/^\/|\/$/g, '');
      url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/children`;
    }
    
    const response = await axios.get(url, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
        "Content-Type": "application/json",
      },
    });
    
    const allItems = response.data.value || [];
    const documents = allItems.filter(item => item.file !== undefined);
    
    return documents;
  } catch (error) {
    console.error("Errore durante il recupero dei documenti:", error);
    throw error;
  }
}

/**
 * Get the content of a document in SharePoint, supporting multiple formats (PDF, Word, Excel).
 * Converts Word and Excel files to PDF automatically and extracts text.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} filePath - The path to the file (e.g., "Cartella_1/file.docx").
 * @returns {Promise<Object>} - A promise that resolves to the file content as text and metadata.
 */
export async function getDocumentContent(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, filePath) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }

    // Limite token per risposta all'agente
    const MAX_TOKENS = 200000;
    const CHARS_PER_TOKEN = 4; // stima approssimativa
    const SAFETY_FACTOR = 0.5
    const ALLOWED_CHARS = Math.floor(MAX_TOKENS * CHARS_PER_TOKEN * SAFETY_FACTOR);

    const cleanPath = filePath.replace(/^\/|\/$/g, '');

    // URL per ottenere i metadati del file
    const metadataUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}?$select=name,file,size`;

    const metadataResponse = await axios.get(metadataUrl, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
      },
    });

    const fileMetadata = metadataResponse.data;
    const mimeType = fileMetadata.file?.mimeType || '';

    // Determina se convertire in PDF
    const shouldConvertToPdf =
      mimeType.includes('wordprocessingml') || // Word
      mimeType.includes('spreadsheetml') ||    // Excel
      mimeType.includes('presentationml');     // PowerPoint

    let text = '';
    let pages = 0;
    let isTruncated = false;
   
    // Gestione excel, con libreria dedicata
    if (mimeType.includes('spreadsheetml')) {
      // URL per scaricare il contenuto
      let downloadUrlXlsx = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/content`;
      const bufferXlsx = await axios.get(downloadUrlXlsx, {
        headers: {
          "Authorization": `Bearer ${tokenToUse}`,
        },
        responseType: 'arraybuffer'
      });
      const workbook = XLSX.read(bufferXlsx.data, { type: 'buffer' });
      let totalUsedChars = 0;

      for (const sheetName of workbook.SheetNames) {
        if (isTruncated) {
          break;
        }
        const sheet = workbook.Sheets[sheetName];
        const sheetRows = XLSX.utils.sheet_to_csv(sheet).split('\n');
        let sheetText = `--- Sheet: ${sheetName} ---\n`;
        for (const row of sheetRows) {
          const rowWithNewline = row + '\n';
          if (totalUsedChars + rowWithNewline.length > ALLOWED_CHARS) {
            isTruncated = true;
            sheetText += "\n[... ATTENZIONE: Righe rimanenti omesse per superamento limite token ...]\n";
            break;
          } else {
            sheetText += rowWithNewline;
            totalUsedChars += rowWithNewline.length;
          }
        }
        text += sheetText;
      }
      return {
        name: fileMetadata.name,
        mimeType: mimeType,
        originalMimeType: mimeType,
        converted: false,
        size: bufferXlsx.data.byteLength,
        text: text,
        truncated: isTruncated,
        pages: 1 // Consideriamo 1 pagina per i file Excel
      };
    }

    // Gestione Word, PowerPoint e altri formati, con conversione a PDF e estrazione testo
    let downloadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/content` +
    (shouldConvertToPdf ? '?format=pdf' : '');  

    const contentResponse = await axios.get(downloadUrl, {
      headers: {
        "Authorization": `Bearer ${tokenToUse}`,
      },
      responseType: 'arraybuffer'
    });

    const contentType = contentResponse.headers['content-type'] || '';
    const buffer = Buffer.from(contentResponse.data);

    // File convertiti a PDF al download
    if (contentType.includes('pdf')) {
      try {
        const parser = new PDFParse({ data: buffer });
        const pdfData = await parser.getText()
        text = pdfData.text;
        pages = pdfData.numpages;
      } catch (pdfError) {
        console.error("Errore nell'estrazione del testo dal PDF:", pdfError);
        text = "[Impossibile estrarre il testo dal PDF]";
      }
    } else { // Altri formati, senza conversione a PDF
      text = buffer.toString('utf-8');
      pages = 1; // Per i file non PDF, consideriamo 1 pagina
    }

    // Troncamento del testo se supera il limite consentito
    if (text.length > ALLOWED_CHARS) {
      text = text.substring(0, ALLOWED_CHARS) + "\n[... ATTENZIONE: Testo troncato per superamento limite token ...]";
      isTruncated = true;
    }
    
    return {
      name: fileMetadata.name,
      mimeType: shouldConvertToPdf ? 'application/pdf' : mimeType,
      originalMimeType: mimeType,
      converted: shouldConvertToPdf,
      size: buffer.byteLength,
      text: text,
      truncated: isTruncated,
      pages: pages
    };
  } catch (error) {
    console.error("Errore durante il recupero del contenuto del documento:", error);
    throw error;
  }
}

/**
 * Upload a document to SharePoint (text or binary).
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} filePath - e.g. "Cartella_1/file.txt"
 * @param {string|Buffer} content - Text or binary content
 * @param {string} contentType - MIME type of the content
 * @param {boolean} overwrite - Whether to overwrite existing files
 * @returns {Promise<Object>}
 */
export async function uploadDocument(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, filePath, content, contentType = 'application/octet-stream', overwrite = false) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    const cleanPath = filePath.replace(/^\/|\/$/g, '');
    
    const isBinary = contentType && !contentType.startsWith('text/');
    const buffer = Buffer.isBuffer(content)
      ? content
      : isBinary
        ? Buffer.from(content, 'base64')
        : Buffer.from(content, 'utf-8');
    
    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/content` +
    (overwrite ? '' : '?@microsoft.graph.conflictBehavior=fail');
    
    const response = await axios.put(uploadUrl, buffer, {
      headers: {
        Authorization: `Bearer ${tokenToUse}`,
        'Content-Type': contentType,
        'Content-Length': buffer.length
      },
      maxBodyLength: Infinity
    });
    
    const file = response.data;
    
    return {
      id: file.id,
      name: file.name,
      size: file.size,
      mimeType: file.file?.mimeType,
      webUrl: file.webUrl,
      path: file.parentReference?.path
    };
  } catch (error) {
    console.error("Errore durante l'upload del documento:", error);
    throw error;
  }
}

/**
 * Update the content of an existing document in SharePoint. Replaces the entire content.
 * Fails if the document does not already exist.
 *
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} filePath - e.g. "Cartella_1/file.txt"
 * @param {string|Buffer} content - New content
 * @param {string} contentType - MIME type
 * @returns {Promise<Object>}
 */
export async function updateDocumentContent(
  authMode,
  accessToken,
  tenantId,
  clientId,
  clientSecret,
  siteId,
  driveId,
  filePath,
  content,
  contentType = 'application/octet-stream'
) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    const cleanPath = filePath.replace(/^\/|\/$/g, '');

    // Determina se è binario dal contentType
    const isBinary = contentType && !contentType.startsWith('text/');
    const buffer = Buffer.isBuffer(content)
      ? content
      : isBinary
        ? Buffer.from(content, 'base64')
        : Buffer.from(content, 'utf-8');

    const updateUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}:/content`;

    const response = await axios.put(updateUrl, buffer, {
      headers: {
        Authorization: `Bearer ${tokenToUse}`,
        'Content-Type': contentType,
        'Content-Length': buffer.length
      },
      maxBodyLength: Infinity
    });

    const file = response.data;

    return {
      id: file.id,
      name: file.name,
      size: file.size,
      mimeType: file.file?.mimeType,
      webUrl: file.webUrl,
      updated: true
    };
  } catch (error) {
    if (error.response?.status === 404) {
      console.error("Documento non esistente:", error);
      throw new Error('Documento non esistente: impossibile aggiornare');
    }
    console.error("Errore durante l'update del documento:", error);
    throw error;
  }
}

/**
 * Delete a document from SharePoint by path.
 * Fails if the document does not exist.
 *
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} driveId - SharePoint drive ID
 * @param {string} filePath - e.g. "Cartella_1/file.txt"
 * @returns {Promise<{deleted: boolean, path: string}>}
 */
export async function deleteDocument(authMode, accessToken, tenantId, clientId, clientSecret, siteId, driveId, filePath) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }
    const cleanPath = filePath.replace(/^\/|\/$/g, '');   
    const deleteUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${cleanPath}`;    

    await axios.delete(deleteUrl, {
      headers: {
        Authorization: `Bearer ${tokenToUse}`,
      }
    });

    return {
      deleted: true,
      path: cleanPath
    };
  } catch (error) {
    if (error.response?.status === 404) {
      console.error("Documento non trovato:", error);
      throw new Error('Documento non trovato: impossibile eliminare');
    }
    if (error.response?.status === 403) {
      console.error("Permessi insufficienti per eliminare il documento:", error);
      throw new Error('Permessi insufficienti per eliminare il documento');
    }
    console.error("Errore durante l'eliminazione del documento:", error);
    throw error;
  }
}

/**
 * Search for documents in Sharepoint containing the specified keywords in the specified attribute.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication (required if authMode is "delegated").
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - The ID of the SharePoint list to search in.
 * @param {string[]} keywords - An array of keywords to search for.
 * @param {string} attributeName - The attribute name to search in (e.g., "name", "fileType").
 * @returns {Promise<Array>} - A promise that resolves to an array of matching document objects.
 */
export async function searchDocumentsByKeywords(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, keywords, attributeName) {
  try {
    let tokenToUse = "";
    if (authMode === "application") {
      tokenToUse = await getAccessToken(tenantId, clientId, clientSecret);
    } else if (authMode === "delegated") {
      tokenToUse = accessToken;
    }

    let allItems = [];
    let nextLink = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields,driveItem&$select=id,fields,driveItem`;

    // Gestione della paginazione con @odata.nextLink
    while (nextLink) {
      const response = await axios.get(nextLink, {
        headers: {
          "Authorization": `Bearer ${tokenToUse}`,
          "Content-Type": "application/json",
        },
      });

      const items = response.data.value || [];
      allItems = allItems.concat(items);
      
      // Controlla se c'è una pagina successiva
      nextLink = response.data['@odata.nextLink'] || null;
    }

    // Filtra solo i documenti (esclude le cartelle)
    const documents = allItems.filter(item => 
      item.driveItem && item.driveItem.file !== undefined
    );

    // Filtra i documenti che contengono le keywords nell'attributo specificato
    const matchingDocuments = documents.filter(doc => {
      let attributeValue = null;
      
      // Mappa alcuni nomi comuni di attributi
      if (attributeName === 'name') {
        // Il nome del file può essere in driveItem.name o fields.FileLeafRef
        attributeValue = doc.driveItem?.name || doc.fields?.FileLeafRef;
      } else if (attributeName === 'contentType') {
        attributeValue = doc.fields?.ContentType;
      } else {
        // Cerca prima in fields, poi in driveItem
        attributeValue = doc.fields?.[attributeName] || doc.driveItem?.[attributeName];
      }
      
      if (attributeValue) {
        return keywords.some(keyword => 
          attributeValue.toString().toLowerCase().includes(keyword.toLowerCase())
        );
      }
      return false;
    });

    return matchingDocuments;
  } catch (error) {
    console.error("Errore durante la ricerca dei documenti:", error);
    throw error;
  }
}