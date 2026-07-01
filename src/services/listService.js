import { getAccessToken, getSharePointToken } from "./authService.js";
import axios from 'axios';

const GRAPH = "https://graph.microsoft.com/v1.0";

/**
 * Resolve the bearer token to use for Microsoft Graph calls, honouring the
 * same application/delegated contract as the rest of the server.
 * @param {string} authMode - "application" or "delegated"
 * @param {string} accessToken - delegated token (required if authMode is "delegated")
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @returns {Promise<string>} - The bearer token.
 */
async function resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret) {
  if (authMode === "delegated") {
    return accessToken;
  }
  return getAccessToken(tenantId, clientId, clientSecret);
}

function authHeaders(token, extra = {}) {
  return {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json",
    ...extra,
  };
}

/**
 * List all SharePoint lists in a site.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @returns {Promise<Array>} - A promise that resolves to an array of list objects.
 */
export async function getLists(authMode, accessToken, tenantId, clientId, clientSecret, siteId) {
  try {
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists?$select=id,name,displayName,description,list,webUrl`;
    const response = await axios.get(url, { headers: authHeaders(token) });
    return response.data.value || [];
  } catch (error) {
    console.error("Errore durante il recupero delle liste:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Create a new SharePoint list. Columns may optionally be supplied inline as an
 * array of Microsoft Graph columnDefinition objects; complex columns
 * (calculated, lookup) are generally more reliable when added afterwards with
 * addColumn.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} displayName - The display name of the list.
 * @param {string} [description] - Optional list description.
 * @param {Array<Object>} [columns] - Optional array of Graph columnDefinition objects.
 * @param {string} [template] - List template (default "genericList").
 * @returns {Promise<Object>} - The created list object.
 */
export async function createList(authMode, accessToken, tenantId, clientId, clientSecret, siteId, displayName, description, columns, template) {
  try {
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists`;
    const body = {
      displayName,
      list: { template: template || "genericList" },
    };
    if (description) body.description = description;
    if (Array.isArray(columns) && columns.length > 0) body.columns = columns;
    const response = await axios.post(url, body, { headers: authHeaders(token) });
    return response.data;
  } catch (error) {
    console.error("Errore durante la creazione della lista:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Retrieve the columns of a list.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID (or list name)
 * @returns {Promise<Array>} - A promise that resolves to an array of column objects.
 */
export async function getListColumns(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId) {
  try {
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/columns`;
    const response = await axios.get(url, { headers: authHeaders(token) });
    return response.data.value || [];
  } catch (error) {
    console.error("Errore durante il recupero delle colonne:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Add a column to an existing list. The columnDefinition is a Microsoft Graph
 * columnDefinition object, so any supported column type works, e.g.:
 *   text:        { name, text: {} }
 *   number:      { name, number: {} }
 *   boolean:     { name, boolean: {} }
 *   dateTime:    { name, dateTime: { format: "dateOnly" } }
 *   currency:    { name, currency: { locale: "en-us" } }
 *   choice:      { name, choice: { choices: [...], displayAs: "dropDownMenu" } }
 *   person:      { name, personOrGroup: { allowMultipleSelection: false } }
 *   lookup:      { name, lookup: { listId, columnName } }
 *   calculated:  { name, calculated: { formula: "=[A]*[B]", outputType: "number" } }
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {Object} columnDefinition - Microsoft Graph columnDefinition object.
 * @returns {Promise<Object>} - The created column object.
 */
export async function addColumn(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, columnDefinition) {
  try {
    if (!columnDefinition || typeof columnDefinition !== "object") {
      throw new Error("columnDefinition must be a Microsoft Graph columnDefinition object");
    }
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/columns`;
    const response = await axios.post(url, columnDefinition, { headers: authHeaders(token) });
    return response.data;
  } catch (error) {
    console.error("Errore durante la creazione della colonna:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Retrieve items from a list, with their field values expanded.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {string} [filter] - Optional OData $filter applied to fields (e.g. "fields/Reviewed eq false").
 * @param {number} [top] - Optional page size ($top).
 * @param {string} [orderby] - Optional OData $orderby (e.g. "fields/Published desc").
 * @returns {Promise<Array>} - A promise that resolves to an array of list items (with .fields).
 */
export async function getListItems(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, filter, top, orderby) {
  try {
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const query = ["$expand=fields"];
    if (filter) query.push(`$filter=${encodeURIComponent(filter)}`);
    if (top) query.push(`$top=${top}`);
    if (orderby) query.push(`$orderby=${encodeURIComponent(orderby)}`);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/items?${query.join("&")}`;
    // $filter/$orderby on list-item fields require the ConsistencyLevel header.
    const response = await axios.get(url, { headers: authHeaders(token, { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" }) });
    return response.data.value || [];
  } catch (error) {
    console.error("Errore durante il recupero degli elementi della lista:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Create an item in a list.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {Object} fields - A map of internal column name -> value.
 * @returns {Promise<Object>} - The created item object.
 */
export async function addListItem(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, fields) {
  try {
    if (!fields || typeof fields !== "object") {
      throw new Error("fields must be an object mapping column names to values");
    }
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/items`;
    const response = await axios.post(url, { fields }, { headers: authHeaders(token) });
    return response.data;
  } catch (error) {
    console.error("Errore durante la creazione dell'elemento:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Update the field values of an existing list item.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {string} itemId - The list item ID.
 * @param {Object} fields - A map of internal column name -> new value.
 * @returns {Promise<Object>} - The updated fields object.
 */
export async function updateListItem(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, itemId, fields) {
  try {
    if (!fields || typeof fields !== "object") {
      throw new Error("fields must be an object mapping column names to values");
    }
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}/fields`;
    const response = await axios.patch(url, fields, { headers: authHeaders(token) });
    return response.data;
  } catch (error) {
    console.error("Errore durante l'aggiornamento dell'elemento:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Delete a list item.
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - User access token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {string} itemId - The list item ID.
 * @returns {Promise<Object>} - A confirmation object.
 */
export async function deleteListItem(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, itemId) {
  try {
    const token = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const url = `${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}`;
    await axios.delete(url, { headers: authHeaders(token) });
    return { deleted: true, itemId };
  } catch (error) {
    console.error("Errore durante l'eliminazione dell'elemento:", error?.response?.data || error.message);
    throw error;
  }
}

/**
 * Create a saved view on a list.
 *
 * Microsoft Graph does not expose list-view creation, so this uses the
 * SharePoint REST API. In application mode a SharePoint-scoped app-only token is
 * acquired automatically; in delegated mode the supplied accessToken must be
 * scoped to SharePoint (not Graph), otherwise the REST call is rejected.
 *
 * @param {string} authMode - Authentication mode ("application" or "delegated").
 * @param {string} accessToken - SharePoint-scoped token for delegated authentication.
 * @param {string} tenantId - tenant ID (required if authMode is "application")
 * @param {string} clientId - application (client) ID (required if authMode is "application")
 * @param {string} clientSecret - application secret (required if authMode is "application")
 * @param {string} siteId - SharePoint site ID (Graph composite: "host,siteCollGuid,webGuid")
 * @param {string} listId - SharePoint list ID (GUID)
 * @param {string} viewTitle - The title of the new view.
 * @param {Array<string>} viewFields - Ordered internal names of the fields to show.
 * @param {string} [viewQuery] - Optional CAML <Query> body (filters/sorts) without the wrapping <Query> tag.
 * @param {number} [rowLimit] - Optional row limit (default 30).
 * @returns {Promise<Object>} - The created view metadata.
 */
export async function createView(authMode, accessToken, tenantId, clientId, clientSecret, siteId, listId, viewTitle, viewFields, viewQuery, rowLimit) {
  try {
    // Resolve the site's web URL and host via Graph (works with a Graph token).
    const graphToken = await resolveGraphToken(authMode, accessToken, tenantId, clientId, clientSecret);
    const siteResp = await axios.get(`${GRAPH}/sites/${siteId}?$select=webUrl`, { headers: authHeaders(graphToken) });
    const webUrl = siteResp.data.webUrl; // e.g. https://contoso.sharepoint.com/sites/CauseWayHQ
    const host = new URL(webUrl).host;

    // SharePoint REST needs a SharePoint-scoped token.
    const spToken = authMode === "application"
      ? await getSharePointToken(tenantId, clientId, clientSecret, host)
      : accessToken;

    const spHeaders = {
      "Authorization": `Bearer ${spToken}`,
      "Content-Type": "application/json;odata=verbose",
      "Accept": "application/json;odata=verbose",
    };

    // 1. Create the view (empty field set, we populate it next).
    const createResp = await axios.post(
      `${webUrl}/_api/web/lists(guid'${listId}')/views`,
      {
        __metadata: { type: "SP.View" },
        Title: viewTitle,
        PersonalView: false,
        RowLimit: rowLimit || 30,
        ViewQuery: viewQuery || "",
      },
      { headers: spHeaders }
    );

    // 2. Set the visible fields in order: clear defaults, then add each.
    if (Array.isArray(viewFields) && viewFields.length > 0) {
      const viewFieldsBase = `${webUrl}/_api/web/lists(guid'${listId}')/views/getbytitle('${encodeURIComponent(viewTitle)}')/viewfields`;
      await axios.post(`${viewFieldsBase}/removeallviewfields`, {}, { headers: spHeaders });
      for (const field of viewFields) {
        await axios.post(
          `${viewFieldsBase}/addviewfield('${encodeURIComponent(field)}')`,
          {},
          { headers: spHeaders }
        );
      }
    }

    return createResp.data?.d || createResp.data;
  } catch (error) {
    console.error("Errore durante la creazione della vista:", error?.response?.data || error.message);
    throw error;
  }
}
