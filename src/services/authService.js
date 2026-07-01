import axios from 'axios';

/**
 * Retrieve an access token from Microsoft Identity Platform.
 * @param {string} tenantId - tenant ID
 * @param {string} clientId - application (client) ID
 * @param {string} clientSecret - application secret
 * @returns {Promise<string>} - A promise that resolves to the access token.
 */
export async function getAccessToken(tenantId, clientId, clientSecret) {
  try {
    tenantId = tenantId || process.env.TENANT_ID;
    clientId = clientId || process.env.CLIENT_ID;
    clientSecret = clientSecret || process.env.CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Tenant ID, Client ID, and Client Secret must be provided.");
    }
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    const response = await axios.post(url, params, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });
    return response.data.access_token;
  } catch (error) {
    console.error("Error obtaining access token:", error);
    throw error;
  }
}

/**
 * Retrieve an access token scoped to a SharePoint resource (tenant host).
 *
 * The Microsoft Graph API cannot create classic SharePoint list *views*; that
 * still requires the SharePoint REST API (`_api/web/lists/.../views`), which
 * rejects Graph-scoped tokens. This helper acquires an app-only token for the
 * `https://{sharepointHost}/.default` resource so view provisioning works.
 *
 * @param {string} tenantId - tenant ID
 * @param {string} clientId - application (client) ID
 * @param {string} clientSecret - application secret
 * @param {string} sharepointHost - the SharePoint host, e.g. "contoso.sharepoint.com"
 * @returns {Promise<string>} - A promise that resolves to the access token.
 */
export async function getSharePointToken(tenantId, clientId, clientSecret, sharepointHost) {
  try {
    tenantId = tenantId || process.env.TENANT_ID;
    clientId = clientId || process.env.CLIENT_ID;
    clientSecret = clientSecret || process.env.CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Tenant ID, Client ID, and Client Secret must be provided.");
    }
    if (!sharepointHost || !sharepointHost.trim()) {
      throw new Error("A SharePoint host (e.g. 'contoso.sharepoint.com') must be provided.");
    }
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('grant_type', 'client_credentials');
    params.append('client_id', clientId);
    params.append('client_secret', clientSecret);
    params.append('scope', `https://${sharepointHost}/.default`);
    const response = await axios.post(url, params, {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
    });
    return response.data.access_token;
  } catch (error) {
    console.error("Error obtaining SharePoint access token:", error);
    throw error;
  }
}