// Shared Microsoft Graph app-only (client-credentials) client. Used by microsoft.js (user/
// license/group provisioning), excel-storage.js (candidate data + log), and graph-mail.js
// (Telenor order / welcome / birthday emails) — one token cache, one auth path, so a tenant
// admin only ever has to reason about a single set of Graph permissions for the whole app.
const config = require('./config');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

let cachedToken = null;
let cachedTokenExpiry = 0;

async function getAccessToken() {
  if (cachedToken && Date.now() < cachedTokenExpiry - 60000) return cachedToken;

  const url = `https://login.microsoftonline.com/${config.microsoft.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: config.microsoft.clientId,
    client_secret: config.microsoft.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if (!res.ok) throw new Error(`Microsoft-token-feil: ${res.status} ${await res.text()}`);
  const data = await res.json();
  cachedToken = data.access_token;
  cachedTokenExpiry = Date.now() + data.expires_in * 1000;
  return cachedToken;
}

async function graphRequest(method, path, body) {
  const token = await getAccessToken();
  return fetch(`${GRAPH_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
}

// Throws with the response body on non-2xx, otherwise returns parsed JSON (or null for 204s) —
// most callers just want the happy path without repeating the same ok-check everywhere.
async function graphJson(method, path, body) {
  const res = await graphRequest(method, path, body);
  if (!res.ok) throw new Error(`Graph ${method} ${path} feilet: ${res.status} ${await res.text()}`);
  if (res.status === 204) return null;
  return res.json();
}

module.exports = { GRAPH_BASE, getAccessToken, graphRequest, graphJson };
