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

function parseRetryAfterMs(res, attempt) {
  const h = res.headers.get('retry-after');
  const secs = h != null ? Number(h) : NaN;
  if (Number.isFinite(secs)) return Math.min(secs * 1000, 30000);
  return Math.min(1000 * 2 ** (attempt - 1), 8000); // exponential fallback, capped
}

// Retries on 429 (throttling, incl. Excel's EditModeCannotAcquireLockTooManyRequests) and 503,
// honouring the Retry-After header when present. Everything else — including 4xx — is returned to
// the caller as-is. extraHeaders lets workbook calls attach a workbook-session-id (see excel-storage).
async function graphRequest(method, path, body, extraHeaders) {
  const maxAttempts = 5;
  for (let attempt = 1; ; attempt += 1) {
    const token = await getAccessToken();
    const res = await fetch(`${GRAPH_BASE}${path}`, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(extraHeaders || {}),
      },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
    if ((res.status === 429 || res.status === 503) && attempt < maxAttempts) {
      await new Promise((r) => setTimeout(r, parseRetryAfterMs(res, attempt)));
      continue;
    }
    return res;
  }
}

// Throws with the response body on non-2xx, otherwise returns parsed JSON, or null when there is
// no body. Several Graph writes return an empty body: 204 No Content (e.g. PATCH /users), but also
// 202 Accepted (e.g. sendMail) — so keying only on 204 made res.json() throw "Unexpected end of
// JSON input" on a perfectly successful send. Read the text once and only parse if non-empty.
async function graphJson(method, path, body, extraHeaders) {
  const res = await graphRequest(method, path, body, extraHeaders);
  if (!res.ok) throw new Error(`Graph ${method} ${path} feilet: ${res.status} ${await res.text()}`);
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

module.exports = { GRAPH_BASE, getAccessToken, graphRequest, graphJson };
