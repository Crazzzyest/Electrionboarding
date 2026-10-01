// Basic Auth gate for the UI and API. Enabled only when APP_PASSWORD is set (see config.auth).
// The DocuSign webhook and /health are excluded by the caller — the webhook authenticates itself
// via HMAC and must stay reachable by DocuSign; /health must stay reachable by monitoring.
const crypto = require('crypto');
const config = require('./config');

function safeEqual(a, b) {
  const ab = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  if (ab.length !== bb.length) return false;
  return crypto.timingSafeEqual(ab, bb);
}

function basicAuth(req, res, next) {
  const { username, password } = config.auth;
  if (!password) return next(); // not configured -> open (warned at startup)

  const header = req.headers.authorization || '';
  const m = /^Basic (.+)$/.exec(header);
  if (m) {
    const decoded = Buffer.from(m[1], 'base64').toString();
    const idx = decoded.indexOf(':');
    const u = decoded.slice(0, idx);
    const p = decoded.slice(idx + 1);
    if (safeEqual(u, username) && safeEqual(p, password)) return next();
  }
  res.set('WWW-Authenticate', 'Basic realm="Electi Onboarding"');
  return res.status(401).send('Autentisering kreves.');
}

// CSRF guard. The browser re-sends cached Basic Auth credentials on ANY request to this origin,
// including a form POST or <img> triggered by a malicious page elsewhere — and several endpoints
// (resend-contract, retry, test-birthday) act without needing a body. Requiring a custom header on
// every state-changing /api call blocks that: a cross-origin page cannot set a custom header
// without a CORS preflight, and this app answers no preflights. The UI adds the header to every
// fetch (see the fetch wrapper in public/index.html).
const CSRF_HEADER = 'x-requested-with';
const CSRF_VALUE = 'electi-onboarding';
const SAFE_METHODS = new Set(['GET', 'HEAD', 'OPTIONS']);

function csrfGuard(req, res, next) {
  if (SAFE_METHODS.has(req.method) || !req.path.startsWith('/api/')) return next();
  if (req.get(CSRF_HEADER) === CSRF_VALUE) return next();
  return res.status(403).json({ success: false, error: 'Mangler CSRF-header — forespørselen ble avvist.' });
}

// Skips the gate for paths that must stay public (webhook, health).
function gate(req, res, next) {
  if (req.path === '/health' || req.path.startsWith('/webhooks/')) return next();
  return basicAuth(req, res, () => csrfGuard(req, res, next));
}

module.exports = { gate, basicAuth, csrfGuard, CSRF_HEADER, CSRF_VALUE };
