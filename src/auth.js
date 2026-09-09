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

// Skips the gate for paths that must stay public (webhook, health).
function gate(req, res, next) {
  if (req.path === '/health' || req.path.startsWith('/webhooks/')) return next();
  return basicAuth(req, res, next);
}

module.exports = { gate, basicAuth };
