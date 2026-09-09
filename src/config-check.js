// Startup configuration checks + the data behind /health, so a deploy with a missing secret is
// caught immediately (loud warning at boot + visible in /health) rather than mid-onboarding. These
// only check that config is PRESENT — they don't make live API calls, to keep /health fast/cheap.
const fs = require('fs');
const config = require('./config');

function checks() {
  return {
    microsoft: Boolean(config.microsoft.tenantId && config.microsoft.clientId && config.microsoft.clientSecret),
    docusign: Boolean(
      config.docusign.integrationKey && config.docusign.userId && config.docusign.accountId
      && (config.docusign.privateKeyBase64 || fs.existsSync(config.docusign.privateKeyPath)),
    ),
    salesscreen: Boolean(config.salesscreen.apiKey),
    excel: Boolean(config.excel.siteId && config.excel.itemPath),
    epostAvsender: Boolean(config.email.sendAsMailbox),
    // In production the HMAC key must be set; in demo/other it's not required.
    docusignHmac: config.docusign.env !== 'production' || Boolean(config.docusign.connectHmacKey),
    tilgangskontroll: Boolean(config.auth.password),
  };
}

function logStartupChecks() {
  if (config.demoMode) {
    console.log('Demo-modus: hopper over konfig-sjekk (ingen eksterne kall).');
    return;
  }
  const c = checks();
  const missing = Object.entries(c).filter(([, ok]) => !ok).map(([k]) => k);
  if (!missing.length) {
    console.log('Konfig-sjekk: alle integrasjoner ser konfigurert ut.');
  } else {
    for (const k of missing) console.warn(`KONFIG-ADVARSEL: "${k}" mangler eller er ufullstendig.`);
  }
  if (!config.auth.password) {
    console.warn('SIKKERHETS-ADVARSEL: APP_PASSWORD er ikke satt — appen er AAPEN uten innlogging.');
  }
}

module.exports = { checks, logStartupChecks };
