// Adapter: DocuSign. Both an outbound adapter (send envelope from template) and an inbound
// webhook source (Connect HMAC verification + event parsing). Uses the official docusign-esign
// SDK for JWT Grant auth and envelope creation — the one deliberate SDK exception in this
// codebase, since hand-rolling JWT RS256 signing isn't worth it. Never touches storage —
// onboarding.js owns all Status_*/timestamp writes.
const fs = require('fs');
const crypto = require('crypto');
const docusign = require('docusign-esign');
const config = require('./config');

function loadPrivateKey() {
  if (config.docusign.privateKeyBase64) {
    return Buffer.from(config.docusign.privateKeyBase64, 'base64');
  }
  return fs.readFileSync(config.docusign.privateKeyPath);
}

function authServer() {
  return config.docusign.env === 'production' ? 'account.docusign.com' : 'account-d.docusign.com';
}

// Cached { token, expiry, accountBasePath, accountId }. The per-account API base path (e.g.
// na1/na2/na3/eu) is NOT a fixed URL to guess — it's discovered via getUserInfo, same as
// DocuSign's own quickstart samples do, so this works regardless of which region Electi's
// production account happens to live in.
let cached = null;

async function authenticate() {
  if (cached && Date.now() < cached.expiry - 60000) return cached;

  const apiClient = new docusign.ApiClient();
  apiClient.setOAuthBasePath(authServer());

  const results = await apiClient.requestJWTUserToken(
    config.docusign.integrationKey,
    config.docusign.userId,
    ['signature', 'impersonation'],
    loadPrivateKey(),
    3600,
  );
  const accessToken = results.body.access_token;

  const userInfo = await apiClient.getUserInfo(accessToken);
  const account = userInfo.accounts.find((a) => a.accountId === config.docusign.accountId)
    || userInfo.accounts.find((a) => a.isDefault)
    || userInfo.accounts[0];
  if (!account) throw new Error('Fant ingen DocuSign-konto for denne brukeren');

  cached = {
    token: accessToken,
    expiry: Date.now() + results.body.expires_in * 1000,
    accountBasePath: `${account.baseUri}/restapi`,
    accountId: account.accountId,
  };
  return cached;
}

async function ensure(candidate, ctx) {
  if (config.demoMode) {
    return { ok: true, externalId: `DEMO-ENVELOPE-${ctx.kandidatId}`, demoMode: true };
  }

  try {
    const auth = await authenticate();
    const apiClient = new docusign.ApiClient();
    apiClient.setBasePath(auth.accountBasePath);
    apiClient.addDefaultHeader('Authorization', `Bearer ${auth.token}`);
    const envelopesApi = new docusign.EnvelopesApi(apiClient);

    const envelopeDefinition = {
      templateId: config.docusign.templateId,
      templateRoles: [{
        email: candidate.privatEpost,
        name: `${candidate.fornavn} ${candidate.etternavn}`,
        roleName: config.docusign.signerRoleName,
        tabs: {
          // Navn auto-fills from `name` on the templateRole (the FullName field), so it needs no tab.
          // Everything else is a Text field the app fills here. E-post is deliberately the @electi.no
          // work address (microsoftUpn), NOT the private delivery address the request was sent to.
          // Stillingsprosent is the number only — the template prints the literal "%" after the field.
          // Tiltredelsesdato comes from startdato, which is optional at registration (may be blank).
          textTabs: [
            { tabLabel: config.docusign.tabLabels.epost, value: candidate.microsoftUpn },
            { tabLabel: config.docusign.tabLabels.telefon, value: candidate.mobil },
            { tabLabel: config.docusign.tabLabels.stilling, value: candidate.stilling },
            { tabLabel: config.docusign.tabLabels.stillingsprosent, value: String(candidate.stillingsprosent) },
            { tabLabel: config.docusign.tabLabels.tiltredelsesdato, value: candidate.startdato || '' },
            { tabLabel: config.docusign.tabLabels.naermesteLeder, value: candidate.naermesteLeder },
          ],
        },
      }],
      customFields: {
        textCustomFields: [{ name: 'kandidatId', value: ctx.kandidatId }],
      },
      status: 'sent',
    };

    const result = await envelopesApi.createEnvelope(auth.accountId, { envelopeDefinition });
    return { ok: true, externalId: result.envelopeId };
  } catch (e) {
    const raw = e.message || String(e);
    const hint = raw.includes('consent_required')
      ? ' — engangssamtykke mangler, se docs/SETUP-CHECKLIST.md (besøk samtykke-URLen én gang)'
      : '';
    return { ok: false, error: raw + hint, retryable: true };
  }
}

// DocuSign Connect signs the raw request body with HMAC-SHA256 using the configured key;
// compares against the X-DocuSign-Signature-1 header. Skipped entirely in DEMO_MODE (see the
// webhook route in index.js).
function verifyConnectSignature(rawBody, signatureHeader) {
  if (!config.docusign.connectHmacKey || !signatureHeader) return false;
  const computed = crypto
    .createHmac('sha256', config.docusign.connectHmacKey)
    .update(rawBody)
    .digest('base64');
  try {
    return crypto.timingSafeEqual(Buffer.from(computed), Buffer.from(signatureHeader));
  } catch {
    return false; // e.g. length mismatch -> definitely not equal
  }
}

// NOTE: exact JSON nesting depends on the Connect configuration screen (aggregate vs per-event
// payload format). Written defensively against the standard per-event JSON shape — verify
// against one real test envelope during setup (see docs/SETUP-CHECKLIST.md) rather than
// trusting this blind.
function parseWebhookEvent(body) {
  const envelopeId = body?.data?.envelopeId || body?.envelopeId;
  const rawStatus = body?.data?.envelopeSummary?.status || body?.event || body?.status || '';
  const status = String(rawStatus).toLowerCase().replace('envelope-', '');
  return { envelopeId, status };
}

module.exports = { ensure, verifyConnectSignature, parseWebhookEvent };
