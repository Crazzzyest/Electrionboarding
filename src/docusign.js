// Adapter: DocuSign. Both an outbound adapter (send envelope from template) and an inbound
// webhook source (Connect HMAC verification + event parsing). Uses the official docusign-esign
// SDK for JWT Grant auth and envelope creation — the one deliberate SDK exception in this
// codebase, since hand-rolling JWT RS256 signing isn't worth it. Never touches storage —
// onboarding.js owns all Status_*/timestamp writes.
const fs = require('fs');
const crypto = require('crypto');
const docusign = require('docusign-esign');
const config = require('./config');
const { formatDateNo } = require('./utils');

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

    // Anchor-based text tabs: DocuSign resolves each field's position from the literal [[...]]
    // string in the document at send time, so re-uploading/editing the template never breaks this
    // again (unlike stored auto-recognised tabs). anchorIgnoreIfNotPresent avoids a hard error if a
    // placeholder is ever removed. locked so the signer can't edit the prefilled value. The [[...]]
    // text is white in the template, so the value renders cleanly on top of it.
    const A = config.docusign.anchors;
    // Font matched to the contract's body (serif, 11pt, black) so the filled value reads as part of
    // the document rather than a pasted-in box. anchorYOffset nudges the value onto the placeholder's
    // baseline. The placeholder text under it is white, so only the value shows.
    const anchorTab = (anchorString, value) => ({
      anchorString,
      anchorUnits: 'pixels',
      anchorXOffset: '0',
      anchorYOffset: '-1',
      anchorIgnoreIfNotPresent: 'true',
      locked: 'true',
      font: 'timesnewroman',
      fontSize: 'size11',
      fontColor: 'black',
      value: value || '',
    });

    const envelopeDefinition = {
      templateId: config.docusign.templateId,
      templateRoles: [{
        email: candidate.privatEpost,
        name: `${candidate.fornavn} ${candidate.etternavn}`,
        roleName: config.docusign.signerRoleName,
        tabs: {
          // E-post is deliberately the @electi.no work address (microsoftUpn), NOT the private
          // delivery address the request was sent to. Stillingsprosent includes the "%" in the value
          // (the doc no longer prints a literal "%", which used to leave a big gap). Tiltredelsesdato
          // is shown in Norwegian dd.mm.yyyy; startdato is optional at registration (may be blank).
          textTabs: [
            anchorTab(A.navn, `${candidate.fornavn} ${candidate.etternavn}`),
            anchorTab(A.epost, candidate.microsoftUpn),
            anchorTab(A.telefon, candidate.mobil),
            anchorTab(A.stilling, candidate.stilling),
            anchorTab(A.stillingsprosent, candidate.stillingsprosent ? `${candidate.stillingsprosent} %` : ''),
            anchorTab(A.tiltredelsesdato, formatDateNo(candidate.startdato)),
            anchorTab(A.naermesteLeder, candidate.naermesteLeder),
          ],
        },
      }],
      customFields: {
        textCustomFields: [{ name: 'kandidatId', value: ctx.kandidatId }],
      },
      // Draft first: the commission-rate PDF is attached by the app (not stored in the template),
      // so the template is never touched — and each signed envelope locks in that day's rate sheet.
      status: 'created',
    };

    const result = await envelopesApi.createEnvelope(auth.accountId, { envelopeDefinition });
    const envelopeId = result.envelopeId;

    const apiBase = `${auth.accountBasePath}/v2.1/accounts/${auth.accountId}/envelopes/${envelopeId}`;
    const authHeader = { Authorization: `Bearer ${auth.token}` };

    // Attach the commission-rate PDF as document 2 (if configured and present on disk). Read as
    // raw bytes; a missing file is not fatal — the contract still goes out without the attachment.
    const { attachmentPath, attachmentName } = config.docusign;
    if (attachmentPath && attachmentName && fs.existsSync(attachmentPath)) {
      const res = await fetch(`${apiBase}/documents/2`, {
        method: 'PUT',
        headers: {
          ...authHeader,
          'Content-Type': 'application/pdf',
          'Content-Disposition': `file; filename="${attachmentName}"; documentid=2`,
        },
        body: fs.readFileSync(attachmentPath),
      });
      if (!res.ok) throw new Error(`DocuSign legg-ved-dokument feilet: ${res.status} ${await res.text()}`);
    } else if (attachmentPath) {
      console.warn(`DocuSign-vedlegg ikke funnet på disk (${attachmentPath}) — kontrakten sendes uten det.`);
    }

    // Now send the (draft) envelope.
    const send = await fetch(apiBase, {
      method: 'PUT',
      headers: { ...authHeader, 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: 'sent' }),
    });
    if (!send.ok) throw new Error(`DocuSign send-konvolutt feilet: ${send.status} ${await send.text()}`);

    return { ok: true, externalId: envelopeId };
  } catch (e) {
    const raw = e.message || String(e);
    const hint = raw.includes('consent_required')
      ? ' — engangssamtykke mangler, se docs/SETUP-CHECKLIST.md (besøk samtykke-URLen én gang)'
      : '';
    return { ok: false, error: raw + hint, retryable: true };
  }
}

// Downloads the fully signed, combined PDF (contract + attachment + DocuSign completion seal) for a
// completed envelope. Used by the webhook to archive the signed contract to SharePoint.
async function downloadCompletedPdf(envelopeId) {
  const auth = await authenticate();
  const res = await fetch(
    `${auth.accountBasePath}/v2.1/accounts/${auth.accountId}/envelopes/${envelopeId}/documents/combined`,
    { headers: { Authorization: `Bearer ${auth.token}` } },
  );
  if (!res.ok) throw new Error(`DocuSign hent signert PDF-feil: ${res.status} ${await res.text()}`);
  return Buffer.from(await res.arrayBuffer());
}

// Voids an envelope (e.g. when a candidate's data was corrected and a fresh contract is sent). A
// completed (already signed) envelope can't be voided — DocuSign returns an error, which the caller
// surfaces. No-op in demo mode.
async function voidEnvelope(envelopeId, reason = 'Erstattet av korrigert kontrakt') {
  if (config.demoMode || !envelopeId) return;
  const auth = await authenticate();
  const res = await fetch(
    `${auth.accountBasePath}/v2.1/accounts/${auth.accountId}/envelopes/${envelopeId}`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${auth.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: 'voided', voidedReason: reason }),
    },
  );
  if (!res.ok) throw new Error(`DocuSign void-feil: ${res.status} ${await res.text()}`);
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

module.exports = { ensure, verifyConnectSignature, parseWebhookEvent, downloadCompletedPdf, voidEnvelope };
