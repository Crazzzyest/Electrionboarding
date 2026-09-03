const express = require('express');
const path = require('path');
const cron = require('node-cron');
const config = require('./config');
const storage = require('./storage');
const onboarding = require('./onboarding');
const docusign = require('./docusign');
const birthday = require('./birthday');
const { KONTRAKT_STATUS, LOGG_HANDLING, LOGG_KILDE } = require('./columns');

const app = express();

// ============================================================
// HEALTH
// ============================================================

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    testMode: config.testMode,
    demoMode: config.demoMode,
    timestamp: new Date().toISOString(),
  });
});

// ============================================================
// CANDIDATES
// ============================================================

app.get('/api/candidates', async (req, res) => {
  try {
    const candidates = await storage.listCandidates();
    res.json({ success: true, candidates });
  } catch (e) {
    console.error('list candidates error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.get('/api/candidates/:row', async (req, res) => {
  try {
    const candidate = await storage.getCandidate(req.params.row);
    if (!candidate) return res.status(404).json({ success: false, error: 'Ikke funnet' });
    const log = await storage.listLog(candidate.kandidatId);
    res.json({ success: true, candidate, log });
  } catch (e) {
    console.error('get candidate error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

const REQUIRED_FIELDS = [
  'fornavn', 'etternavn', 'fodselsdato', 'privatEpost', 'mobil',
  'stilling', 'stillingsprosent', 'avdeling', 'naermesteLeder', 'registrertAv',
];

app.post('/api/candidates', express.json(), async (req, res) => {
  try {
    const missing = REQUIRED_FIELDS.filter((f) => !req.body[f]);
    if (missing.length) {
      return res.status(400).json({ success: false, error: `Mangler felt: ${missing.join(', ')}` });
    }

    // Decide the @electi.no address up front — the employment contract states it, so it has to
    // exist before the envelope is sent, not after the Microsoft account is created.
    const microsoft = require('./microsoft');
    const candidate = await storage.createCandidate({
      ...req.body,
      microsoftUpn: microsoft.buildUpn(req.body),
    });
    await storage.appendLog(candidate.kandidatId, 'registrering', LOGG_HANDLING.FULLFORT, 'Kandidat registrert', LOGG_KILDE.REGISTRERING);

    // Registering *is* what sends the contract.
    onboarding.sendContract(candidate.row).catch((e) => console.error('sendContract error:', e));

    res.json({ success: true, candidate });
  } catch (e) {
    console.error('create candidate error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// Retry: either one named step (force bypasses the "contract must be signed" gate — the manual
// override escape hatch for edge cases like a contract signed on paper outside DocuSign), or
// every pending step at once.
app.post('/api/candidates/:row/retry', express.json(), async (req, res) => {
  try {
    const { step, force } = req.body || {};
    const { row } = req.params;

    if (step) {
      const result = await onboarding.runStep(row, step, { force: Boolean(force), trigger: LOGG_KILDE.MANUELL });
      return res.json({ success: true, result });
    }

    const result = await onboarding.runOnboardingSteps(row, { trigger: LOGG_KILDE.MANUELL });
    res.json({ success: true, result });
  } catch (e) {
    console.error('retry error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/candidates/:row/resend-contract', async (req, res) => {
  try {
    const result = await onboarding.sendContract(req.params.row);
    res.json({ success: true, result });
  } catch (e) {
    console.error('resend-contract error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// DOCUSIGN WEBHOOK
// ============================================================

// Shared by the real Connect webhook and the DEMO_MODE "simulate signed" endpoint below, so the
// demo path exercises the exact same logic as production rather than a parallel implementation.
async function handleEnvelopeEvent(event) {
  const candidate = await storage.findCandidateByEnvelopeId(event.envelopeId);
  if (!candidate) {
    console.error(`DocuSign webhook: ingen kandidat funnet for envelope ${event.envelopeId}`);
    return { matched: false };
  }

  if (event.status === 'completed') {
    await storage.updateCandidateFields(candidate.row, {
      statusKontrakt: KONTRAKT_STATUS.SIGNERT,
      kontraktSignertDato: new Date().toISOString(),
    });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FULLFORT, 'Signert', LOGG_KILDE.WEBHOOK);
    onboarding.runOnboardingSteps(candidate.row, { trigger: LOGG_KILDE.WEBHOOK })
      .catch((e) => console.error('runOnboardingSteps error:', e));
  } else if (event.status === 'declined' || event.status === 'voided') {
    await storage.updateCandidateFields(candidate.row, { statusKontrakt: KONTRAKT_STATUS.AVSLATT });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FEILET, `Status: ${event.status}`, LOGG_KILDE.WEBHOOK);
  }

  return { matched: true };
}

app.post('/webhooks/docusign',
  express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }),
  async (req, res) => {
    // DOCUSIGN_CONNECT_SKIP_HMAC=true bypasses signature verification — a TEST-ONLY escape hatch
    // for when the shared HMAC secret can't be confirmed (e.g. during first setup). Leaves the
    // endpoint unauthenticated, so turn it off again for real use.
    const skipHmac = process.env.DOCUSIGN_CONNECT_SKIP_HMAC === 'true';
    if (skipHmac) {
      console.warn('DOCUSIGN_CONNECT_SKIP_HMAC=true — webhook-signatur verifiseres IKKE (kun for test).');
    } else if (!config.demoMode) {
      const signature = req.headers['x-docusign-signature-1'];
      if (!docusign.verifyConnectSignature(req.rawBody, signature)) {
        // Diagnostics (safe: logs only lengths + short prefixes of derived signatures, never the key).
        const key = config.docusign.connectHmacKey || '';
        const computed = key && req.rawBody
          ? require('crypto').createHmac('sha256', key).update(req.rawBody).digest('base64')
          : '(mangler nokkel eller body)';
        console.error(
          `DocuSign webhook 401 (HMAC): header=${signature ? signature.slice(0, 12) + '…' : 'MANGLER'} `
          + `computed=${computed.slice(0, 12)}… keyLen=${key.length} bodyLen=${req.rawBody ? req.rawBody.length : 0}`,
        );
        return res.status(401).end();
      }
    }

    // Ack fast — DocuSign expects a quick 200 and retries on timeout, which the idempotency
    // design already has to handle regardless.
    res.status(200).json({ received: true });

    const event = docusign.parseWebhookEvent(req.body);
    handleEnvelopeEvent(event).catch((e) => console.error('handleEnvelopeEvent error:', e));
  });

// DEMO_MODE only — fabricates a signed event and runs it through the real webhook handler
// function above (not a duplicate code path), so the demo behaves exactly like production would.
app.post('/api/demo/simulate-signed/:row', async (req, res) => {
  if (!config.demoMode) return res.status(403).json({ success: false, error: 'Kun tilgjengelig i DEMO_MODE' });
  try {
    const candidate = await storage.getCandidate(req.params.row);
    if (!candidate) return res.status(404).json({ success: false, error: 'Ikke funnet' });

    let { docusignEnvelopeId: envelopeId } = candidate;
    if (!envelopeId) {
      envelopeId = `DEMO-ENVELOPE-${candidate.kandidatId}`;
      await storage.updateCandidateFields(candidate.row, {
        docusignEnvelopeId: envelopeId,
        statusKontrakt: KONTRAKT_STATUS.SENDT,
        kontraktSendtDato: new Date().toISOString(),
      });
    }

    const result = await handleEnvelopeEvent({ envelopeId, status: 'completed' });
    res.json({ success: true, result });
  } catch (e) {
    console.error('simulate-signed error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// BIRTHDAYS
// ============================================================

app.post('/api/check-birthdays', async (req, res) => {
  try {
    const result = await birthday.checkBirthdaysToday();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('check-birthdays error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// STATIC FRONTEND
// ============================================================

// no-cache on HTML so a browser tab always revalidates and picks up a new deploy on a normal
// refresh — otherwise a stale cached index.html keeps running old JS after a redeploy. Other
// assets can still be cached normally.
app.use(express.static(path.join(__dirname, 'public'), {
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html')) res.setHeader('Cache-Control', 'no-cache');
  },
}));

// ============================================================
// CRON JOBS (skipped in demo mode)
// ============================================================

if (config.demoMode) {
  console.log('Demo mode: cron jobs disabled, no external API calls.');
} else {
  cron.schedule('0 7 * * *', async () => {
    try {
      const result = await birthday.checkBirthdaysToday();
      console.log(`Cron: bursdagssjekk fullført, ${result.sent} varsel sendt.`);
    } catch (e) {
      console.error('Cron check-birthdays error:', e.message);
    }
  }, { timezone: 'Europe/Oslo' });
}

// ============================================================
// START SERVER
// ============================================================

const server = app.listen(config.port, () => {
  console.log(`Electi Onboarding listening on port ${config.port}`);
  console.log(`Test mode: ${config.testMode}`);
  console.log(`Demo mode: ${config.demoMode}`);
  if (!config.demoMode) {
    console.log('Cron jobs: bursdagssjekk (07:00 Europe/Oslo)');
  }
});

async function shutdown(signal) {
  console.log(`\n${signal} received. Shutting down...`);
  server.close(() => {
    console.log('Server closed.');
    process.exit(0);
  });
  setTimeout(() => { process.exit(1); }, 10000);
}

process.on('SIGTERM', () => shutdown('SIGTERM'));
process.on('SIGINT', () => shutdown('SIGINT'));
