const express = require('express');
const path = require('path');
const cron = require('node-cron');
const config = require('./config');
const storage = require('./storage');
const onboarding = require('./onboarding');
const offboarding = require('./offboarding');
const docusign = require('./docusign');
const contractArchive = require('./contract-archive');
const birthday = require('./birthday');
const monitoring = require('./monitoring');
const auth = require('./auth');
const configCheck = require('./config-check');
const { validateCandidate } = require('./validation');
const { KONTRAKT_STATUS, LOGG_HANDLING, LOGG_KILDE } = require('./columns');

const app = express();

// Access control (Basic Auth) in front of everything except /health and the DocuSign webhook.
// No-op when APP_PASSWORD isn't set (a loud warning is logged at startup).
app.use(auth.gate);

// ============================================================
// HEALTH
// ============================================================

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    testMode: config.testMode,
    demoMode: config.demoMode,
    checks: config.demoMode ? undefined : configCheck.checks(),
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
  'fornavn', 'etternavn', 'fodselsdato', 'privatEpost', 'mobil', 'kontonummer',
  'stilling', 'stillingsprosent', 'avdeling', 'naermesteLeder', 'registrertAv',
];

app.post('/api/candidates', express.json(), async (req, res) => {
  try {
    const missing = REQUIRED_FIELDS.filter((f) => !req.body[f]);
    if (missing.length) {
      return res.status(400).json({ success: false, error: `Mangler felt: ${missing.join(', ')}` });
    }

    const formatErrors = validateCandidate(req.body);
    if (formatErrors.length) {
      return res.status(400).json({ success: false, error: formatErrors.join(' ') });
    }

    // Duplicate guard: same private e-post, or same name + birthdate, already registered.
    // Skipped when ALLOW_DUPLICATES=true (testing). On in production.
    const existing = config.allowDuplicates ? [] : await storage.listCandidates();
    const epost = String(req.body.privatEpost || '').trim().toLowerCase();
    const dupe = existing.find((c) => (
      (epost && String(c.privatEpost || '').trim().toLowerCase() === epost)
      || (String(c.fornavn || '').trim().toLowerCase() === String(req.body.fornavn || '').trim().toLowerCase()
        && String(c.etternavn || '').trim().toLowerCase() === String(req.body.etternavn || '').trim().toLowerCase()
        && c.fodselsdato === req.body.fodselsdato)
    ));
    if (dupe) {
      return res.status(409).json({ success: false, error: `Kandidaten finnes allerede (${dupe.kandidatId}).` });
    }

    // Decide the @electi.no address up front — the employment contract states it, so it has to
    // exist before the envelope is sent, not after the Microsoft account is created.
    const microsoft = require('./microsoft');
    const candidate = await storage.createCandidate({
      ...req.body,
      microsoftUpn: microsoft.buildUpn(req.body),
    });
    await storage.appendLog(candidate.kandidatId, 'registrering', LOGG_HANDLING.FULLFORT, 'Kandidat registrert', LOGG_KILDE.REGISTRERING);

    // Registrering trigger nå ALLE flytene med én gang, som uavhengige løp — kontrakten sendes
    // (egen flyt) parallelt med Microsoft/SalesScreen/Telenor/Velkomst, i stedet for at de venter
    // på en signert kontrakt. Begge kjører fire-and-forget slik at skjemaet svarer umiddelbart.
    onboarding.sendContract(candidate.row).catch((e) => console.error('sendContract error:', e));
    onboarding.runOnboardingSteps(candidate.row, { trigger: LOGG_KILDE.REGISTRERING })
      .catch((e) => console.error('runOnboardingSteps error:', e));

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

// Edit a candidate's data. Only whitelisted fields are writable. If the Microsoft account hasn't
// been created yet, the @electi.no address is recomputed from the (possibly changed) name so the
// contract and later provisioning stay consistent; once the account exists, the UPN is left alone.
const EDITABLE_FIELDS = [
  'fornavn', 'etternavn', 'fodselsdato', 'privatEpost', 'mobil', 'kontonummer', 'adresse',
  'stilling', 'stillingsprosent', 'avdeling', 'naermesteLeder', 'startdato', 'registrertAv',
];

app.put('/api/candidates/:row', express.json(), async (req, res) => {
  try {
    const candidate = await storage.getCandidate(req.params.row);
    if (!candidate) return res.status(404).json({ success: false, error: 'Ikke funnet' });

    const updates = {};
    for (const f of EDITABLE_FIELDS) {
      if (req.body[f] !== undefined) updates[f] = req.body[f];
    }
    if (!Object.keys(updates).length) {
      return res.status(400).json({ success: false, error: 'Ingen felter å oppdatere' });
    }

    const formatErrors = validateCandidate(updates);
    if (formatErrors.length) {
      return res.status(400).json({ success: false, error: formatErrors.join(' ') });
    }

    // Recompute the work address from the new name only while the account doesn't exist yet.
    const microsoft = require('./microsoft');
    if (candidate.statusMicrosoft365 !== 'OK') {
      const merged = { ...candidate, ...updates };
      updates.microsoftUpn = microsoft.buildUpn(merged);
    }

    await storage.updateCandidateFields(candidate.row, updates);
    await storage.appendLog(candidate.kandidatId, 'registrering', LOGG_HANDLING.FULLFORT, 'Kandidat redigert', LOGG_KILDE.MANUELL);
    const updated = await storage.getCandidate(candidate.row);
    res.json({ success: true, candidate: updated });
  } catch (e) {
    console.error('edit candidate error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/candidates/:row/resend-contract', async (req, res) => {
  try {
    // Resend: void any previous (unsigned) envelope and send a fresh contract with current data.
    const result = await onboarding.sendContract(req.params.row, { resend: true });
    res.json({ success: true, result });
  } catch (e) {
    console.error('resend-contract error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// ============================================================
// OFFBOARDING (egen fane)
// ============================================================

app.get('/api/offboardings', async (req, res) => {
  try {
    const offboardings = await storage.listOffboardings();
    res.json({ success: true, offboardings });
  } catch (e) {
    console.error('list offboardings error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.get('/api/offboardings/:row', async (req, res) => {
  try {
    const off = await storage.getOffboarding(req.params.row);
    if (!off) return res.status(404).json({ success: false, error: 'Ikke funnet' });
    const log = await storage.listLog(off.offboardingId);
    res.json({ success: true, offboarding: off, log });
  } catch (e) {
    console.error('get offboarding error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

const OFF_REQUIRED_FIELDS = ['navn', 'microsoftUpn', 'sluttdato', 'registrertAv'];

app.post('/api/offboardings', express.json(), async (req, res) => {
  try {
    const missing = OFF_REQUIRED_FIELDS.filter((f) => !req.body[f]);
    if (missing.length) {
      return res.status(400).json({ success: false, error: `Mangler felt: ${missing.join(', ')}` });
    }

    const off = await storage.createOffboarding({
      ...req.body,
      harProvisjon: req.body.harProvisjon === true || req.body.harProvisjon === 'true' || req.body.harProvisjon === 'Ja',
    });
    await storage.appendLog(off.offboardingId, 'offboarding', LOGG_HANDLING.FULLFORT, 'Registrert', LOGG_KILDE.REGISTRERING);

    // Timing policy: 'immediate' fires the steps now; 'scheduled' waits for the daily cron to pick
    // it up once sluttdato has arrived.
    if (config.offboarding.timing !== 'scheduled') {
      offboarding.runOffboarding(off.row, { trigger: LOGG_KILDE.REGISTRERING })
        .catch((e) => console.error('runOffboarding error:', e));
    }

    res.json({ success: true, offboarding: off, scheduled: config.offboarding.timing === 'scheduled' });
  } catch (e) {
    console.error('create offboarding error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

app.post('/api/offboardings/:row/retry', express.json(), async (req, res) => {
  try {
    const { step } = req.body || {};
    const { row } = req.params;
    const result = step
      ? await offboarding.runOffboardingStep(row, step, { trigger: LOGG_KILDE.MANUELL })
      : await offboarding.runOffboarding(row, { trigger: LOGG_KILDE.MANUELL });
    res.json({ success: true, result });
  } catch (e) {
    console.error('offboarding retry error:', e);
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
    // Registrering trigger allerede alle stegene, så signering skal KUN registrere at kontrakten
    // er signert — den kjører dem bevisst ikke på nytt (et tidligere feilet steg skal rettes via
    // "Kjør på nytt" i UI, ikke fyres automatisk av en signering).
    await storage.updateCandidateFields(candidate.row, {
      statusKontrakt: KONTRAKT_STATUS.SIGNERT,
      kontraktSignertDato: new Date().toISOString(),
    });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FULLFORT, 'Signert', LOGG_KILDE.WEBHOOK);

    // Archive the signed PDF to SharePoint. Idempotent: skip if already stored (webhook retry).
    if (!candidate.signertKontraktUrl) {
      try {
        const arch = await contractArchive.archiveSignedContract(candidate, event.envelopeId);
        if (arch.webUrl) {
          await storage.updateCandidateFields(candidate.row, { signertKontraktUrl: arch.webUrl });
        }
        await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FULLFORT,
          arch.demoMode ? 'Signert kontrakt (demo, ikke arkivert)' : 'Signert kontrakt lagret i SharePoint', LOGG_KILDE.WEBHOOK);
      } catch (e) {
        console.error('Arkivering av signert kontrakt feilet:', e);
        await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FEILET,
          `Kunne ikke lagre signert kontrakt: ${e.message}`, LOGG_KILDE.WEBHOOK);
      }
    }
  } else if (event.status === 'declined' || event.status === 'voided') {
    await storage.updateCandidateFields(candidate.row, { statusKontrakt: KONTRAKT_STATUS.AVSLATT });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FEILET, `Status: ${event.status}`, LOGG_KILDE.WEBHOOK);
  }

  return { matched: true };
}

// Reconciliation: catch signings whose webhook was missed (server down, transient failure). Polls
// DocuSign for the status of every candidate still on "Sendt" and, if it has moved on, runs the same
// handleEnvelopeEvent path the webhook would have. Idempotent — a truly-still-sent envelope is a no-op.
async function reconcilePendingContracts() {
  if (config.demoMode) return { checked: 0, updated: 0 };
  const candidates = await storage.listCandidates();
  const pending = candidates.filter((c) => c.statusKontrakt === KONTRAKT_STATUS.SENDT && c.docusignEnvelopeId);
  let updated = 0;
  for (const c of pending) {
    try {
      const status = await docusign.getEnvelopeStatus(c.docusignEnvelopeId);
      if (status && status !== 'sent' && status !== 'delivered' && status !== 'created') {
        await handleEnvelopeEvent({ envelopeId: c.docusignEnvelopeId, status });
        updated += 1;
      }
    } catch (e) {
      console.error(`reconcile envelope ${c.docusignEnvelopeId} feilet:`, e.message);
    }
  }
  return { checked: pending.length, updated };
}

app.post('/webhooks/docusign',
  express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }),
  async (req, res) => {
    // DOCUSIGN_CONNECT_SKIP_HMAC=true bypasses signature verification — a TEST-ONLY escape hatch
    // for when the shared HMAC secret can't be confirmed (e.g. during first setup). It is IGNORED in
    // production (docusign.env === 'production'): prod always verifies the signature even if the flag
    // is accidentally left on, so the webhook can never run unauthenticated against real data.
    const skipHmac = process.env.DOCUSIGN_CONNECT_SKIP_HMAC === 'true' && config.docusign.env !== 'production';
    if (process.env.DOCUSIGN_CONNECT_SKIP_HMAC === 'true' && config.docusign.env === 'production') {
      console.warn('DOCUSIGN_CONNECT_SKIP_HMAC=true ignoreres i produksjon — signatur verifiseres alltid.');
    }
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

// Sends a sample birthday notification to managementEmail on demand (UI button), to verify setup.
app.post('/api/test-birthday', async (req, res) => {
  try {
    const result = await birthday.sendTestBirthday();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('test-birthday error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// Reconcile contracts stuck on "Sendt" against DocuSign (catches missed webhooks). Manual trigger;
// also runs on a schedule (see cron below).
app.post('/api/reconcile', async (req, res) => {
  try {
    const result = await reconcilePendingContracts();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('reconcile error:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// Scan for failed/stuck steps and email a summary. Manual trigger; also runs on a schedule.
app.post('/api/check-alerts', async (req, res) => {
  try {
    const result = await monitoring.checkAlerts();
    res.json({ success: true, ...result });
  } catch (e) {
    console.error('check-alerts error:', e);
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

  // Only relevant when offboarding uses scheduled timing — runs the ones whose sluttdato has come.
  if (config.offboarding.timing === 'scheduled') {
    cron.schedule('0 7 * * *', async () => {
      try {
        const result = await offboarding.runDueOffboardings();
        console.log(`Cron: offboarding-sjekk fullført, ${result.ran} kjørt.`);
      } catch (e) {
        console.error('Cron offboarding error:', e.message);
      }
    }, { timezone: 'Europe/Oslo' });
  }

  // Reconcile stuck contracts every 30 min (catches missed webhooks).
  cron.schedule('*/30 * * * *', async () => {
    try {
      const result = await reconcilePendingContracts();
      if (result.updated) console.log(`Cron: avstemming — ${result.updated} av ${result.checked} oppdatert.`);
    } catch (e) {
      console.error('Cron reconcile error:', e.message);
    }
  }, { timezone: 'Europe/Oslo' });

  // Daily drift alert at 08:00 — emails a summary if any step is failed or a contract is stuck.
  cron.schedule('0 8 * * *', async () => {
    try {
      const result = await monitoring.checkAlerts();
      if (result.alerts) console.log(`Cron: drifts-varsel — ${result.alerts} problem(er)${result.sent ? ' (sendt)' : ''}.`);
    } catch (e) {
      console.error('Cron check-alerts error:', e.message);
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
    console.log('Cron jobs: bursdagssjekk (07:00), avstemming (hver 30. min), drifts-varsel (08:00) — Europe/Oslo');
  }
  configCheck.logStartupChecks();
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
