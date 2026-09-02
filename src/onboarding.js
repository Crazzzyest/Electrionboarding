// Orchestration / state machine. The only module that writes Status_*/timestamp columns —
// adapters never touch storage directly, which keeps them swappable and prevents split-brain
// writes. The DocuSign webhook, the manual retry endpoint, and (indirectly) the demo "simulate
// signed" endpoint all converge on runStep/runOnboardingSteps, which is what makes "webhook
// event or manual retry" behave identically and safely.
const storage = require('./storage');
const docusign = require('./docusign');
const microsoft = require('./microsoft');
const telenor = require('./telenor');
const salesscreen = require('./salesscreen');
const { sendWelcomeEmail } = require('./emails');
const { STEG_STATUS, KONTRAKT_STATUS, LOGG_HANDLING, LOGG_KILDE } = require('./columns');

// Simple in-memory mutex: sufficient at this scale, and self-heals on crash/restart (unlike a
// sheet-persisted "Kjører" lock, which could get stuck forever if the process died mid-step).
const rowsInProgress = new Set();

// The welcome email needs a *valid* temp password at the moment it sends, not the one
// generated when the Microsoft365 step ran (never persisted, and possibly long gone if this
// step runs later or is retried independently) — so it mints a fresh one via Graph right here.
async function ensureVelkommen(candidate, ctx) {
  try {
    const tempPassword = await microsoft.resetTempPassword(candidate.microsoftUserId);
    return await sendWelcomeEmail(candidate, { ...ctx, tempPassword });
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

// extraUpdates(result) lets a step persist more than just its status/timestamp — e.g. Microsoft
// returns both a Graph object id AND a UPN, and the UPN (not the opaque id) is what the welcome
// email needs to show. Keyed on the *candidate field name*, not the sheet column, same as every
// other write path in this file.
const STEP_ADAPTERS = {
  microsoft365: {
    ensure: (c, ctx) => microsoft.ensure(c, ctx),
    statusField: 'statusMicrosoft365',
    doneField: 'microsoft365FullfortDato',
    extraUpdates: (result) => ({
      microsoftUserId: result.externalId,
      microsoftUpn: result.details && result.details.upn,
    }),
  },
  telenor: {
    ensure: (c, ctx) => telenor.ensure(c, ctx),
    statusField: 'statusTelenor',
    doneField: 'telenorBestillingSendtDato',
    // The real order mail states the new hire's @electi.no address. The address is known from
    // registration, but waiting for Microsoft365 means we send Telenor an address that actually
    // exists, instead of one we merely predicted and which may have failed to be created.
    dependsOn: 'microsoft365',
  },
  salesscreen: {
    ensure: (c, ctx) => salesscreen.ensure(c, ctx),
    statusField: 'statusSalesscreen',
    doneField: 'salesscreenFullfortDato',
    extraUpdates: (result) => ({ salesscreenUserId: result.externalId }),
  },
  velkommen: {
    ensure: (c, ctx) => ensureVelkommen(c, ctx),
    statusField: 'statusVelkommen',
    doneField: 'velkommenSendtDato',
    dependsOn: 'microsoft365', // needs the UPN + a freshly minted temp password
  },
};

async function runStep(row, stepName, { force = false, trigger = LOGG_KILDE.MANUELL } = {}) {
  const adapter = STEP_ADAPTERS[stepName];
  if (!adapter) throw new Error(`Ukjent steg: ${stepName}`);

  const candidate = await storage.getCandidate(row);
  if (!candidate) return { skipped: true, reason: 'Kandidat ikke funnet' };

  if (!force && candidate.statusKontrakt !== KONTRAKT_STATUS.SIGNERT) {
    return { skipped: true, reason: 'Venter på signert kontrakt' };
  }

  if (candidate[adapter.statusField] === STEG_STATUS.OK) {
    await storage.appendLog(candidate.kandidatId, stepName, LOGG_HANDLING.HOPPET_OVER, 'Allerede OK', trigger);
    return { skipped: true, ok: true, reason: 'Allerede fullført' };
  }

  if (adapter.dependsOn) {
    const dep = STEP_ADAPTERS[adapter.dependsOn];
    if (candidate[dep.statusField] !== STEG_STATUS.OK) {
      return { skipped: true, reason: `Venter på ${adapter.dependsOn}` };
    }
  }

  await storage.updateCandidateFields(row, { [adapter.statusField]: STEG_STATUS.KJORER });
  await storage.appendLog(candidate.kandidatId, stepName, LOGG_HANDLING.FORSOKT, '', trigger);

  let result;
  try {
    result = await adapter.ensure(candidate, { kandidatId: candidate.kandidatId, trigger });
  } catch (e) {
    result = { ok: false, error: e.message, retryable: true };
  }

  if (result.ok) {
    const updates = { [adapter.statusField]: STEG_STATUS.OK };
    if (adapter.doneField) updates[adapter.doneField] = new Date().toISOString();
    if (adapter.extraUpdates) {
      for (const [field, value] of Object.entries(adapter.extraUpdates(result))) {
        if (value !== undefined && value !== null) updates[field] = value;
      }
    }
    await storage.updateCandidateFields(row, updates);
    await storage.appendLog(
      candidate.kandidatId, stepName, LOGG_HANDLING.FULLFORT,
      result.demoMode ? '(demo)' : '', trigger,
    );
  } else {
    await storage.updateCandidateFields(row, {
      [adapter.statusField]: STEG_STATUS.FEILET,
      sisteFeilmelding: result.error || 'Ukjent feil',
    });
    await storage.appendLog(candidate.kandidatId, stepName, LOGG_HANDLING.FEILET, result.error || 'Ukjent feil', trigger);
  }

  return result;
}

async function runOnboardingSteps(row, { trigger = LOGG_KILDE.MANUELL } = {}) {
  if (rowsInProgress.has(row)) {
    return { skipped: true, reason: 'Kjører allerede' };
  }
  rowsInProgress.add(row);

  try {
    const candidate = await storage.getCandidate(row);
    if (!candidate) return { skipped: true, reason: 'Kandidat ikke funnet' };
    if (candidate.statusKontrakt !== KONTRAKT_STATUS.SIGNERT) {
      return { skipped: true, reason: 'Venter på signert kontrakt' };
    }

    // microsoft365 and salesscreen are independent of each other; telenor and velkommen both
    // need the Microsoft account to exist first (see dependsOn in STEP_ADAPTERS), so they run
    // after it rather than alongside.
    const results = {};
    for (const step of ['microsoft365', 'salesscreen', 'telenor', 'velkommen']) {
      results[step] = await runStep(row, step, { trigger });
    }

    return { ok: true, results };
  } finally {
    rowsInProgress.delete(row);
  }
}

async function sendContract(row) {
  const candidate = await storage.getCandidate(row);
  if (!candidate) return { skipped: true, reason: 'Kandidat ikke funnet' };

  if (candidate.docusignEnvelopeId) {
    return { skipped: true, ok: true, reason: 'Kontrakt allerede sendt' };
  }

  await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FORSOKT, '', LOGG_KILDE.REGISTRERING);

  let result;
  try {
    result = await docusign.ensure(candidate, { kandidatId: candidate.kandidatId });
  } catch (e) {
    result = { ok: false, error: e.message, retryable: true };
  }

  if (result.ok) {
    await storage.updateCandidateFields(row, {
      statusKontrakt: KONTRAKT_STATUS.SENDT,
      docusignEnvelopeId: result.externalId,
      kontraktSendtDato: new Date().toISOString(),
    });
    await storage.appendLog(
      candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FULLFORT,
      result.demoMode ? '(demo)' : '', LOGG_KILDE.REGISTRERING,
    );
  } else {
    await storage.updateCandidateFields(row, {
      statusKontrakt: KONTRAKT_STATUS.FEILET,
      sisteFeilmelding: result.error || 'Ukjent feil',
    });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FEILET, result.error || 'Ukjent feil', LOGG_KILDE.REGISTRERING);
  }

  return result;
}

module.exports = { runStep, runOnboardingSteps, sendContract };
