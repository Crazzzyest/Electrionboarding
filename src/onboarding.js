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

// Turns a raw adapter/Graph error into a short, human message for the UI and the log. The raw
// error is still written to the server console for debugging — the client-facing status just
// should not show JSON blobs or stack traces, especially for steps that are simply waiting on an
// admin-granted Microsoft permission rather than actually broken.
function friendlyError(raw) {
  const m = String(raw || '');
  if (/Mail\.Send|Mail\.Read|sendAsMailbox/i.test(m)) {
    return 'E-post ikke aktivert ennå — venter på Mail-tillatelse og avsender-postboks fra Electis admin.';
  }
  if (/Authorization_RequestDenied|Insufficient privileges|(^|\D)403(\D|$)/.test(m)) {
    return 'Venter på Microsoft-tillatelse fra Electis admin (se docs/MICROSOFT-ADMIN-SETUP.md).';
  }
  if (/ikke konfigurert|not configured|ikke satt/i.test(m)) {
    return m.split('{')[0].trim(); // already human — just drop any JSON tail
  }
  return m.split('{')[0].trim() || 'Ukjent feil';
}

async function runStep(row, stepName, { force = false, trigger = LOGG_KILDE.MANUELL } = {}) {
  const adapter = STEP_ADAPTERS[stepName];
  if (!adapter) throw new Error(`Ukjent steg: ${stepName}`);

  const candidate = await storage.getCandidate(row);
  if (!candidate) return { skipped: true, reason: 'Kandidat ikke funnet' };

  // NB: steg gates ikke lenger på signert kontrakt — kontrakten er nå en egen, sidestilt flyt
  // (se runOnboardingSteps). De eneste rekkefølge-kravene som står igjen er de tekniske
  // dependsOn (telenor/velkommen trenger Microsoft-kontoen). `force` beholdes som manuell
  // override for edge-cases.
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
    // Clear any stale error from an earlier failed attempt now that the step has succeeded.
    const updates = { [adapter.statusField]: STEG_STATUS.OK, sisteFeilmelding: '' };
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
    const raw = result.error || 'Ukjent feil';
    console.error(`Steg "${stepName}" feilet:`, raw);
    const melding = friendlyError(raw);
    await storage.updateCandidateFields(row, {
      [adapter.statusField]: STEG_STATUS.FEILET,
      sisteFeilmelding: melding,
    });
    await storage.appendLog(candidate.kandidatId, stepName, LOGG_HANDLING.FEILET, melding, trigger);
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

    // Uavhengige, sidestilte flyter — de starter så snart skjemaet sendes inn, ikke etter en
    // signert kontrakt. To grener kjører i parallell:
    //   • microsoft365 → deretter telenor + velkommen (som teknisk trenger Microsoft-kontoen)
    //   • salesscreen (helt uavhengig)
    // Hvert steg er idempotent og hopper over seg selv hvis det allerede er OK, så en retry eller
    // et senere webhook-kall gjør ingen skade.
    const results = {};
    const microsoftBranch = (async () => {
      results.microsoft365 = await runStep(row, 'microsoft365', { trigger });
      const [tel, vel] = await Promise.all([
        runStep(row, 'telenor', { trigger }),
        runStep(row, 'velkommen', { trigger }),
      ]);
      results.telenor = tel;
      results.velkommen = vel;
    })();
    const salesscreenBranch = (async () => {
      results.salesscreen = await runStep(row, 'salesscreen', { trigger });
    })();

    await Promise.all([microsoftBranch, salesscreenBranch]);

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
    const raw = result.error || 'Ukjent feil';
    console.error('Steg "kontrakt" feilet:', raw);
    const melding = friendlyError(raw);
    await storage.updateCandidateFields(row, {
      statusKontrakt: KONTRAKT_STATUS.FEILET,
      sisteFeilmelding: melding,
    });
    await storage.appendLog(candidate.kandidatId, 'kontrakt', LOGG_HANDLING.FEILET, melding, LOGG_KILDE.REGISTRERING);
  }

  return result;
}

module.exports = { runStep, runOnboardingSteps, sendContract };
