// Orchestration for the offboarding flow (scope-addition after V1). Mirrors onboarding.js: it is
// the only module that writes the Offboarding table's Status_*/timestamp columns, adapters just
// return an {ok, ...} result. Steps are idempotent — an already-OK step is skipped and a retry is
// safe. Every policy choice (what to do with the M365 account, immediate vs scheduled, who gets
// the commission/SalesScreen notices) comes from config.offboarding, i.e. Sliplane env vars.
const config = require('./config');
const storage = require('./storage');
const microsoft = require('./microsoft');
const mail = require('./graph-mail');
const { STEG_STATUS, LOGG_HANDLING, LOGG_KILDE } = require('./columns');

const rowsInProgress = new Set();

function friendlyError(raw) {
  const m = String(raw || '');
  if (/Authorization_RequestDenied|Insufficient privileges|(^|\D)403(\D|$)/.test(m)) {
    return 'Venter på Microsoft-tillatelse fra Electis admin (se docs/MICROSOFT-ADMIN-SETUP.md).';
  }
  if (/Mail\.Send|sendAsMailbox|ikke konfigurert|not configured|ikke satt/i.test(m)) {
    return m.split('{')[0].trim();
  }
  return m.split('{')[0].trim() || 'Ukjent feil';
}

// ---- Step adapters --------------------------------------------------------------------------

async function offboardMicrosoft(o) {
  if (config.demoMode) return { ok: true, demoMode: true };
  if (!o.microsoftUpn) {
    return { ok: false, error: 'Mangler Electi e-post (UPN) — kan ikke finne kontoen', retryable: true };
  }
  try {
    const res = await microsoft.offboardUser(o.microsoftUpn, config.offboarding.microsoftAction);
    return { ok: true, details: res.details };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

function offboardTelenorHtml(o) {
  return `
    <p>Hei,</p>
    <p>Følgende selger slutter og skal avsluttes:</p>
    <p>
      ${o.navn}<br>
      ${o.microsoftUpn || ''}<br>
      Sluttdato: ${o.sluttdato || 'ikke oppgitt'}
    </p>
    <p>Mvh<br>${o.registrertAv || 'Electi'}</p>
  `;
}

async function offboardTelenor(o) {
  if (config.demoMode) return { ok: true, demoMode: true };
  const to = config.offboarding.telenorTo.length ? config.offboarding.telenorTo : config.telenor.orderTo;
  const cc = config.offboarding.telenorCc.length ? config.offboarding.telenorCc : config.telenor.orderCc;
  try {
    await mail.sendEmail(to, config.offboarding.telenorSubject, offboardTelenorHtml(o), { cc });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

// SalesScreen's connect API is write-only with no confirmed deactivation endpoint, so this step is
// an honest human handoff: a reminder mail asking someone to deactivate the user in the UI.
async function offboardSalesscreen(o) {
  if (config.demoMode) return { ok: true, demoMode: true };
  const to = config.offboarding.salesscreenVarselEpost;
  if (!to) {
    return {
      ok: false,
      error: 'OFFBOARDING_SALESSCREEN_EPOST (eller OFFBOARDING_PROVISJON_EPOST) er ikke satt',
      retryable: true,
    };
  }
  try {
    await mail.sendEmail(
      to,
      `Deaktiver i SalesScreen: ${o.navn}`,
      `<p>${o.navn} (${o.microsoftUpn || 'ukjent e-post'}) slutter ${o.sluttdato || '(dato ikke oppgitt)'}.</p>
       <p>SalesScreen-API-et støtter ikke deaktivering, så brukeren må deaktiveres manuelt i SalesScreen.</p>`,
    );
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

async function offboardProvisjon(o) {
  if (config.demoMode) return { ok: true, demoMode: true };
  const to = config.offboarding.provisjonEpost;
  if (!to) {
    return { ok: false, error: 'OFFBOARDING_PROVISJON_EPOST er ikke satt', retryable: true };
  }
  try {
    await mail.sendEmail(
      to,
      `Provisjonskrav ved avslutning: ${o.navn}`,
      `<p>${o.navn} slutter ${o.sluttdato || '(dato ikke oppgitt)'} og har krav på provisjon.</p>
       <p>Provisjonspapirene må behandles manuelt.</p>`,
    );
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

const STEP_ADAPTERS = {
  microsoft: { ensure: offboardMicrosoft, statusField: 'statusMicrosoft' },
  telenor: { ensure: offboardTelenor, statusField: 'statusTelenor' },
  salesscreen: { ensure: offboardSalesscreen, statusField: 'statusSalesscreen' },
  provisjon: { ensure: offboardProvisjon, statusField: 'statusProvisjon' },
};

// ---- Runner ---------------------------------------------------------------------------------

async function runOffboardingStep(row, stepName, { trigger = LOGG_KILDE.MANUELL } = {}) {
  const adapter = STEP_ADAPTERS[stepName];
  if (!adapter) throw new Error(`Ukjent offboarding-steg: ${stepName}`);

  const o = await storage.getOffboarding(row);
  if (!o) return { skipped: true, reason: 'Offboarding ikke funnet' };

  if (o[adapter.statusField] === STEG_STATUS.OK) {
    await storage.appendLog(o.offboardingId, `off:${stepName}`, LOGG_HANDLING.HOPPET_OVER, 'Allerede OK', trigger);
    return { skipped: true, ok: true, reason: 'Allerede fullført' };
  }

  await storage.updateOffboardingFields(row, { [adapter.statusField]: STEG_STATUS.KJORER });
  await storage.appendLog(o.offboardingId, `off:${stepName}`, LOGG_HANDLING.FORSOKT, '', trigger);

  let result;
  try {
    result = await adapter.ensure(o);
  } catch (e) {
    result = { ok: false, error: e.message, retryable: true };
  }

  if (result.ok) {
    await storage.updateOffboardingFields(row, { [adapter.statusField]: STEG_STATUS.OK, sisteFeilmelding: '' });
    await storage.appendLog(o.offboardingId, `off:${stepName}`, LOGG_HANDLING.FULLFORT, result.demoMode ? '(demo)' : '', trigger);
  } else {
    const melding = friendlyError(result.error || 'Ukjent feil');
    console.error(`Offboarding-steg "${stepName}" feilet:`, result.error);
    await storage.updateOffboardingFields(row, { [adapter.statusField]: STEG_STATUS.FEILET, sisteFeilmelding: melding });
    await storage.appendLog(o.offboardingId, `off:${stepName}`, LOGG_HANDLING.FEILET, melding, trigger);
  }

  return result;
}

async function runOffboarding(row, { trigger = LOGG_KILDE.MANUELL } = {}) {
  if (rowsInProgress.has(row)) return { skipped: true, reason: 'Kjører allerede' };
  rowsInProgress.add(row);

  try {
    const o = await storage.getOffboarding(row);
    if (!o) return { skipped: true, reason: 'Offboarding ikke funnet' };

    // All four steps are independent of one another, so run them in parallel. provisjon is a no-op
    // (pre-set to OK at creation) when no commission is owed, so it simply skips.
    const results = {};
    await Promise.all(
      ['microsoft', 'telenor', 'salesscreen', 'provisjon'].map(async (step) => {
        results[step] = await runOffboardingStep(row, step, { trigger });
      }),
    );

    // Stamp the completion date once every step is OK.
    const after = await storage.getOffboarding(row);
    const allOk = ['statusMicrosoft', 'statusTelenor', 'statusSalesscreen', 'statusProvisjon']
      .every((f) => after[f] === STEG_STATUS.OK);
    if (allOk && !after.utfortDato) {
      await storage.updateOffboardingFields(row, { utfortDato: new Date().toISOString() });
    }

    return { ok: true, results };
  } finally {
    rowsInProgress.delete(row);
  }
}

// Scheduled timing: run every offboarding whose sluttdato has arrived (<= today in Oslo) and that
// isn't already fully done. Idempotent — already-OK steps skip, so re-running daily is safe.
async function runDueOffboardings() {
  const { todayInTimezone } = require('./utils');
  const t = todayInTimezone('Europe/Oslo');
  const todayStr = `${t.year}-${String(t.month).padStart(2, '0')}-${String(t.day).padStart(2, '0')}`;

  const all = await storage.listOffboardings();
  let ran = 0;
  for (const o of all) {
    if (o.utfortDato) continue; // already complete
    if (!o.sluttdato) continue; // no date to be due against
    if (o.sluttdato.slice(0, 10) > todayStr) continue; // not due yet
    await runOffboarding(o.row, { trigger: LOGG_KILDE.CRON });
    ran += 1;
  }
  return { ran };
}

module.exports = { runOffboarding, runOffboardingStep, runDueOffboardings };
