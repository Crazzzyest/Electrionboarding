// Drift monitoring: a daily scan that emails a summary when any candidate has a failed step (or a
// contract that has sat unsigned too long), so problems are noticed instead of discovered by chance.
// Reconciliation of missed webhooks lives in index.js (it needs the shared handleEnvelopeEvent).
const config = require('./config');
const mail = require('./graph-mail');
const storage = require('./storage');
const { STEG_STATUS, KONTRAKT_STATUS } = require('./columns');

const STEP_FIELDS = [
  ['statusMicrosoft365', 'Microsoft365'],
  ['statusTelenor', 'Telenor'],
  ['statusSalesscreen', 'SalesScreen'],
  ['statusVelkommen', 'Velkommen'],
];

// Days a contract may sit as "Sendt" (unsigned) before it's flagged as possibly stuck.
const UNSIGNED_STALE_DAYS = 7;

function daysSince(iso) {
  if (!iso) return 0;
  const t = Date.parse(iso);
  if (Number.isNaN(t)) return 0;
  return (Date.now() - t) / 86400000;
}

async function checkAlerts() {
  const candidates = await storage.listCandidates();
  const problems = [];

  for (const c of candidates) {
    const failedSteps = STEP_FIELDS.filter(([f]) => c[f] === STEG_STATUS.FEILET).map(([, label]) => label);
    if (c.statusKontrakt === KONTRAKT_STATUS.FEILET) failedSteps.unshift('Kontrakt');
    const staleUnsigned = c.statusKontrakt === KONTRAKT_STATUS.SENDT
      && daysSince(c.kontraktSendtDato) > UNSIGNED_STALE_DAYS;
    if (failedSteps.length || staleUnsigned) {
      problems.push({ c, failedSteps, staleUnsigned });
    }
  }

  if (!problems.length) return { alerts: 0, sent: false };

  const to = config.alertEmail || config.email.managementEmail;
  if (config.demoMode) {
    console.log(`[DEMO] Drifts-varsel: ${problems.length} kandidat(er) med problemer.`);
    return { alerts: problems.length, sent: false, demo: true };
  }
  if (!to) {
    console.warn('checkAlerts: ingen mottaker (ALERT_EMAIL/managementEmail ikke satt).');
    return { alerts: problems.length, sent: false };
  }

  const rows = problems.map(({ c, failedSteps, staleUnsigned }) => {
    const bits = [];
    if (failedSteps.length) bits.push(`feilet: ${failedSteps.join(', ')}`);
    if (staleUnsigned) bits.push(`kontrakt usignert i &gt; ${UNSIGNED_STALE_DAYS} dager`);
    const feil = c.sisteFeilmelding ? ` — <em>${c.sisteFeilmelding}</em>` : '';
    return `<li><strong>${c.fornavn} ${c.etternavn}</strong> (${c.kandidatId}): ${bits.join('; ')}${feil}</li>`;
  }).join('');

  await mail.sendEmail(
    to,
    `Onboarding: ${problems.length} kandidat(er) trenger oppmerksomhet`,
    `<p>Følgende kandidater har et steg som feilet eller en kontrakt som henger:</p><ul>${rows}</ul>
     <p>Åpne onboarding-appen og bruk «Kjør på nytt» på de aktuelle stegene.</p>`,
  );
  return { alerts: problems.length, sent: true, to };
}

module.exports = { checkAlerts };
