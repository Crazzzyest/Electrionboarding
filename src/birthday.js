// Daily cron logic: scans Ansatte for signed candidates whose birthday is today and who
// haven't already gotten a notification this year, emails management, and stamps the year as
// the idempotency key. Not part of the onboarding step graph in onboarding.js — a fully
// separate daily process over the same roster.
const config = require('./config');
const mail = require('./graph-mail');
const storage = require('./storage');
const { buildBirthdayEmailHtml } = require('./emails');
const { KONTRAKT_STATUS } = require('./columns');
const { todayInTimezone, isSameMonthDay } = require('./utils');

async function checkBirthdaysToday() {
  const candidates = await storage.listCandidates();
  const today = todayInTimezone('Europe/Oslo');
  let sent = 0;

  for (const c of candidates) {
    if (c.statusKontrakt !== KONTRAKT_STATUS.SIGNERT) continue; // only actual hires
    if (!isSameMonthDay(c.fodselsdato, today)) continue;
    if (Number(c.sisteBursdagsvarselAr) === today.year) continue; // already notified this year

    if (config.demoMode) {
      console.log(`[DEMO] Bursdagsvarsel: ${c.fornavn} ${c.etternavn}`);
    } else if (!config.email.managementEmail) {
      console.error('Bursdagsvarsel: managementEmail ikke konfigurert — se docs/SETUP-CHECKLIST.md');
      continue;
    } else {
      await mail.sendEmail(
        config.email.managementEmail,
        `Bursdag i dag: ${c.fornavn} ${c.etternavn}`,
        buildBirthdayEmailHtml(c),
      );
    }

    await storage.updateCandidateFields(c.row, { sisteBursdagsvarselAr: String(today.year) });
    sent++;
  }

  return { sent };
}

// Sends a sample birthday notification to managementEmail regardless of whether anyone actually has
// a birthday today — so the recipient/mail setup and the template can be verified on demand from the
// UI. The real daily scan is checkBirthdaysToday(); this only exercises the notification itself.
async function sendTestBirthday() {
  const sample = { fornavn: 'Test', etternavn: 'Testesen', avdeling: 'Salg' };
  if (config.demoMode) {
    console.log('[DEMO] Test-bursdagsvarsel (ingen ekte e-post sendt)');
    return { sent: 0, demo: true, to: config.email.managementEmail };
  }
  if (!config.email.managementEmail) {
    throw new Error('managementEmail er ikke satt (sett MANAGEMENT_EMAIL) — ingen mottaker for bursdagsvarsel.');
  }
  await mail.sendEmail(
    config.email.managementEmail,
    'TEST — Bursdagsvarsel (Electi onboarding)',
    `<p><em>Dette er en test av bursdagsvarselet.</em></p>${buildBirthdayEmailHtml(sample)}`,
  );
  return { sent: 1, to: config.email.managementEmail };
}

module.exports = { checkBirthdaysToday, sendTestBirthday };
