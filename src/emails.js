// HTML builders for the welcome email and birthday notification, plus sendWelcomeEmail — wired
// into onboarding.js's "velkommen" step with the same {ok, ...} shape as the real integration
// adapters, even though sending an email isn't a third-party integration in the same sense.
const config = require('./config');
const mail = require('./graph-mail');

function buildWelcomeEmailHtml(candidate, tempPassword) {
  return `
    <p>Hei ${candidate.fornavn},</p>
    <p>Velkommen til Electi! Din nye Microsoft-konto er klar:</p>
    <p>
      <strong>Logg inn:</strong> <a href="https://portal.office.com">https://portal.office.com</a><br>
      <strong>Brukernavn:</strong> ${candidate.microsoftUpn}<br>
      <strong>Midlertidig passord:</strong> ${tempPassword}<br>
      (du blir bedt om å bytte passord ved første innlogging)
    </p>
    <p>E-posten din finner du på <a href="https://outlook.office.com">https://outlook.office.com</a>. Merk: det kan ta noen minutter før innboksen din blir tilgjengelig etter at kontoen er opprettet.</p>
    <p>Vi gleder oss til å ha deg med på laget!</p>
  `;
}

function buildBirthdayEmailHtml(candidate) {
  return `
    <p>${candidate.fornavn} ${candidate.etternavn} har bursdag i dag! 🎉</p>
    <p>Avdeling: ${candidate.avdeling}</p>
  `;
}

// ctx.tempPassword is provided by onboarding.js's "velkommen" step wiring, which mints a fresh
// one-time password via microsoft.resetTempPassword() right before calling this — see the note
// there on why the original account-creation password can't just be threaded through instead.
async function sendWelcomeEmail(candidate, ctx) {
  if (config.demoMode) {
    return { ok: true, demoMode: true };
  }
  try {
    await mail.sendEmail(
      candidate.privatEpost,
      'Velkommen til Electi!',
      buildWelcomeEmailHtml(candidate, ctx.tempPassword),
    );
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

module.exports = { buildWelcomeEmailHtml, buildBirthdayEmailHtml, sendWelcomeEmail };
