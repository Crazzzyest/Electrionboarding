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
    <h3 style="margin-top:20px;">Tjenester du skal koble deg til</h3>
    <p>Bruk e-postadressen din <strong>${candidate.microsoftUpn}</strong> når du kobler deg til disse:</p>
    <ul>
      <li><strong>Hyre</strong> (bilutleie): <a href="${config.email.hyreJoinUrl}">${config.email.hyreJoinUrl}</a><br>
        Skriv inn @electi.no-adressen din, så mottar du en invitasjon til Electis bedriftskonto.</li>
      <li><strong>Airbnb for Work</strong> (reise): <a href="${config.email.airbnbJoinUrl}">${config.email.airbnbJoinUrl}</a><br>
        Legg til @electi.no-adressen din for å koble deg til Electis medarbeiderportal.</li>
    </ul>
    <p>Da kan Electi bestille bil og reise på vegne av deg.</p>
    <p><strong>SalesScreen</strong> (salgsdashbord): kontoen din er opprettet med @electi.no-adressen.
      Du får en egen e-post fra SalesScreen med lenke for å sette passord — følg den for å logge inn på
      <a href="https://app.salesscreen.com">app.salesscreen.com</a>.</p>
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
