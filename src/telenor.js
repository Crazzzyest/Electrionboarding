// Adapter: Telenor order email. No external API — "integration" here means composing and sending
// the same mail Magnus sends by hand today, to Telenor's order contacts.
//
// Modelled on a real example from 2026-08-21. The real mail is deliberately minimal: subject
// "Ny selger", body listing name, phone and the new hire's @electi.no address. Anything beyond
// that (address, department, start date, product/system names) is NOT in Electi's actual mail, so
// it is not invented here either.
//
// Note this needs the new hire's *work* address (@electi.no), not their private one — which is why
// onboarding.js runs this step after Microsoft365, so the address is real rather than predicted.
// Never touches storage — onboarding.js owns all Status_*/timestamp writes.
const config = require('./config');
const mail = require('./graph-mail');

// Stable token used only for the duplicate-send check. Kept out of the subject line so the mail
// Telenor receives looks exactly like the one Magnus sends by hand; it rides along in the body
// instead, where a mail search still finds it.
function idToken(kandidatId) {
  return `[ELECTI-ONB-${kandidatId}]`;
}

function formatPhone(mobil) {
  const digits = String(mobil || '').replace(/\D/g, '');
  return digits.length === 8 ? digits : String(mobil || '').trim();
}

function buildOrderHtml(candidate) {
  return `
    <p>Hei,</p>
    <p>Trenger brukere til f&oslash;lgende:</p>
    <p>
      ${candidate.fornavn} ${candidate.etternavn}<br>
      ${formatPhone(candidate.mobil)}<br>
      ${candidate.microsoftUpn}
    </p>
    <p>Mvh<br>${candidate.registrertAv || 'Electi'}</p>
    <p style="color:#ffffff;font-size:1px;">${idToken(candidate.kandidatId)}</p>
  `;
}

async function ensure(candidate, ctx) {
  if (config.demoMode) {
    return { ok: true, demoMode: true };
  }

  // The order is useless to Telenor without the work address, and it is the one field we cannot
  // fabricate — fail loudly rather than send an order they cannot action.
  if (!candidate.microsoftUpn) {
    return {
      ok: false,
      error: 'Mangler @electi.no-adresse (Microsoft365-steget må fullføres først)',
      retryable: true,
    };
  }

  try {
    const token = idToken(candidate.kandidatId);
    if (await mail.searchMail(token)) {
      return { ok: true, details: { alreadySent: true } };
    }

    await mail.sendEmail(
      config.telenor.orderTo,
      config.telenor.subject,
      buildOrderHtml(candidate),
      { cc: config.telenor.orderCc },
    );
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

module.exports = { ensure };
