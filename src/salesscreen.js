// Adapter: SalesScreen. Confirmed 2026-08-21 against their real Postman-published API spec
// (docs.salesscreen.com renders via client-side JS, so a plain fetch only sees an empty shell —
// had to read it through an actual browser and pull the collection JSON their docs page itself
// fetches). POST /User/Add is an upsert keyed on "userId": send the same userId again and it
// updates the existing user rather than erroring or duplicating, and it auto-creates the team if
// the team "key" doesn't exist yet. That upsert behavior is why there's no separate
// find-before-create step here, unlike the other adapters — there's no lookup endpoint to call
// even if we wanted one. Never touches storage — onboarding.js owns all Status_*/timestamp writes.
const config = require('./config');

async function createOrUpdateUser(candidate) {
  // Use the @electi.no WORK address as the SalesScreen id/login — SalesScreen is a work tool and the
  // employee signs in with their work identity (same as Hyre/Airbnb and the rest of onboarding). The
  // UPN is decided at registration, so it's available even before the mailbox exists. Falls back to
  // the private address only if a UPN is somehow missing.
  const loginEmail = candidate.microsoftUpn || candidate.privatEpost;
  const url = `${config.salesscreen.baseUrl}${config.salesscreen.createUserEndpoint}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { apiKey: config.salesscreen.apiKey, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      userId: loginEmail,
      email: loginEmail,
      firstname: candidate.fornavn,
      lastname: candidate.etternavn,
      team: config.salesscreen.team,
    }),
  });
  if (!res.ok) throw new Error(`SalesScreen POST /User/Add-feil: ${res.status} ${await res.text()}`);
  return res;
}

async function ensure(candidate, ctx) {
  const loginEmail = candidate.microsoftUpn || candidate.privatEpost;
  if (config.demoMode) {
    return { ok: true, externalId: loginEmail, demoMode: true };
  }

  if (!config.salesscreen.apiKey) {
    return {
      ok: false,
      error: 'SALESSCREEN_API_KEY er ikke satt — se docs/SETUP-CHECKLIST.md',
      retryable: true,
    };
  }

  try {
    await createOrUpdateUser(candidate);
    return { ok: true, externalId: loginEmail };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

module.exports = { ensure };
