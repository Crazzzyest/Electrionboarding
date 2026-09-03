// Adapter: Microsoft Graph (app-only / client-credentials — Electi's own Azure AD tenant).
// Does idempotent sub-checks internally so a partially-failed prior run never redoes completed
// work: find-or-create user, assign license if missing, add to groups if missing. Never touches
// storage — onboarding.js owns all Status_*/timestamp writes (see the adapter contract there).
const config = require('./config');
const { slugifyName, generateTempPassword } = require('./utils');
const { GRAPH_BASE, graphRequest } = require('./graph-client');

function buildUpn(candidate) {
  const local = `${slugifyName(candidate.fornavn)}.${slugifyName(candidate.etternavn)}`;
  return `${local}@${config.microsoft.domain}`;
}

async function findUserByUpn(upn) {
  const res = await graphRequest('GET', `/users/${encodeURIComponent(upn)}?$select=id,userPrincipalName,assignedLicenses`);
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Graph GET user-feil: ${res.status} ${await res.text()}`);
  return res.json();
}

async function createUser(candidate, upn) {
  const tempPassword = generateTempPassword();
  const body = {
    accountEnabled: true,
    displayName: `${candidate.fornavn} ${candidate.etternavn}`,
    mailNickname: `${slugifyName(candidate.fornavn)}.${slugifyName(candidate.etternavn)}`,
    userPrincipalName: upn,
    passwordProfile: {
      forceChangePasswordNextSignIn: true,
      password: tempPassword,
    },
    usageLocation: config.microsoft.usageLocation,
    givenName: candidate.fornavn,
    surname: candidate.etternavn,
    jobTitle: candidate.stilling || undefined,
    department: candidate.avdeling || undefined,
  };

  const res = await graphRequest('POST', '/users', body);
  if (!res.ok) throw new Error(`Graph POST user-feil: ${res.status} ${await res.text()}`);
  const user = await res.json();
  return { user, tempPassword };
}

async function ensureLicense(userId) {
  if (!config.microsoft.licenseSkuId) {
    return { assigned: false, note: 'Ingen licenseSkuId konfigurert — se docs/MICROSOFT-ADMIN-SETUP.md' };
  }

  const userRes = await graphRequest('GET', `/users/${userId}?$select=assignedLicenses`);
  if (!userRes.ok) throw new Error(`Graph GET assignedLicenses-feil: ${userRes.status} ${await userRes.text()}`);
  const user = await userRes.json();
  const hasLicense = (user.assignedLicenses || []).some((l) => l.skuId === config.microsoft.licenseSkuId);
  if (hasLicense) return { assigned: true };

  const res = await graphRequest('POST', `/users/${userId}/assignLicense`, {
    addLicenses: [{ skuId: config.microsoft.licenseSkuId }],
    removeLicenses: [],
  });
  if (!res.ok) throw new Error(`Graph assignLicense-feil: ${res.status} ${await res.text()}`);
  return { assigned: true };
}

// Mints a fresh one-time password for an existing user. Used by the "velkommen" step, which
// may run well after (or be retried independently of) account creation — the original creation
// password is deliberately never persisted, so the welcome step needs a way to get a valid
// password at the moment it actually sends, not depend on a value from an earlier step call.
async function resetTempPassword(userId) {
  if (config.demoMode) return 'Demo1234!';

  const tempPassword = generateTempPassword();
  const res = await graphRequest('PATCH', `/users/${userId}`, {
    passwordProfile: {
      forceChangePasswordNextSignIn: true,
      password: tempPassword,
    },
  });
  if (!res.ok) throw new Error(`Graph reset-passord-feil: ${res.status} ${await res.text()}`);
  return tempPassword;
}

async function ensureGroups(userId, avdeling) {
  const groupIds = config.microsoft.groupIdsByAvdeling[avdeling] || [];
  if (!groupIds.length) {
    return { added: [], note: `Ingen grupper konfigurert for avdeling "${avdeling}"` };
  }

  const checkRes = await graphRequest('POST', `/users/${userId}/checkMemberGroups`, { groupIds });
  if (!checkRes.ok) throw new Error(`Graph checkMemberGroups-feil: ${checkRes.status} ${await checkRes.text()}`);
  const { value: alreadyMember } = await checkRes.json();

  const added = [];
  for (const groupId of groupIds) {
    if (alreadyMember.includes(groupId)) continue;
    const res = await graphRequest('POST', `/groups/${groupId}/members/$ref`, {
      '@odata.id': `${GRAPH_BASE}/directoryObjects/${userId}`,
    });
    if (!res.ok) throw new Error(`Graph add-to-group-feil (${groupId}): ${res.status} ${await res.text()}`);
    added.push(groupId);
  }
  return { added };
}

// A newly created Azure AD user is not instantly visible to every Graph endpoint (eventual
// consistency): assignLicense and checkMemberGroups can return 404 "Request_ResourceNotFound" for
// a few seconds after POST /users. Retry those specific 404s a handful of times so the first
// automatic run (the DocuSign webhook) completes, instead of failing and needing a manual retry.
function isResourceNotFound(err) {
  const m = (err && err.message) || '';
  return / 404 /.test(m) || /Request_ResourceNotFound/i.test(m) || /does not exist/i.test(m);
}

async function withPropagationRetry(fn, attempts = 5, delayMs = 3000) {
  for (let i = 0; ; i += 1) {
    try {
      return await fn();
    } catch (e) {
      if (i >= attempts - 1 || !isResourceNotFound(e)) throw e;
      await new Promise((r) => setTimeout(r, delayMs));
    }
  }
}

async function ensure(candidate, ctx) {
  if (config.demoMode) {
    return {
      ok: true,
      externalId: `DEMO-MS-${ctx.kandidatId}`,
      details: {
        upn: buildUpn(candidate),
        tempPassword: 'Demo1234!',
        licenseAssigned: true,
        groupsAdded: [],
      },
      demoMode: true,
    };
  }

  try {
    // Prefer the UPN decided at registration — that exact address is written into the signed
    // employment contract and sent to Telenor, so the account must match it rather than the
    // other way round. buildUpn is only a fallback for rows created before this was stored.
    const upn = candidate.microsoftUpn || buildUpn(candidate);
    let user = await findUserByUpn(upn);
    let tempPassword = null;

    if (!user) {
      const created = await createUser(candidate, upn);
      user = created.user;
      tempPassword = created.tempPassword;
    }

    // Wrapped in propagation-retry: on a just-created account these can 404 until Azure AD catches
    // up. For an already-existing user the first call succeeds and no retry happens.
    const license = await withPropagationRetry(() => ensureLicense(user.id));
    const groups = await withPropagationRetry(() => ensureGroups(user.id, candidate.avdeling));

    return {
      ok: true,
      externalId: user.id,
      details: {
        upn: user.userPrincipalName || upn,
        tempPassword, // only non-null the run that actually created the account — never persisted
        // True when the account already existed. Normally that just means this step is being
        // retried. It would also be true if a *different* person already holds this address
        // (two hires with the same name) — rare at Electi's size, but it would mean the new hire
        // is adopted onto someone else's mailbox, so it is surfaced in the log rather than hidden.
        alreadyExisted: !tempPassword,
        licenseAssigned: license.assigned,
        licenseNote: license.note,
        groupsAdded: groups.added,
        groupsNote: groups.note,
      },
    };
  } catch (e) {
    return { ok: false, error: e.message, retryable: true };
  }
}

module.exports = { ensure, buildUpn, resetTempPassword };
