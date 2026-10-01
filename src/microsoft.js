// Adapter: Microsoft Graph (app-only / client-credentials — Electi's own Azure AD tenant).
// Does idempotent sub-checks internally so a partially-failed prior run never redoes completed
// work: find-or-create user, assign license if missing, add to groups if missing. Never touches
// storage — onboarding.js owns all Status_*/timestamp writes (see the adapter contract there).
const config = require('./config');
const { slugifyName, generateTempPassword } = require('./utils');
const { GRAPH_BASE, graphRequest } = require('./graph-client');

function buildUpn(candidate, suffix = '') {
  const local = `${slugifyName(candidate.fornavn)}.${slugifyName(candidate.etternavn)}${suffix}`;
  return `${local}@${config.microsoft.domain}`;
}

// Every account this app creates is stamped with the candidate's kandidatId in Graph's employeeId
// field. That stamp is the ONLY proof an existing account is ours: without it, registering a
// "new hire" with the same name as a current employee would adopt that employee's account and the
// welcome step would reset their password and mail it to whatever private address was typed in.
function isOwnedBy(user, kandidatId) {
  return Boolean(user && kandidatId && user.employeeId === kandidatId);
}

async function findUserByUpn(upn) {
  const res = await graphRequest('GET', `/users/${encodeURIComponent(upn)}?$select=id,userPrincipalName,assignedLicenses,employeeId`);
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Graph GET user-feil: ${res.status} ${await res.text()}`);
  return res.json();
}

// Any user already using the address, as sign-in name, primary mail or alias (proxyAddresses) —
// a UPN that is free can still collide with someone's alias, and Exchange would reject it.
async function findUsersByAddress(address) {
  const a = address.replace(/'/g, "''");
  const filter = `userPrincipalName eq '${a}' or mail eq '${a}' or proxyAddresses/any(p:p eq 'smtp:${a}')`;
  const res = await graphRequest('GET', `/users?$filter=${encodeURIComponent(filter)}&$select=id,userPrincipalName,employeeId`);
  if (!res.ok) throw new Error(`Graph adresse-oppslag feilet: ${res.status} ${await res.text()}`);
  return (await res.json()).value || [];
}

// Picks the first free @electi.no address for a candidate: ola.nordmann, then ola.nordmann2, 3, ...
// "Free" means neither another candidate in storage (takenUpns) nor any existing Microsoft account
// holds it — except an account this same candidate already owns (a retry after a partial failure).
async function allocateUpn(candidate, takenUpns = []) {
  const taken = new Set(takenUpns.map((u) => String(u || '').trim().toLowerCase()).filter(Boolean));
  for (let n = 1; n <= 50; n += 1) {
    const upn = buildUpn(candidate, n === 1 ? '' : String(n));
    if (taken.has(upn.toLowerCase())) continue;
    if (config.demoMode) return upn;
    const holders = await findUsersByAddress(upn);
    if (holders.every((u) => isOwnedBy(u, candidate.kandidatId))) return upn;
  }
  throw new Error('Fant ingen ledig @electi.no-adresse for dette navnet');
}

async function createUser(candidate, upn, kandidatId) {
  const tempPassword = generateTempPassword();
  const body = {
    accountEnabled: true,
    displayName: `${candidate.fornavn} ${candidate.etternavn}`,
    mailNickname: upn.split('@')[0],
    userPrincipalName: upn,
    employeeId: kandidatId, // ownership stamp — see isOwnedBy
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
async function resetTempPassword(userId, kandidatId) {
  if (config.demoMode) return 'Demo1234!';

  // Never reset a password on an account this app didn't create for this candidate — that would
  // hand a working employee's login to whoever's private e-mail is on the candidate row.
  const check = await graphRequest('GET', `/users/${userId}?$select=id,employeeId`);
  if (!check.ok) throw new Error(`Graph GET user-feil: ${check.status} ${await check.text()}`);
  if (!isOwnedBy(await check.json(), kandidatId)) {
    throw new Error('Kontoen ble ikke opprettet av onboarding for denne kandidaten — passordet tilbakestilles ikke.');
  }

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
  // Match the "Avdeling/Team" value case-insensitively and trimmed, so "Salg", "salg" and " SALG "
  // all map to the same configured group — the exact-case match was an easy footgun in the form.
  const map = config.microsoft.groupIdsByAvdeling;
  const norm = String(avdeling || '').trim().toLowerCase();
  const key = Object.keys(map).find((k) => k.trim().toLowerCase() === norm);
  const groupIds = (key && map[key]) || [];
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

// Offboarding guard. The offboarding form takes a free-text address, and disabling/deleting is
// destructive — so only sellers may be offboarded: an account this app created (employeeId stamp),
// or a member of one of the configured seller groups (covers sellers hired before this app — they
// must be added to the group; checked 2026-09-28 the Salg group held only 3 members, and the seller
// licence is no substitute: managers hold it too). Anyone
// else (management, admins, shared mailboxes) is refused. A missing account passes — there is
// nothing to disable, and offboardUser already treats that as done.
async function checkOffboardable(upn) {
  if (config.demoMode) return { ok: true };
  const user = await findUserByUpn(upn);
  if (!user) return { ok: true, notFound: true };
  if (/^ONB-/.test(user.employeeId || '')) return { ok: true };

  const groupIds = [...new Set(Object.values(config.microsoft.groupIdsByAvdeling).flat())];
  if (groupIds.length) {
    const res = await graphRequest('POST', `/users/${user.id}/checkMemberGroups`, { groupIds });
    if (!res.ok) throw new Error(`Graph checkMemberGroups-feil: ${res.status} ${await res.text()}`);
    const { value } = await res.json();
    if ((value || []).length) return { ok: true };
  }
  return {
    ok: false,
    reason: `${upn} er ikke opprettet av onboarding og er ikke med i Salg-gruppen, så offboarding er avvist. `
      + 'Er personen selger, legg vedkommende i Salg-gruppen i Microsoft 365 og prøv igjen.',
  };
}

// Offboarding: block or remove a departing employee's Microsoft 365 account. 'disable' sets
// accountEnabled=false and strips the licence (frees the seat) while keeping the mailbox for
// retention; 'delete' removes the user entirely. Idempotent — a not-found user, or an account
// already disabled/unlicensed, is treated as success rather than an error, so a retry is safe.
async function offboardUser(upn, action = 'disable') {
  const user = await findUserByUpn(upn);
  if (!user) {
    return { ok: true, details: { notFound: true, note: `Fant ingen konto for ${upn} (allerede fjernet?)` } };
  }

  if (action === 'delete') {
    const res = await graphRequest('DELETE', `/users/${user.id}`);
    if (!res.ok && res.status !== 404) throw new Error(`Graph DELETE user-feil: ${res.status} ${await res.text()}`);
    return { ok: true, externalId: user.id, details: { action: 'delete' } };
  }

  // disable: block sign-in
  const patch = await graphRequest('PATCH', `/users/${user.id}`, { accountEnabled: false });
  if (!patch.ok) throw new Error(`Graph disable-konto-feil: ${patch.status} ${await patch.text()}`);

  // free the licence seat (only if one is assigned — removing an absent licence errors)
  const removed = [];
  for (const lic of user.assignedLicenses || []) {
    if (!lic.skuId) continue;
    const r = await graphRequest('POST', `/users/${user.id}/assignLicense`, {
      addLicenses: [], removeLicenses: [lic.skuId],
    });
    if (!r.ok) throw new Error(`Graph removeLicense-feil: ${r.status} ${await r.text()}`);
    removed.push(lic.skuId);
  }

  return { ok: true, externalId: user.id, details: { action: 'disable', licensesRemoved: removed } };
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
      const created = await createUser(candidate, upn, ctx.kandidatId);
      user = created.user;
      tempPassword = created.tempPassword;
    } else if (!isOwnedBy(user, ctx.kandidatId)) {
      // Someone else already has this address (a current employee with the same name, or an
      // account created outside this app). Refuse rather than adopt it — adopting would give the
      // new hire that person's mailbox and let the welcome step reset their password.
      return {
        ok: false,
        error: `Adressen ${upn} tilhører allerede en annen konto. Rediger kandidatens navn for å få en ny adresse.`,
        retryable: false,
      };
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
        // True when the account already existed — only possible for an account stamped with this
        // candidate's kandidatId, i.e. a retry after a partial failure (foreign accounts are refused).
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

module.exports = {
  ensure, buildUpn, allocateUpn, isOwnedBy, resetTempPassword, checkOffboardable, offboardUser,
};
