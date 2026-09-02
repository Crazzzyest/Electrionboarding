// Mail-via-Graph wrapper, replacing Gmail for the three places this app sends email: Telenor
// order, welcome email, birthday notification. App-only Mail.Send always sends as a specific
// mailbox (there's no "service account with no inbox" concept), so every call goes through
// config.email.sendAsMailbox — see the config.js comment on why that has to be a real mailbox.
const config = require('./config');
const { graphJson } = require('./graph-client');

function mailboxPath() {
  if (!config.email.sendAsMailbox) {
    throw new Error('E-post er ikke konfigurert (mangler sendAsMailbox) — se docs/SETUP-CHECKLIST.md');
  }
  return `/users/${encodeURIComponent(config.email.sendAsMailbox)}`;
}

// `to` and `cc` accept a single address or an array, matching the old google.js signature.
async function sendEmail(to, subject, html, { cc } = {}) {
  const toList = Array.isArray(to) ? to : [to];
  const message = {
    subject,
    body: { contentType: 'HTML', content: html },
    toRecipients: toList.map((address) => ({ emailAddress: { address } })),
  };
  if (cc && cc.length) {
    message.ccRecipients = (Array.isArray(cc) ? cc : [cc]).map((address) => ({ emailAddress: { address } }));
  }

  await graphJson('POST', `${mailboxPath()}/sendMail`, { message, saveToSentItems: true });
}

// Used only for the Telenor-order idempotency check (see telenor.js) — Graph's unqualified
// $search on /messages covers subject and body, matching what the old Gmail `q:` search did for
// the same purpose. Requires Mail.Read (or Mail.ReadWrite) in addition to Mail.Send.
async function searchMail(query) {
  const data = await graphJson('GET', `${mailboxPath()}/messages?$search=${encodeURIComponent(`"${query}"`)}&$top=1`);
  return (data.value || []).length > 0;
}

module.exports = { sendEmail, searchMail };
