const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '..', '.env'), quiet: true });

const config = {
  testMode: process.env.TEST_MODE === 'true',
  demoMode: process.env.DEMO_MODE === 'true',

  // Data store: an Excel workbook in SharePoint/OneDrive, via Microsoft Graph — chosen 2026-08-29
  // over Google Sheets specifically to avoid a second identity/auth story (Google OAuth) on top
  // of the Microsoft one this app already needs for user provisioning. Uses the SAME app-only
  // Graph credentials as everything else (see graph-client.js) — one credential, one consent,
  // one thing that can expire, not two.
  //
  // NOT CONFIRMED, all placeholders: Electi hasn't said where this file should live yet.
  // - siteId: the SharePoint site hosting the file (Graph: GET /sites/{hostname}:/{site-path}).
  //   Leave blank to address a personal OneDrive via driveId instead.
  // - driveId: only needed if NOT using siteId (e.g. a personal OneDrive rather than a
  //   SharePoint site's default document library).
  // - itemPath: the file's path within that site/drive's root, e.g. "/Onboarding/Electi-Onboarding.xlsx".
  // - ansatteTable / loggTable: the Excel Table (Insert -> Table in Excel, not just a cell range —
  //   Graph's row-add/row-update API needs a named Table, not a bare range) on each worksheet.
  // See docs/SETUP-CHECKLIST.md — excel-storage.js refuses to guess at any of this.
  excel: {
    // Verified live 2026-09-04: workbook created in the root SharePoint site's document library,
    // with real Tables "Ansatte" (32 cols) and "Logg" (6 cols). The Logg worksheet is named
    // "Loggark" because "Logg" is a reserved sheet name in Norwegian Excel — the TABLE is "Logg".
    siteId: 'electi.sharepoint.com,115eb03c-2edd-4386-9056-0888d93b4bbb,ae813d16-db52-48ae-a9c4-9153f46d721f',
    driveId: '',
    itemPath: '/Onboarding/Electi-Onboarding.xlsx',
    ansatteTable: 'Ansatte',
    loggTable: 'Logg',
    offboardingTable: 'Offboarding',
  },

  storage: {
    // Which storage backend to use when NOT in demo mode:
    //   'excel' — the production target: an Excel workbook in SharePoint/OneDrive via Graph
    //             (needs the Files permission + excel.siteId/driveId/itemPath above).
    //   'local' — a permission-free JSON file (see local-storage.js), for demos and early
    //             Sliplane deploys made before that permission/location exists. NOT durable on
    //             ephemeral hosts (resets on redeploy/restart) — a bridge, not production storage.
    backend: process.env.STORAGE_BACKEND || 'excel',
    // Where the 'local' backend keeps its JSON file. Overridable via env for a mounted volume.
    localFile: process.env.STORAGE_LOCAL_FILE
      || path.resolve(__dirname, '..', 'data', 'onboarding-data.json'),
  },

  email: {
    // Who receives birthday notifications. Overridable via env (MANAGEMENT_EMAIL) so Electi can
    // point it at real management later without a redeploy; defaults to Edson's address for now.
    managementEmail: process.env.MANAGEMENT_EMAIL || 'edson.reistad@effektivaltruisme.no',
    // TODO: which real mailbox automated mail is sent FROM (Telenor order, welcome, birthday).
    // Graph's app-only Mail.Send always sends as a specific mailbox — there's no "service account
    // with no inbox" option — so this has to be a real, licensed mailbox Electi is fine with
    // seeing as the sender. See docs/SETUP-CHECKLIST.md.
    sendAsMailbox: 'edson.reistad@electi.no', // verified live 2026-09-04: licensed, sends OK
  },

  microsoft: {
    tenantId: process.env.MICROSOFT_TENANT_ID,
    clientId: process.env.MICROSOFT_CLIENT_ID,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    // Verified 2026-08-25 against the live tenant: electi.no is the default verified domain
    // (electi.eu and electi.onmicrosoft.com also exist).
    domain: 'electi.no',
    usageLocation: 'NO',
    // EXCHANGEDESKLESS (Exchange Online Kiosk) — verified live: it is the only licence with a real
    // mailbox plan (EXCHANGE_S_DESKLESS) that has free seats, and its 41/45 usage matches Electi's
    // ~25+ seller headcount, so it is almost certainly the sellers' licence.
    // NOTE: the SKU the admin first supplied (1f2f344a…, "STREAM") grants NO mailbox at all —
    // see docs/MICROSOFT-ADMIN-SETUP.md. CONFIRM this choice with Electi before go-live.
    // WARNING: only 4 free seats remained on 2026-08-25.
    licenseSkuId: '80b2d799-d2ba-4d2a-8842-fb0d0f3a4b82',
    // Verified live: "Salg" is a real group in the tenant. (The Object ID the admin first supplied,
    // 457752a7…, does not resolve to any group or user in this tenant.)
    // Key must match the "Avdeling/Team" value typed in the registration form, exactly.
    // Other groups that exist, if Electi wants new hires in them too:
    //   e0fe9397-35a2-457a-8e9c-2b2273acf22c  All Company
    //   bab2c260-6036-4564-b48a-2a0a32350c17  Electi Business Partner AS
    groupIdsByAvdeling: {
      Salg: ['b5adad24-9fdc-47e2-95a4-a85084c0c08d'],
    },
  },

  docusign: {
    env: process.env.DOCUSIGN_ENV || 'demo', // 'demo' or 'production'
    integrationKey: process.env.DOCUSIGN_INTEGRATION_KEY,
    userId: process.env.DOCUSIGN_USER_ID,
    accountId: process.env.DOCUSIGN_ACCOUNT_ID,
    privateKeyBase64: process.env.DOCUSIGN_PRIVATE_KEY_BASE64,
    privateKeyPath: process.env.DOCUSIGN_PRIVATE_KEY_PATH
      || path.resolve(__dirname, '..', 'docusign-private-key.pem'),
    // Trimmed: a trailing newline/space from pasting the key into the host's env is the most
    // common cause of a Connect HMAC mismatch (401 on every webhook), and it is invisible.
    connectHmacKey: (process.env.DOCUSIGN_CONNECT_HMAC_KEY || '').trim(),
    templateId: 'c53d4f46-11ed-48ba-8800-053aa3afe989', // "Arbeidsavtale" template, rebuilt 2026-09-02
    signerRoleName: 'Arbeidstaker', // matches the rebuilt template's role name (was "Ansatt")
    // Discovered 2026-09-02 by reading the rebuilt template's real tab definitions via the API. The
    // template uses DocuSign's automatic field-recognition, so these are random internal labels, not
    // the "[[STILLING]]" strings from the Word template. Navn auto-fills from the recipient name
    // (FullName field, no tabLabel needed); everything else is a Text field the app fills below.
    // CAUTION: these labels belong to THIS processed version of the template. If the document is ever
    // re-uploaded/regenerated in DocuSign, they change and need re-discovery (run the discovery
    // script in scratchpad; see docs/SETUP-CHECKLIST.md).
    tabLabels: {
      epost: 'atb.docusignFields.label-text 2wyx5n7797jmtkb4rkj',
      telefon: 'atb.docusignFields.label-text ew53e0dgonmtkb63ts',
      stilling: 'atb.docusignFields.label-text mzd2ecr1skbmtkb6jfl',
      stillingsprosent: 'atb.docusignFields.label-text gb8jkpyohkmtkb6vsb',
      tiltredelsesdato: 'atb.docusignFields.label-text 57byde06i6fmtkb77ck',
      naermesteLeder: 'atb.docusignFields.label-text be441u2b9oomtkb7iun',
    },
  },

  salesscreen: {
    apiKey: process.env.SALESSCREEN_API_KEY,
    // Confirmed 2026-08-21 by reading their real Postman-published API spec (docs.salesscreen.com
    // renders via JS, so this needed a real browser, not a plain fetch) — POST /User/Add upserts
    // a user by "userId" (auto-creates the team too if the key doesn't exist yet). No lookup
    // endpoint exists, so idempotency relies entirely on that upsert behavior — see salesscreen.js.
    baseUrl: 'https://connect.salesscreen.com/api/v1',
    createUserEndpoint: '/User/Add',
    // New hires should land in Electi's leadership/manager team ("ledergruppen"). The SalesScreen
    // "connect" API is WRITE-ONLY — verified 2026-09-05: every list endpoint (/Team, /Team/List,
    // /Group, /User/List, …) returns 404, and SalesScreen's own API FAQ says team management must
    // be done in the platform, not via the API — so the exact team can't be discovered
    // programmatically. It lives in SalesScreen under Settings -> Mappings.
    //
    // CRITICAL: "key" is SalesScreen's match/auto-create identifier. POST /User/Add auto-CREATES a
    // team if the key doesn't already exist — so a wrong key silently makes a DUPLICATE team with
    // the same display name instead of putting the user in the real one. Read the real leadership
    // team's key from Settings -> Mappings and set SALESSCREEN_TEAM_KEY (+ _NAME) on the host. The
    // 'Telenor' fallback is only the old confirmed default from Magnus (2026-08-21); override it.
    team: {
      key: process.env.SALESSCREEN_TEAM_KEY || 'Telenor',
      name: process.env.SALESSCREEN_TEAM_NAME || process.env.SALESSCREEN_TEAM_KEY || 'Telenor',
    },
  },

  telenor: {
    // Modelled directly on a real order Magnus sent 2026-08-21 ("Ny selger"). That mail is far
    // simpler than the original proposal implied: subject "Ny selger", and a body listing only
    // name, phone and the new hire's @electi.no address. It does NOT mention Salesforce or Ransel
    // by name — see the open question in docs/SETUP-CHECKLIST.md.
    // Real recipients by default. For safe testing, redirect the order mail to a test inbox by
    // setting TELENOR_ORDER_TO (and optionally TELENOR_ORDER_CC), comma-separated, on the host — so
    // no real order ever reaches Telenor during a test signing. Remove those env vars to restore the
    // real recipients for go-live.
    orderTo: (process.env.TELENOR_ORDER_TO || 'c-view.admin@telenor.no,kasper-maehlum.kvesetberget@telenor.no')
      .split(',').map((s) => s.trim()).filter(Boolean),
    orderCc: (process.env.TELENOR_ORDER_CC || 'audun.paulsen@electi.no,daniel.kolstad@electi.no')
      .split(',').map((s) => s.trim()).filter(Boolean),
    subject: 'Ny selger',
  },

  // Offboarding (scope-addition after V1). Every policy choice is an env var so Electi can set them
  // in Sliplane without a code change. Recipients fall back to onboarding equivalents where sensible.
  offboarding: {
    // 'immediate' runs the steps the moment an offboarding is registered; 'scheduled' waits until
    // the sluttdato — a daily cron (07:00 Europe/Oslo) then runs the ones that have come due.
    timing: (process.env.OFFBOARDING_TIMING || 'immediate').trim().toLowerCase(),
    // What to do with the Microsoft 365 account: 'disable' blocks sign-in AND frees the licence
    // seat but keeps the mailbox (retention-friendly); 'delete' removes the user entirely.
    microsoftAction: (process.env.OFFBOARDING_MICROSOFT_ACTION || 'disable').trim().toLowerCase(),
    // Where the commission-handling notice goes when a departing seller has provisjonskrav.
    provisjonEpost: (process.env.OFFBOARDING_PROVISJON_EPOST || '').trim(),
    // SalesScreen's connect API is write-only with no confirmed deactivation endpoint, so that step
    // is a human handoff: a reminder mail to this address. Defaults to the provisjon recipient.
    salesscreenVarselEpost: (process.env.OFFBOARDING_SALESSCREEN_EPOST
      || process.env.OFFBOARDING_PROVISJON_EPOST || '').trim(),
    // Telenor cancellation mail. Recipients default to the same order contacts as onboarding.
    telenorSubject: process.env.OFFBOARDING_TELENOR_SUBJECT || 'Avslutning selger',
    telenorTo: (process.env.OFFBOARDING_TELENOR_TO || '').split(',').map((s) => s.trim()).filter(Boolean),
    telenorCc: (process.env.OFFBOARDING_TELENOR_CC || '').split(',').map((s) => s.trim()).filter(Boolean),
  },

  port: parseInt(process.env.PORT, 10) || 3000,
};

module.exports = config;
