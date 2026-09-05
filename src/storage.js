// Domain-level candidate storage. This is the one place that knows whether data lives in the
// Excel workbook or in the in-memory DEMO_MODE store — onboarding.js and the adapters only ever
// see plain candidate objects, never a table row or a demo-state reference. excel-storage.js is
// the actual swap boundary (it mirrors this file's expected interface exactly), so switching
// backends again later would only mean writing a new module with the same shape.
const config = require('./config');
// Storage backend for non-demo runs: 'excel' (production, SharePoint/OneDrive via Graph) or
// 'local' (permission-free JSON file for demos / early Sliplane deploys). demoMode still bypasses
// both with the in-memory store in the functions below. Both modules share one interface, so the
// only thing that changes here is which one is required.
const backend = config.storage.backend === 'local'
  ? require('./local-storage')
  : require('./excel-storage');
const demoData = require('./demo-data');
const {
  COL, NUM_COLS, STEG_STATUS, KONTRAKT_STATUS, OFF_COL, OFF_NUM_COLS,
} = require('./columns');
const { generateKandidatId, generateOffboardingId } = require('./utils');

const FIELD_TO_COL = {
  statusKontrakt: COL.STATUS_KONTRAKT,
  docusignEnvelopeId: COL.DOCUSIGN_ENVELOPE_ID,
  kontraktSendtDato: COL.KONTRAKT_SENDT_DATO,
  kontraktSignertDato: COL.KONTRAKT_SIGNERT_DATO,
  statusMicrosoft365: COL.STATUS_MICROSOFT365,
  microsoftUserId: COL.MICROSOFT_USER_ID,
  microsoftUpn: COL.MICROSOFT_UPN,
  microsoft365FullfortDato: COL.MICROSOFT365_FULLFORT_DATO,
  statusTelenor: COL.STATUS_TELENOR,
  telenorBestillingSendtDato: COL.TELENOR_BESTILLING_SENDT_DATO,
  statusSalesscreen: COL.STATUS_SALESSCREEN,
  salesscreenUserId: COL.SALESSCREEN_USER_ID,
  salesscreenFullfortDato: COL.SALESSCREEN_FULLFORT_DATO,
  statusVelkommen: COL.STATUS_VELKOMMEN,
  velkommenSendtDato: COL.VELKOMMEN_SENDT_DATO,
  sisteBursdagsvarselAr: COL.SISTE_BURSDAGSVARSEL_AR,
  notater: COL.NOTATER,
  sisteFeilmelding: COL.SISTE_FEILMELDING,
  kontonummer: COL.KONTONUMMER,
};

function rowToCandidate(values, rowNumber) {
  const get = (col) => values[col - 1] || '';
  return {
    row: rowNumber,
    kandidatId: get(COL.KANDIDAT_ID),
    fornavn: get(COL.FORNAVN),
    etternavn: get(COL.ETTERNAVN),
    fodselsdato: get(COL.FODSELSDATO),
    privatEpost: get(COL.PRIVAT_EPOST),
    mobil: get(COL.MOBIL),
    adresse: get(COL.ADRESSE),
    stilling: get(COL.STILLING),
    stillingsprosent: get(COL.STILLINGSPROSENT),
    avdeling: get(COL.AVDELING),
    naermesteLeder: get(COL.NAERMESTE_LEDER),
    startdato: get(COL.STARTDATO),
    registrertAv: get(COL.REGISTRERT_AV),
    registrertDato: get(COL.REGISTRERT_DATO),
    statusKontrakt: get(COL.STATUS_KONTRAKT) || KONTRAKT_STATUS.IKKE_SENDT,
    docusignEnvelopeId: get(COL.DOCUSIGN_ENVELOPE_ID),
    kontraktSendtDato: get(COL.KONTRAKT_SENDT_DATO),
    kontraktSignertDato: get(COL.KONTRAKT_SIGNERT_DATO),
    statusMicrosoft365: get(COL.STATUS_MICROSOFT365) || STEG_STATUS.VENTER,
    microsoftUserId: get(COL.MICROSOFT_USER_ID),
    microsoftUpn: get(COL.MICROSOFT_UPN),
    microsoft365FullfortDato: get(COL.MICROSOFT365_FULLFORT_DATO),
    statusTelenor: get(COL.STATUS_TELENOR) || STEG_STATUS.VENTER,
    telenorBestillingSendtDato: get(COL.TELENOR_BESTILLING_SENDT_DATO),
    statusSalesscreen: get(COL.STATUS_SALESSCREEN) || STEG_STATUS.VENTER,
    salesscreenUserId: get(COL.SALESSCREEN_USER_ID),
    salesscreenFullfortDato: get(COL.SALESSCREEN_FULLFORT_DATO),
    statusVelkommen: get(COL.STATUS_VELKOMMEN) || STEG_STATUS.VENTER,
    velkommenSendtDato: get(COL.VELKOMMEN_SENDT_DATO),
    sisteBursdagsvarselAr: get(COL.SISTE_BURSDAGSVARSEL_AR),
    notater: get(COL.NOTATER),
    sisteFeilmelding: get(COL.SISTE_FEILMELDING),
    kontonummer: get(COL.KONTONUMMER),
  };
}

async function listCandidates() {
  if (config.demoMode) {
    return demoData.state.candidates.map((c) => ({ ...c }));
  }
  const rows = await backend.getSheetData(config.excel.ansatteTable);
  return rows
    .slice(1) // header row
    .map((values, i) => rowToCandidate(values, i + 2))
    .filter((c) => c.kandidatId);
}

async function getCandidate(row) {
  const all = await listCandidates();
  return all.find((c) => c.row === Number(row)) || null;
}

async function findCandidateByEnvelopeId(envelopeId) {
  const all = await listCandidates();
  return all.find((c) => c.docusignEnvelopeId === envelopeId) || null;
}

async function createCandidate(fields) {
  const all = await listCandidates();
  const year = new Date().getFullYear();
  const countThisYear = all.filter((c) => c.kandidatId.startsWith(`ONB-${year}-`)).length;
  const kandidatId = generateKandidatId(year, countThisYear + 1);
  const now = new Date().toISOString();

  const candidate = {
    kandidatId,
    fornavn: fields.fornavn,
    etternavn: fields.etternavn,
    fodselsdato: fields.fodselsdato,
    privatEpost: fields.privatEpost,
    mobil: fields.mobil,
    adresse: fields.adresse || '',
    stilling: fields.stilling,
    stillingsprosent: fields.stillingsprosent,
    avdeling: fields.avdeling,
    naermesteLeder: fields.naermesteLeder,
    startdato: fields.startdato || '',
    registrertAv: fields.registrertAv,
    registrertDato: now,
    statusKontrakt: KONTRAKT_STATUS.IKKE_SENDT,
    docusignEnvelopeId: '',
    kontraktSendtDato: '',
    kontraktSignertDato: '',
    statusMicrosoft365: STEG_STATUS.VENTER,
    microsoftUserId: '',
    // Set at registration, not when the account is created: the employment contract itself states
    // the new hire's @electi.no address, so it has to be decided before the contract goes out.
    // It is deterministic (firstname.lastname@domain), so the Microsoft step later creates exactly
    // this address rather than inventing its own. See microsoft.js.
    microsoftUpn: fields.microsoftUpn || '',
    microsoft365FullfortDato: '',
    statusTelenor: STEG_STATUS.VENTER,
    telenorBestillingSendtDato: '',
    statusSalesscreen: STEG_STATUS.VENTER,
    salesscreenUserId: '',
    salesscreenFullfortDato: '',
    statusVelkommen: STEG_STATUS.VENTER,
    velkommenSendtDato: '',
    sisteBursdagsvarselAr: '',
    notater: '',
    sisteFeilmelding: '',
    kontonummer: fields.kontonummer || '',
  };

  if (config.demoMode) {
    const row = Math.max(1, ...demoData.state.candidates.map((c) => c.row)) + 1;
    demoData.state.candidates.push({ ...candidate, row });
    return { ...candidate, row };
  }

  const row = new Array(NUM_COLS).fill('');
  row[COL.KANDIDAT_ID - 1] = candidate.kandidatId;
  row[COL.FORNAVN - 1] = candidate.fornavn;
  row[COL.ETTERNAVN - 1] = candidate.etternavn;
  row[COL.FODSELSDATO - 1] = candidate.fodselsdato;
  row[COL.PRIVAT_EPOST - 1] = candidate.privatEpost;
  row[COL.MOBIL - 1] = candidate.mobil;
  row[COL.ADRESSE - 1] = candidate.adresse;
  row[COL.STILLING - 1] = candidate.stilling;
  row[COL.STILLINGSPROSENT - 1] = candidate.stillingsprosent;
  row[COL.AVDELING - 1] = candidate.avdeling;
  row[COL.NAERMESTE_LEDER - 1] = candidate.naermesteLeder;
  row[COL.STARTDATO - 1] = candidate.startdato;
  row[COL.REGISTRERT_AV - 1] = candidate.registrertAv;
  row[COL.REGISTRERT_DATO - 1] = candidate.registrertDato;
  row[COL.STATUS_KONTRAKT - 1] = candidate.statusKontrakt;
  row[COL.MICROSOFT_UPN - 1] = candidate.microsoftUpn;
  row[COL.STATUS_MICROSOFT365 - 1] = candidate.statusMicrosoft365;
  row[COL.STATUS_TELENOR - 1] = candidate.statusTelenor;
  row[COL.STATUS_SALESSCREEN - 1] = candidate.statusSalesscreen;
  row[COL.STATUS_VELKOMMEN - 1] = candidate.statusVelkommen;
  row[COL.KONTONUMMER - 1] = candidate.kontonummer;

  await backend.appendRow(config.excel.ansatteTable, row);

  const refreshed = await listCandidates();
  return refreshed.find((c) => c.kandidatId === kandidatId);
}

async function updateCandidateFields(row, fields) {
  if (config.demoMode) {
    const candidate = demoData.state.candidates.find((c) => c.row === Number(row));
    if (candidate) Object.assign(candidate, fields);
    return;
  }

  const updates = Object.entries(fields)
    .filter(([key]) => FIELD_TO_COL[key])
    .map(([key, value]) => ({ col: FIELD_TO_COL[key], value }));

  if (!updates.length) return;
  await backend.updateCells(config.excel.ansatteTable, row, updates);
}

async function appendLog(kandidatId, steg, handling, melding, kilde) {
  const entry = {
    tidspunkt: new Date().toISOString(),
    kandidatId,
    steg,
    handling,
    melding: melding || '',
    kilde,
  };

  if (config.demoMode) {
    demoData.state.log.unshift(entry);
    return;
  }

  await backend.appendRow(config.excel.loggTable, [
    entry.tidspunkt, entry.kandidatId, entry.steg, entry.handling, entry.melding, entry.kilde,
  ]);
}

async function listLog(kandidatId) {
  if (config.demoMode) {
    return demoData.state.log.filter((e) => e.kandidatId === kandidatId);
  }
  const rows = await backend.getSheetData(config.excel.loggTable);
  return rows
    .slice(1)
    .map(([tidspunkt, id, steg, handling, melding, kilde]) => ({ tidspunkt, kandidatId: id, steg, handling, melding, kilde }))
    .filter((e) => e.kandidatId === kandidatId)
    .reverse();
}

// ---- Offboarding (separate table, same backend) ----------------------------------------------

const OFF_FIELD_TO_COL = {
  statusMicrosoft: OFF_COL.STATUS_MICROSOFT,
  statusTelenor: OFF_COL.STATUS_TELENOR,
  statusSalesscreen: OFF_COL.STATUS_SALESSCREEN,
  statusProvisjon: OFF_COL.STATUS_PROVISJON,
  utfortDato: OFF_COL.UTFORT_DATO,
  sisteFeilmelding: OFF_COL.SISTE_FEILMELDING,
};

function rowToOffboarding(values, rowNumber) {
  const get = (col) => values[col - 1] || '';
  return {
    row: rowNumber,
    offboardingId: get(OFF_COL.OFFBOARDING_ID),
    navn: get(OFF_COL.NAVN),
    microsoftUpn: get(OFF_COL.MICROSOFT_UPN),
    sluttdato: get(OFF_COL.SLUTTDATO),
    harProvisjon: get(OFF_COL.HAR_PROVISJON),
    registrertAv: get(OFF_COL.REGISTRERT_AV),
    registrertDato: get(OFF_COL.REGISTRERT_DATO),
    statusMicrosoft: get(OFF_COL.STATUS_MICROSOFT) || STEG_STATUS.VENTER,
    statusTelenor: get(OFF_COL.STATUS_TELENOR) || STEG_STATUS.VENTER,
    statusSalesscreen: get(OFF_COL.STATUS_SALESSCREEN) || STEG_STATUS.VENTER,
    statusProvisjon: get(OFF_COL.STATUS_PROVISJON) || STEG_STATUS.VENTER,
    utfortDato: get(OFF_COL.UTFORT_DATO),
    sisteFeilmelding: get(OFF_COL.SISTE_FEILMELDING),
  };
}

async function listOffboardings() {
  if (config.demoMode) {
    return demoData.state.offboardings.map((o) => ({ ...o }));
  }
  const rows = await backend.getSheetData(config.excel.offboardingTable);
  return rows
    .slice(1)
    .map((values, i) => rowToOffboarding(values, i + 2))
    .filter((o) => o.offboardingId);
}

async function getOffboarding(row) {
  const all = await listOffboardings();
  return all.find((o) => o.row === Number(row)) || null;
}

async function createOffboarding(fields) {
  const all = await listOffboardings();
  const year = new Date().getFullYear();
  const countThisYear = all.filter((o) => o.offboardingId.startsWith(`OFB-${year}-`)).length;
  const offboardingId = generateOffboardingId(year, countThisYear + 1);
  const now = new Date().toISOString();

  const offboarding = {
    offboardingId,
    navn: fields.navn,
    microsoftUpn: fields.microsoftUpn || '',
    sluttdato: fields.sluttdato || '',
    // Stored as a plain "Ja"/"Nei" so the sheet is human-readable.
    harProvisjon: fields.harProvisjon ? 'Ja' : 'Nei',
    registrertAv: fields.registrertAv || '',
    registrertDato: now,
    statusMicrosoft: STEG_STATUS.VENTER,
    statusTelenor: STEG_STATUS.VENTER,
    statusSalesscreen: STEG_STATUS.VENTER,
    // If no commission is owed, that step is a no-op from the start — mark it OK so the row can
    // reach "fully done" without a manual skip.
    statusProvisjon: fields.harProvisjon ? STEG_STATUS.VENTER : STEG_STATUS.OK,
    utfortDato: '',
    sisteFeilmelding: '',
  };

  if (config.demoMode) {
    const row = Math.max(1, ...demoData.state.offboardings.map((o) => o.row), 1) + 1;
    demoData.state.offboardings.push({ ...offboarding, row });
    return { ...offboarding, row };
  }

  const row = new Array(OFF_NUM_COLS).fill('');
  row[OFF_COL.OFFBOARDING_ID - 1] = offboarding.offboardingId;
  row[OFF_COL.NAVN - 1] = offboarding.navn;
  row[OFF_COL.MICROSOFT_UPN - 1] = offboarding.microsoftUpn;
  row[OFF_COL.SLUTTDATO - 1] = offboarding.sluttdato;
  row[OFF_COL.HAR_PROVISJON - 1] = offboarding.harProvisjon;
  row[OFF_COL.REGISTRERT_AV - 1] = offboarding.registrertAv;
  row[OFF_COL.REGISTRERT_DATO - 1] = offboarding.registrertDato;
  row[OFF_COL.STATUS_MICROSOFT - 1] = offboarding.statusMicrosoft;
  row[OFF_COL.STATUS_TELENOR - 1] = offboarding.statusTelenor;
  row[OFF_COL.STATUS_SALESSCREEN - 1] = offboarding.statusSalesscreen;
  row[OFF_COL.STATUS_PROVISJON - 1] = offboarding.statusProvisjon;

  await backend.appendRow(config.excel.offboardingTable, row);

  const refreshed = await listOffboardings();
  return refreshed.find((o) => o.offboardingId === offboardingId);
}

async function updateOffboardingFields(row, fields) {
  if (config.demoMode) {
    const o = demoData.state.offboardings.find((x) => x.row === Number(row));
    if (o) Object.assign(o, fields);
    return;
  }

  const updates = Object.entries(fields)
    .filter(([key]) => OFF_FIELD_TO_COL[key])
    .map(([key, value]) => ({ col: OFF_FIELD_TO_COL[key], value }));

  if (!updates.length) return;
  await backend.updateCells(config.excel.offboardingTable, row, updates);
}

module.exports = {
  listCandidates,
  getCandidate,
  findCandidateByEnvelopeId,
  createCandidate,
  updateCandidateFields,
  appendLog,
  listLog,
  listOffboardings,
  getOffboarding,
  createOffboarding,
  updateOffboardingFields,
};
