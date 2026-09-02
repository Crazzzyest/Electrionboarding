// Column indices (1-based, matching the spreadsheet) for the "Ansatte" tab.
const COL = {
  KANDIDAT_ID: 1,
  FORNAVN: 2,
  ETTERNAVN: 3,
  FODSELSDATO: 4,
  PRIVAT_EPOST: 5,
  MOBIL: 6,
  ADRESSE: 7,
  STILLING: 8,
  STILLINGSPROSENT: 9,
  AVDELING: 10,
  NAERMESTE_LEDER: 11,
  STARTDATO: 12,
  REGISTRERT_AV: 13,
  REGISTRERT_DATO: 14,
  STATUS_KONTRAKT: 15,
  DOCUSIGN_ENVELOPE_ID: 16,
  KONTRAKT_SENDT_DATO: 17,
  KONTRAKT_SIGNERT_DATO: 18,
  STATUS_MICROSOFT365: 19,
  MICROSOFT_USER_ID: 20,
  MICROSOFT_UPN: 21,
  MICROSOFT365_FULLFORT_DATO: 22,
  STATUS_TELENOR: 23,
  TELENOR_BESTILLING_SENDT_DATO: 24,
  STATUS_SALESSCREEN: 25,
  SALESSCREEN_USER_ID: 26,
  SALESSCREEN_FULLFORT_DATO: 27,
  STATUS_VELKOMMEN: 28,
  VELKOMMEN_SENDT_DATO: 29,
  SISTE_BURSDAGSVARSEL_AR: 30,
  NOTATER: 31,
  SISTE_FEILMELDING: 32,
};

const NUM_COLS = 32;

// Column headers, in order — written as row 1 when the sheet is first set up.
const HEADERS = [
  'KandidatID', 'Fornavn', 'Etternavn', 'Fødselsdato', 'Privat e-post', 'Mobil', 'Adresse',
  'Stilling', 'Stillingsprosent', 'Avdeling', 'Nærmeste leder', 'Startdato', 'Registrert av', 'Registrert dato',
  'Status kontrakt', 'DocuSign EnvelopeID', 'Kontrakt sendt dato', 'Kontrakt signert dato',
  'Status Microsoft365', 'Microsoft UserID', 'Microsoft UPN', 'Microsoft365 fullført dato',
  'Status Telenor', 'Telenor bestilling sendt dato',
  'Status SalesScreen', 'SalesScreen UserID', 'SalesScreen fullført dato',
  'Status Velkommen', 'Velkommen sendt dato',
  'Siste bursdagsvarsel år', 'Notater', 'Siste feilmelding',
];

const STEG_STATUS = {
  VENTER: 'Venter',
  KJORER: 'Kjører',
  OK: 'OK',
  FEILET: 'Feilet',
};

const KONTRAKT_STATUS = {
  IKKE_SENDT: 'Ikke sendt',
  SENDT: 'Sendt',
  SIGNERT: 'Signert',
  AVSLATT: 'Avslått',
  FEILET: 'Feilet',
};

// "Logg" tab columns, fixed order, append-only (no COL map needed — never updated by index).
const LOGG_HEADERS = ['Tidspunkt', 'KandidatID', 'Steg', 'Handling', 'Melding', 'Kilde'];

const LOGG_HANDLING = {
  FORSOKT: 'forsøkt',
  FULLFORT: 'fullført',
  FEILET: 'feilet',
  HOPPET_OVER: 'hoppet_over_allerede_ok',
};

const LOGG_KILDE = {
  WEBHOOK: 'webhook',
  MANUELL: 'manuell',
  CRON: 'cron',
  REGISTRERING: 'registrering',
};

module.exports = {
  COL, NUM_COLS, HEADERS,
  STEG_STATUS, KONTRAKT_STATUS,
  LOGG_HEADERS, LOGG_HANDLING, LOGG_KILDE,
};
