// In-memory "fake sheet" used only when DEMO_MODE=true, so the full candidate lifecycle
// (register -> sign -> steps run -> retry) is interactively demoable with zero real
// credentials. storage.js reads/writes this instead of calling google.js when demoMode is on.
const { STEG_STATUS, KONTRAKT_STATUS } = require('./columns');
const { todayInTimezone } = require('./utils');

function isoNow() {
  return new Date().toISOString();
}

function buildSeedCandidates() {
  const today = todayInTimezone('Europe/Oslo');
  const bursdagIAr = `1992-${String(today.month).padStart(2, '0')}-${String(today.day).padStart(2, '0')}`;

  return [
    {
      row: 2,
      kandidatId: 'ONB-DEMO-001',
      fornavn: 'Kari',
      etternavn: 'Nordmann',
      fodselsdato: bursdagIAr,
      privatEpost: 'kari.demo@example.com',
      mobil: '900 00 001',
      adresse: 'Eksempelveien 1, 0001 Oslo',
      stilling: 'Selger',
      stillingsprosent: '100',
      avdeling: 'Salg',
      naermesteLeder: 'Demo Salgsleder, Salgssjef',
      startdato: '',
      registrertAv: 'Demo Salgsleder',
      registrertDato: isoNow(),
      statusKontrakt: KONTRAKT_STATUS.SIGNERT,
      docusignEnvelopeId: 'DEMO-ENVELOPE-001',
      kontraktSendtDato: isoNow(),
      kontraktSignertDato: isoNow(),
      statusMicrosoft365: STEG_STATUS.OK,
      microsoftUserId: 'DEMO-MS-001',
      microsoftUpn: 'kari.nordmann@electi.no',
      microsoft365FullfortDato: isoNow(),
      statusTelenor: STEG_STATUS.OK,
      telenorBestillingSendtDato: isoNow(),
      statusSalesscreen: STEG_STATUS.OK,
      salesscreenUserId: 'kari.demo@example.com',
      salesscreenFullfortDato: isoNow(),
      statusVelkommen: STEG_STATUS.OK,
      velkommenSendtDato: isoNow(),
      sisteBursdagsvarselAr: '',
      notater: 'Demo: fullført løp. Fødselsdato er satt til i dag for å vise bursdagsvarselet.',
      sisteFeilmelding: '',
    },
    {
      row: 3,
      kandidatId: 'ONB-DEMO-002',
      fornavn: 'Ola',
      etternavn: 'Eksempel',
      fodselsdato: '1995-03-14',
      privatEpost: 'ola.demo@example.com',
      mobil: '900 00 002',
      adresse: 'Testgata 2, 0002 Oslo',
      stilling: 'Selger',
      stillingsprosent: '100',
      avdeling: 'Salg',
      naermesteLeder: 'Demo Salgsleder, Salgssjef',
      startdato: '',
      registrertAv: 'Demo Salgsleder',
      registrertDato: isoNow(),
      statusKontrakt: KONTRAKT_STATUS.IKKE_SENDT,
      docusignEnvelopeId: '',
      kontraktSendtDato: '',
      kontraktSignertDato: '',
      statusMicrosoft365: STEG_STATUS.VENTER,
      microsoftUserId: '',
      microsoftUpn: '',
      microsoft365FullfortDato: '',
      statusTelenor: STEG_STATUS.VENTER,
      telenorBestillingSendtDato: '',
      statusSalesscreen: STEG_STATUS.VENTER,
      salesscreenUserId: '',
      salesscreenFullfortDato: '',
      statusVelkommen: STEG_STATUS.VENTER,
      velkommenSendtDato: '',
      sisteBursdagsvarselAr: '',
      notater: 'Demo: kontrakt akkurat "sendt", venter på simulert signering.',
      sisteFeilmelding: '',
    },
  ];
}

const state = {
  candidates: buildSeedCandidates(),
  offboardings: [],
  log: [],
  nextSeq: 3,
};

function resetDemoState() {
  state.candidates = buildSeedCandidates();
  state.offboardings = [];
  state.log = [];
  state.nextSeq = 3;
}

module.exports = { state, resetDemoState };
