process.env.DEMO_MODE = 'true'; // in-memory store, no Graph calls
const test = require('node:test');
const assert = require('node:assert');
const demoData = require('../src/demo-data');
const storage = require('../src/storage');

test.beforeEach(() => demoData.resetDemoState());

test('createCandidate assigns an ONB id and is listed', async () => {
  const c = await storage.createCandidate({
    fornavn: 'Test', etternavn: 'Person', privatEpost: 't@x.no', mobil: '99999999',
    stilling: 'Selger', stillingsprosent: '100', avdeling: 'Salg', naermesteLeder: 'Leder', registrertAv: 'Meg',
    kontonummer: '12345678901', microsoftUpn: 'test.person@electi.no',
  });
  assert.match(c.kandidatId, /^ONB-\d{4}-\d{3}$/);
  assert.strictEqual(c.kontonummer, '12345678901');
  const all = await storage.listCandidates();
  assert.ok(all.some((x) => x.kandidatId === c.kandidatId));
});

test('updateCandidateFields persists a change', async () => {
  const c = await storage.createCandidate({
    fornavn: 'A', etternavn: 'B', privatEpost: 'a@b.no', mobil: '1', stilling: 'S',
    stillingsprosent: '100', avdeling: 'Salg', naermesteLeder: 'L', registrertAv: 'M',
  });
  await storage.updateCandidateFields(c.row, { statusMicrosoft365: 'OK', microsoftUpn: 'a.b@electi.no' });
  const got = await storage.getCandidate(c.row);
  assert.strictEqual(got.statusMicrosoft365, 'OK');
  assert.strictEqual(got.microsoftUpn, 'a.b@electi.no');
});

test('createOffboarding with commission leaves provisjon pending; without, pre-set OK', async () => {
  const withProv = await storage.createOffboarding({
    navn: 'Med Provisjon', microsoftUpn: 'm.p@electi.no', sluttdato: '2026-10-01', registrertAv: 'Meg', harProvisjon: true,
  });
  assert.strictEqual(withProv.harProvisjon, 'Ja');
  assert.strictEqual(withProv.statusProvisjon, 'Venter');

  const noProv = await storage.createOffboarding({
    navn: 'Uten Provisjon', microsoftUpn: 'u.p@electi.no', sluttdato: '2026-10-05', registrertAv: 'Meg', harProvisjon: false,
  });
  assert.strictEqual(noProv.harProvisjon, 'Nei');
  assert.strictEqual(noProv.statusProvisjon, 'OK'); // no paperwork needed -> already done
});
