process.env.DEMO_MODE = 'true';
const test = require('node:test');
const assert = require('node:assert');
const demoData = require('../src/demo-data');
const storage = require('../src/storage');
const monitoring = require('../src/monitoring');

test.beforeEach(() => demoData.resetDemoState());

test('checkAlerts reports nothing when no step has failed', async () => {
  const r = await monitoring.checkAlerts();
  assert.strictEqual(r.alerts, 0);
});

test('checkAlerts detects a failed step', async () => {
  const c = await storage.createCandidate({
    fornavn: 'Feil', etternavn: 'Person', privatEpost: 'f@x.no', mobil: '99999999',
    stilling: 'Selger', stillingsprosent: '100', avdeling: 'Salg', naermesteLeder: 'L', registrertAv: 'M',
  });
  await storage.updateCandidateFields(c.row, { statusTelenor: 'Feilet', sisteFeilmelding: 'Noe gikk galt' });
  const r = await monitoring.checkAlerts();
  assert.ok(r.alerts >= 1, 'should flag the candidate with a failed step');
  assert.strictEqual(r.demo, true); // demo mode: computed but not emailed
});
