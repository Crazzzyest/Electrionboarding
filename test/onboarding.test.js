process.env.DEMO_MODE = 'true'; // adapters short-circuit to demo, no external calls
const test = require('node:test');
const assert = require('node:assert');
const demoData = require('../src/demo-data');
const storage = require('../src/storage');
const onboarding = require('../src/onboarding');
const offboarding = require('../src/offboarding');

test.beforeEach(() => demoData.resetDemoState());

async function freshCandidate() {
  return storage.createCandidate({
    fornavn: 'Kari', etternavn: 'Testesen', privatEpost: 'kari@x.no', mobil: '99999999',
    stilling: 'Selger', stillingsprosent: '100', avdeling: 'Salg', naermesteLeder: 'Leder',
    registrertAv: 'Meg', microsoftUpn: 'kari.testesen@electi.no',
  });
}

test('runStep completes microsoft365 in demo, and is idempotent on re-run', async () => {
  const c = await freshCandidate();
  const first = await onboarding.runStep(c.row, 'microsoft365', {});
  assert.strictEqual(first.ok, true);
  assert.strictEqual((await storage.getCandidate(c.row)).statusMicrosoft365, 'OK');

  const second = await onboarding.runStep(c.row, 'microsoft365', {});
  assert.strictEqual(second.skipped, true, 'already-OK step must be skipped, not re-run');
  assert.strictEqual(second.ok, true);
});

test('runOnboardingSteps drives all steps to OK', async () => {
  const c = await freshCandidate();
  await onboarding.runOnboardingSteps(c.row, {});
  const got = await storage.getCandidate(c.row);
  for (const f of ['statusMicrosoft365', 'statusTelenor', 'statusSalesscreen', 'statusVelkommen']) {
    assert.strictEqual(got[f], 'OK', `${f} should be OK`);
  }
});

test('offboarding runs all steps to OK and stamps utfortDato', async () => {
  const o = await storage.createOffboarding({
    navn: 'Slutter Person', microsoftUpn: 's.p@electi.no', sluttdato: '2026-10-01', registrertAv: 'Meg', harProvisjon: true,
  });
  await offboarding.runOffboarding(o.row, {});
  const got = await storage.getOffboarding(o.row);
  for (const f of ['statusMicrosoft', 'statusTelenor', 'statusSalesscreen', 'statusProvisjon']) {
    assert.strictEqual(got[f], 'OK', `${f} should be OK`);
  }
  assert.ok(got.utfortDato, 'utfortDato should be set once every step is OK');
});

test('friendlyError maps known Graph errors to human text', () => {
  assert.match(
    onboarding.friendlyError('Graph POST /users/x@electri.no/sendMail feilet: 404 {"error":{"code":"ErrorInvalidUser"}}'),
    /EMAIL_SEND_AS/,
  );
  assert.match(onboarding.friendlyError('403 Authorization_RequestDenied'), /Microsoft-tillatelse/);
  assert.strictEqual(onboarding.friendlyError(''), 'Ukjent feil');
  assert.strictEqual(onboarding.friendlyError('Noe gikk galt {"x":1}'), 'Noe gikk galt'); // JSON tail stripped
});
