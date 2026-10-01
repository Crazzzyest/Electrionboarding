process.env.DEMO_MODE = 'true'; // no external calls
const test = require('node:test');
const assert = require('node:assert');
const { escapeHtml } = require('../src/utils');
const { csrfGuard, CSRF_VALUE } = require('../src/auth');
const microsoft = require('../src/microsoft');
const { validateCandidate } = require('../src/validation');
const { buildWelcomeEmailHtml } = require('../src/emails');

test('escapeHtml neutralises markup and attribute breakouts', () => {
  assert.strictEqual(
    escapeHtml(`<a href="x">'&`),
    '&lt;a href=&quot;x&quot;&gt;&#39;&amp;',
  );
  assert.strictEqual(escapeHtml(null), '');
});

test('welcome email does not render a name as HTML', () => {
  const html = buildWelcomeEmailHtml({ fornavn: '<a href="https://evil">Ola</a>', microsoftUpn: 'ola@electi.no' }, 'pw');
  assert.ok(!html.includes('<a href="https://evil">'));
  assert.ok(html.includes('&lt;a href='));
});

function runGuard(req) {
  let status = null;
  let passed = false;
  const res = { status(s) { status = s; return { json() {} }; } };
  csrfGuard({ get: (h) => (req.headers || {})[h.toLowerCase()], ...req }, res, () => { passed = true; });
  return { status, passed };
}

test('csrfGuard blocks state-changing /api calls without the header', () => {
  assert.deepStrictEqual(runGuard({ method: 'POST', path: '/api/candidates/5/resend-contract' }), { status: 403, passed: false });
  assert.deepStrictEqual(runGuard({ method: 'PUT', path: '/api/candidates/5' }), { status: 403, passed: false });
});

test('csrfGuard lets through GETs, the header, and non-api paths', () => {
  assert.strictEqual(runGuard({ method: 'GET', path: '/api/candidates' }).passed, true);
  assert.strictEqual(runGuard({ method: 'POST', path: '/api/candidates', headers: { 'x-requested-with': CSRF_VALUE } }).passed, true);
  assert.strictEqual(runGuard({ method: 'POST', path: '/webhooks/docusign' }).passed, true);
});

test('isOwnedBy only accepts an account stamped with the same kandidatId', () => {
  assert.strictEqual(microsoft.isOwnedBy({ employeeId: 'ONB-2026-007' }, 'ONB-2026-007'), true);
  assert.strictEqual(microsoft.isOwnedBy({ employeeId: 'ONB-2026-001' }, 'ONB-2026-007'), false);
  assert.strictEqual(microsoft.isOwnedBy({ employeeId: null }, 'ONB-2026-007'), false); // pre-existing employee
  assert.strictEqual(microsoft.isOwnedBy({ employeeId: undefined }, undefined), false);
});

test('allocateUpn skips addresses already held by other candidates', async () => {
  const c = { fornavn: 'Ola', etternavn: 'Nordmann' };
  assert.strictEqual(await microsoft.allocateUpn(c, []), 'ola.nordmann@electi.no');
  assert.strictEqual(await microsoft.allocateUpn(c, ['OLA.NORDMANN@electi.no']), 'ola.nordmann2@electi.no');
  assert.strictEqual(
    await microsoft.allocateUpn(c, ['ola.nordmann@electi.no', 'ola.nordmann2@electi.no']),
    'ola.nordmann3@electi.no',
  );
});

test('names without letters are rejected (would give an empty address)', () => {
  assert.deepStrictEqual(validateCandidate({ fornavn: '<>!', etternavn: 'Nordmann' }), ['Fornavn må inneholde bokstaver.']);
  assert.deepStrictEqual(validateCandidate({ fornavn: 'Øyvind', etternavn: 'Å' }), []);
});
