const test = require('node:test');
const assert = require('node:assert');
const { validateCandidate } = require('../src/validation');

test('a well-formed candidate has no validation errors', () => {
  assert.deepStrictEqual(validateCandidate({
    privatEpost: 'ola@example.no', mobil: '98765432', kontonummer: '12345678901', stillingsprosent: '100',
  }), []);
});

test('kontonummer must be 11 digits', () => {
  assert.ok(validateCandidate({ kontonummer: '123' }).some((e) => /Kontonummer/.test(e)));
  assert.deepStrictEqual(validateCandidate({ kontonummer: '1234 56 78901' }), []); // spaces ignored, 11 digits
});

test('invalid e-post and short mobil are rejected', () => {
  assert.ok(validateCandidate({ privatEpost: 'ikke-epost' }).some((e) => /e-post/.test(e)));
  assert.ok(validateCandidate({ mobil: '123' }).some((e) => /Mobil/.test(e)));
});

test('stillingsprosent must be 1-100', () => {
  assert.ok(validateCandidate({ stillingsprosent: '0' }).length);
  assert.ok(validateCandidate({ stillingsprosent: '150' }).length);
  assert.deepStrictEqual(validateCandidate({ stillingsprosent: '50' }), []);
});

test('absent fields are not validated (partial edits)', () => {
  assert.deepStrictEqual(validateCandidate({}), []);
  assert.deepStrictEqual(validateCandidate({ stilling: 'Selger' }), []);
});
