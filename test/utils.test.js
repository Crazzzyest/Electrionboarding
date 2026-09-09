const test = require('node:test');
const assert = require('node:assert');
const {
  generateKandidatId, generateOffboardingId, generateTempPassword,
  excelSerialToISO, formatDateNo, slugifyName, isSameMonthDay, todayInTimezone,
} = require('../src/utils');

test('generateKandidatId zero-pads the sequence', () => {
  assert.strictEqual(generateKandidatId(2026, 1), 'ONB-2026-001');
  assert.strictEqual(generateKandidatId(2026, 42), 'ONB-2026-042');
});

test('generateOffboardingId uses OFB prefix', () => {
  assert.strictEqual(generateOffboardingId(2026, 2), 'OFB-2026-002');
});

test('excelSerialToISO converts Excel date serials and passes through real dates', () => {
  assert.strictEqual(excelSerialToISO(46272), '2026-09-07'); // the bug we fixed
  assert.strictEqual(excelSerialToISO('46272'), '2026-09-07');
  assert.strictEqual(excelSerialToISO('1992-05-14'), '1992-05-14'); // already ISO -> unchanged
  assert.strictEqual(excelSerialToISO(''), ''); // empty -> empty
  assert.strictEqual(excelSerialToISO('  '), ''); // whitespace -> empty
});

test('formatDateNo renders ISO as dd.mm.yyyy and leaves other strings alone', () => {
  assert.strictEqual(formatDateNo('2026-09-07'), '07.09.2026');
  assert.strictEqual(formatDateNo(''), '');
  assert.strictEqual(formatDateNo('ukjent'), 'ukjent');
});

test('slugifyName transliterates Norwegian characters and strips separators', () => {
  assert.strictEqual(slugifyName('Ole Bjørn'), 'olebjorn');
  assert.strictEqual(slugifyName('Ærlig Åse'), 'aerligase');
  assert.strictEqual(slugifyName('Anne-Kari'), 'annekari');
});

test('generateTempPassword is 12 chars with mixed character classes', () => {
  for (let i = 0; i < 20; i += 1) {
    const p = generateTempPassword();
    assert.strictEqual(p.length, 12);
    assert.match(p, /[A-Z]/);
    assert.match(p, /[a-z]/);
    assert.match(p, /[0-9]/);
    assert.match(p, /[!@#$%&*]/);
  }
});

test('isSameMonthDay matches month/day ignoring year and timezone', () => {
  assert.ok(isSameMonthDay('1992-09-07', { month: 9, day: 7 }));
  assert.ok(!isSameMonthDay('1992-09-08', { month: 9, day: 7 }));
  assert.ok(!isSameMonthDay('', { month: 9, day: 7 }));
});

test('todayInTimezone returns numeric y/m/d', () => {
  const t = todayInTimezone('Europe/Oslo');
  assert.strictEqual(typeof t.year, 'number');
  assert.ok(t.month >= 1 && t.month <= 12);
  assert.ok(t.day >= 1 && t.day <= 31);
});
