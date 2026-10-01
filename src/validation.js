// Format validation for candidate data — runs on both POST (new) and PUT (edit). Presence of
// required fields is checked separately in index.js; these only validate FORMAT, and only when a
// value is actually provided, so partial edits validate just the fields being changed.
const { slugifyName } = require('./utils');

function digitsOnly(s) {
  return String(s == null ? '' : s).replace(/\D/g, '');
}

function validateCandidate(body) {
  const errors = [];
  const has = (k) => body[k] !== undefined && body[k] !== null && String(body[k]).trim() !== '';

  // The @electi.no address is built from the name, so each name part must contain at least one
  // letter or digit that survives slugifying (a name of only symbols would give ".@electi.no").
  for (const [k, label] of [['fornavn', 'Fornavn'], ['etternavn', 'Etternavn']]) {
    if (has(k) && !slugifyName(body[k])) errors.push(`${label} må inneholde bokstaver.`);
  }

  if (has('privatEpost') && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(body.privatEpost).trim())) {
    errors.push('Ugyldig privat e-post.');
  }
  if (has('mobil') && digitsOnly(body.mobil).length < 8) {
    errors.push('Mobil må ha minst 8 sifre.');
  }
  if (has('kontonummer') && digitsOnly(body.kontonummer).length !== 11) {
    errors.push('Kontonummer må være 11 sifre.');
  }
  if (has('stillingsprosent')) {
    const p = Number(body.stillingsprosent);
    if (!Number.isFinite(p) || p < 1 || p > 100) errors.push('Stillingsprosent må være mellom 1 og 100.');
  }
  return errors;
}

module.exports = { validateCandidate };
