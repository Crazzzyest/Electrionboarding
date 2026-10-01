const crypto = require('crypto');

function generateKandidatId(year, seq) {
  return `ONB-${year}-${String(seq).padStart(3, '0')}`;
}

function generateOffboardingId(year, seq) {
  return `OFB-${year}-${String(seq).padStart(3, '0')}`;
}

// crypto.randomInt, not Math.random: this password guards a real Microsoft 365 account until the
// employee changes it, so it must not be predictable. Fisher-Yates shuffle (also crypto-backed) so
// the fixed-category prefix isn't predictable either.
function generateTempPassword() {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghijkmnpqrstuvwxyz';
  const digits = '23456789';
  const symbols = '!@#$%&*';
  const pick = (chars) => chars[crypto.randomInt(chars.length)];

  const chars = [pick(upper), pick(lower), pick(digits), pick(symbols)];
  const all = upper + lower + digits + symbols;
  for (let i = 0; i < 10; i++) chars.push(pick(all));

  for (let i = chars.length - 1; i > 0; i--) {
    const j = crypto.randomInt(i + 1);
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join('');
}

// Escape a value for interpolation into HTML (emails and anything rendered as HTML). Every
// user-supplied field (names, stilling, "registrert av", error messages echoed from external APIs)
// goes through this, so a name like <a href=...> can't inject links or markup into mail sent from
// Electi's own mailbox.
function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Excel (via Graph) returns date-typed cells as a serial NUMBER (e.g. 46272), not the "YYYY-MM-DD"
// string the form wrote — Excel coerces a date-looking string into a date cell on write. This turns
// such a serial back into an ISO date; a value that is not a bare number is returned unchanged (so a
// real "1992-05-14" string, or an empty cell, passes straight through). Excel's epoch is 1899-12-30
// (which also absorbs the 1900 leap-year bug for any date at/after 1900-03-01 — every date we handle).
function excelSerialToISO(value) {
  const s = String(value == null ? '' : value).trim();
  if (!/^\d{2,6}$/.test(s)) return s;
  const ms = Date.UTC(1899, 11, 30) + Number(s) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}

// "2026-09-07" -> "07.09.2026" (Norwegian display). Anything not ISO-shaped is returned unchanged.
function formatDateNo(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || '').trim());
  return m ? `${m[3]}.${m[2]}.${m[1]}` : String(value || '').trim();
}

// Norwegian-aware slug for building UPNs/mail nicknames from names (Ole Bjørn -> olebjorn).
// Explicit character map rather than Unicode-range regex/NFD tricks, which are easy to get
// subtly wrong with invisible combining characters — this is easy to read and verify instead.
const TRANSLIT_MAP = {
  æ: 'ae', ø: 'o', å: 'a',
  é: 'e', è: 'e', ê: 'e', ë: 'e',
  á: 'a', à: 'a', â: 'a', ä: 'a',
  ö: 'o', ó: 'o', ò: 'o', ô: 'o',
  ü: 'u', ú: 'u', ù: 'u', û: 'u',
  ñ: 'n', ç: 'c', ý: 'y', ß: 'ss',
};

function slugifyName(name) {
  return String(name || '')
    .toLowerCase()
    .split('')
    .map((ch) => TRANSLIT_MAP[ch] || ch)
    .join('')
    .replace(/[^a-z0-9]+/g, '');
}

// Returns { year, month, day } for "today" in a given IANA timezone, independent of the
// server/container's own system timezone.
function todayInTimezone(timeZone) {
  const fmt = new Intl.DateTimeFormat('en-CA', { timeZone, year: 'numeric', month: '2-digit', day: '2-digit' });
  const parts = fmt.formatToParts(new Date());
  const get = (type) => Number(parts.find((p) => p.type === type).value);
  return { year: get('year'), month: get('month'), day: get('day') };
}

// Compares an ISO "YYYY-MM-DD..." date string's month/day against a { month, day } object,
// without going through Date's local-time getters (avoids timezone-shift bugs on the stored date).
function isSameMonthDay(isoDateStr, { month, day }) {
  if (!isoDateStr) return false;
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(isoDateStr);
  if (!match) return false;
  return Number(match[2]) === month && Number(match[3]) === day;
}

module.exports = {
  generateKandidatId,
  generateOffboardingId,
  excelSerialToISO,
  formatDateNo,
  generateTempPassword,
  escapeHtml,
  slugifyName,
  todayInTimezone,
  isSameMonthDay,
};
