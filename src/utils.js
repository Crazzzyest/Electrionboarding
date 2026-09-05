function generateKandidatId(year, seq) {
  return `ONB-${year}-${String(seq).padStart(3, '0')}`;
}

function generateOffboardingId(year, seq) {
  return `OFB-${year}-${String(seq).padStart(3, '0')}`;
}

function generateTempPassword() {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghijkmnpqrstuvwxyz';
  const digits = '23456789';
  const symbols = '!@#$%&*';
  const pick = (chars) => chars[Math.floor(Math.random() * chars.length)];

  let pwd = pick(upper) + pick(lower) + pick(digits) + pick(symbols);
  const all = upper + lower + digits + symbols;
  for (let i = 0; i < 8; i++) pwd += pick(all);

  // Shuffle so the fixed-category prefix isn't predictable.
  return pwd.split('').sort(() => Math.random() - 0.5).join('');
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
  generateTempPassword,
  slugifyName,
  todayInTimezone,
  isSameMonthDay,
};
