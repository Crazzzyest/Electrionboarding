// Low-level Excel-via-Graph wrapper. Deliberately mirrors google.js's exact interface
// (getSheetData/appendRow/updateCells, same row-numbering convention: header = row 1, first
// data row = row 2) so storage.js needed a one-line import change and nothing else — this file
// is the actual swap boundary the plan called for, not storage.js.
//
// One real API difference from Sheets worth knowing: Excel Tables have no per-cell PATCH.
// updateCells() has to read the row, merge the requested columns into it, and PATCH the whole
// row back — see updateCells() below. Because onboarding/offboarding steps now run in parallel,
// several updateCells() calls CAN hit the same row at once, and read-merge-write would then lose
// updates (each reads the same stale row and PATCHes back the whole thing, last writer wins). A
// per-row async lock (withRowLock) serialises those so each re-reads after the previous PATCH.
const config = require('./config');
const { graphJson } = require('./graph-client');

// Per-key promise chain: withRowLock(key, fn) runs fn only after the previous fn for the same key
// has settled, so concurrent read-merge-write cycles on one row queue instead of clobbering. The
// stored tail is error-swallowed so one failure doesn't wedge the chain; the caller still gets the
// real result/error.
const rowChains = new Map();
function withRowLock(key, fn) {
  const prev = rowChains.get(key) || Promise.resolve();
  const next = prev.then(fn, fn);
  rowChains.set(key, next.catch(() => {}));
  return next;
}

function itemBasePath() {
  if (config.excel.siteId) return `/sites/${config.excel.siteId}/drive/root:${config.excel.itemPath}:`;
  if (config.excel.driveId) return `/drives/${config.excel.driveId}/root:${config.excel.itemPath}:`;
  throw new Error('Excel-lagring er ikke konfigurert (mangler siteId/driveId + itemPath) — se docs/SETUP-CHECKLIST.md');
}

function tablePath(tableName) {
  return `${itemBasePath()}/workbook/tables/${encodeURIComponent(tableName)}`;
}

// A persistent workbook session is the documented fix for EditModeCannotAcquireLockTooManyRequests:
// without one, Graph opens (and locks) a fresh session per write, so parallel steps writing to the
// same workbook collide on the edit lock. With persistChanges:true, all our reads/writes share one
// session and one lock. The session id is cached; a single in-flight createSession promise prevents
// parallel callers from opening several. Graph expires idle sessions (~7 min), so on a session error
// we drop the cache and retry once with a fresh one.
let sessionPromise = null;
function getSessionId() {
  if (!sessionPromise) {
    sessionPromise = graphJson('POST', `${itemBasePath()}/workbook/createSession`, { persistChanges: true })
      .then((d) => d.id)
      .catch((e) => { sessionPromise = null; throw e; });
  }
  return sessionPromise;
}

function isSessionError(err) {
  const m = (err && err.message) || '';
  return /session/i.test(m) || / 404 /.test(m) || /InvalidSession|invalidSessionId|expired/i.test(m);
}

// All workbook calls go through this so they carry the session id (and recover if it expired).
async function wb(method, path, body) {
  const sid = await getSessionId();
  try {
    return await graphJson(method, path, body, { 'workbook-session-id': sid });
  } catch (e) {
    if (!isSessionError(e)) throw e;
    sessionPromise = null; // force a fresh session and retry once
    const sid2 = await getSessionId();
    return graphJson(method, path, body, { 'workbook-session-id': sid2 });
  }
}

// Data rows only, in table order — the Tables API never includes the header row here (that's
// separate table metadata), unlike Sheets' values.get which includes row 1.
async function getTableRowsRaw(tableName) {
  const data = await wb('GET', `${tablePath(tableName)}/rows?$select=values`);
  // Each row's `values` is a single-row 2D array ([[cell1, cell2, ...]]) even for one row —
  // a Range-API quirk carried over into the Tables API.
  return (data.value || []).map((r) => r.values[0]);
}

// Returns [header, ...dataRows] like google.js's getSheetData, so callers that do `.slice(1)`
// keep working unchanged. The header row's actual content doesn't matter to any caller (only
// its presence, to keep the index math lined up) so it's a blank placeholder, not a real read.
async function getSheetData(tableName) {
  const rows = await getTableRowsRaw(tableName);
  const width = rows[0] ? rows[0].length : 0;
  return [new Array(width).fill(''), ...rows];
}

async function appendRow(tableName, rowValues) {
  await wb('POST', `${tablePath(tableName)}/rows/add`, { values: [rowValues] });
}

async function updateCell(tableName, row, col, value) {
  await updateCells(tableName, row, [{ col, value }]);
}

async function updateCells(tableName, row, updates) {
  if (!updates.length) return;
  // Serialise read-merge-write per row so parallel steps don't clobber each other's columns.
  await withRowLock(`${tableName}#${row}`, async () => {
    const tableRowIndex = row - 2; // Sheets-style absolute row (header=1) -> 0-based table row index
    const rows = await getTableRowsRaw(tableName);
    const current = rows[tableRowIndex];
    if (!current) throw new Error(`Fant ikke rad ${row} i tabellen "${tableName}"`);

    const merged = [...current];
    for (const { col, value } of updates) merged[col - 1] = value;

    await wb('PATCH', `${tablePath(tableName)}/rows/itemAt(index=${tableRowIndex})`, { values: [merged] });
  });
}

module.exports = { getSheetData, appendRow, updateCell, updateCells };
