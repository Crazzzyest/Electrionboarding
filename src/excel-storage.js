// Low-level Excel-via-Graph wrapper. Deliberately mirrors google.js's exact interface
// (getSheetData/appendRow/updateCells, same row-numbering convention: header = row 1, first
// data row = row 2) so storage.js needed a one-line import change and nothing else — this file
// is the actual swap boundary the plan called for, not storage.js.
//
// One real API difference from Sheets worth knowing: Excel Tables have no per-cell PATCH.
// updateCells() has to read the row, merge the requested columns into it, and PATCH the whole
// row back — see updateCells() below. At this app's write volume (one onboarding step at a
// time, never concurrent writes to the same row) that read-before-write isn't a correctness
// risk, just an extra round trip.
const config = require('./config');
const { graphJson } = require('./graph-client');

function itemBasePath() {
  if (config.excel.siteId) return `/sites/${config.excel.siteId}/drive/root:${config.excel.itemPath}:`;
  if (config.excel.driveId) return `/drives/${config.excel.driveId}/root:${config.excel.itemPath}:`;
  throw new Error('Excel-lagring er ikke konfigurert (mangler siteId/driveId + itemPath) — se docs/SETUP-CHECKLIST.md');
}

function tablePath(tableName) {
  return `${itemBasePath()}/workbook/tables/${encodeURIComponent(tableName)}`;
}

// Data rows only, in table order — the Tables API never includes the header row here (that's
// separate table metadata), unlike Sheets' values.get which includes row 1.
async function getTableRowsRaw(tableName) {
  const data = await graphJson('GET', `${tablePath(tableName)}/rows?$select=values`);
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
  await graphJson('POST', `${tablePath(tableName)}/rows/add`, { values: [rowValues] });
}

async function updateCell(tableName, row, col, value) {
  await updateCells(tableName, row, [{ col, value }]);
}

async function updateCells(tableName, row, updates) {
  if (!updates.length) return;
  const tableRowIndex = row - 2; // Sheets-style absolute row (header=1) -> 0-based table row index
  const rows = await getTableRowsRaw(tableName);
  const current = rows[tableRowIndex];
  if (!current) throw new Error(`Fant ikke rad ${row} i tabellen "${tableName}"`);

  const merged = [...current];
  for (const { col, value } of updates) merged[col - 1] = value;

  await graphJson('PATCH', `${tablePath(tableName)}/rows/itemAt(index=${tableRowIndex})`, { values: [merged] });
}

module.exports = { getSheetData, appendRow, updateCell, updateCells };
