// Low-level local-file storage. A permission-free stand-in for excel-storage.js, for demos and
// early deploys made BEFORE the SharePoint/OneDrive Excel file (and its Graph Files permission)
// are set up. Mirrors excel-storage.js's exact interface (getSheetData/appendRow/updateCell/
// updateCells; header = row 1, first data row = row 2), so storage.js only picks a different
// backend module and changes nothing else.
//
// Data lives in one JSON file: { "<tableName>": [ [cell, cell, ...], ... ] } holding DATA rows
// only (no header row), same row model as excel-storage's getTableRowsRaw.
//
// NOT durable on ephemeral hosts: a Sliplane container resets this file on redeploy/restart. It
// is a bridge to get "form -> contract" live without waiting on Microsoft tenant permissions,
// not the production store. Excel in SharePoint/OneDrive remains the target (see config.storage).
const fs = require('fs');
const path = require('path');
const config = require('./config');

const FILE = config.storage.localFile;

function readAll() {
  try {
    return JSON.parse(fs.readFileSync(FILE, 'utf8'));
  } catch {
    return {}; // missing/empty/corrupt file -> start clean
  }
}

function writeAll(data) {
  fs.mkdirSync(path.dirname(FILE), { recursive: true });
  fs.writeFileSync(FILE, JSON.stringify(data, null, 2), 'utf8');
}

function getRows(tableName) {
  const data = readAll();
  return Array.isArray(data[tableName]) ? data[tableName] : [];
}

// Returns [blankHeader, ...dataRows] like excel-storage.getSheetData, so callers that do .slice(1)
// keep working unchanged. The header's content is irrelevant to callers (only its presence, to
// keep the row-index math lined up), so it is a blank placeholder.
async function getSheetData(tableName) {
  const rows = getRows(tableName);
  const width = rows[0] ? rows[0].length : 0;
  return [new Array(width).fill(''), ...rows];
}

async function appendRow(tableName, rowValues) {
  const data = readAll();
  if (!Array.isArray(data[tableName])) data[tableName] = [];
  data[tableName].push(rowValues);
  writeAll(data);
}

async function updateCell(tableName, row, col, value) {
  await updateCells(tableName, row, [{ col, value }]);
}

async function updateCells(tableName, row, updates) {
  if (!updates.length) return;
  const data = readAll();
  const rows = Array.isArray(data[tableName]) ? data[tableName] : [];
  const idx = row - 2; // absolute row (header=1) -> 0-based data-row index
  const current = rows[idx];
  if (!current) throw new Error(`Fant ikke rad ${row} i tabellen "${tableName}"`);
  for (const { col, value } of updates) current[col - 1] = value;
  data[tableName] = rows;
  writeAll(data);
}

module.exports = { getSheetData, appendRow, updateCell, updateCells };
