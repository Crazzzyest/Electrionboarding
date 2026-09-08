// Archives the fully signed contract PDF to SharePoint when an envelope completes. Uses the same
// Graph app-only credentials as the Excel storage and mail (one identity, Files.ReadWrite.All), and
// the same SharePoint site as the workbook. Called from the DocuSign webhook's "completed" branch.
const config = require('./config');
const docusign = require('./docusign');
const { getAccessToken, GRAPH_BASE } = require('./graph-client');

// Strip characters SharePoint/OneDrive disallow in file names, keep it human-readable.
function safeName(s) {
  return String(s || '').replace(/[\\/:*?"<>|#%]/g, '').replace(/\s+/g, ' ').trim();
}

function itemBase() {
  if (!config.excel.siteId) throw new Error('SharePoint siteId ikke satt — kan ikke arkivere kontrakt');
  return `${GRAPH_BASE}/sites/${config.excel.siteId}/drive`;
}

// Create the contracts folder if it doesn't exist yet (conflictBehavior: fail => 409 if present,
// which we treat as success). The parent (/Onboarding) already holds the workbook.
async function ensureFolder(token) {
  const folder = config.contracts.folder.replace(/^\//, ''); // "Onboarding/Kontrakter"
  const parts = folder.split('/');
  const name = parts.pop();
  const parentPath = parts.join('/'); // "Onboarding"
  const parentRef = parentPath ? `root:/${encodeURI(parentPath)}:` : 'root';
  const res = await fetch(`${itemBase()}/${parentRef}/children`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
  });
  if (!res.ok && res.status !== 409) {
    throw new Error(`Kunne ikke opprette kontraktmappe: ${res.status} ${await res.text()}`);
  }
}

async function archiveSignedContract(candidate, envelopeId) {
  if (config.demoMode) return { ok: true, demoMode: true };

  const pdf = await docusign.downloadCompletedPdf(envelopeId);
  const fileName = `${candidate.kandidatId} - ${safeName(candidate.fornavn)} ${safeName(candidate.etternavn)}.pdf`;
  const path = `${config.contracts.folder}/${fileName}`;

  const token = await getAccessToken();
  await ensureFolder(token);

  const res = await fetch(`${itemBase()}/root:${encodeURI(path)}:/content`, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/pdf' },
    body: pdf,
  });
  if (!res.ok) throw new Error(`SharePoint-opplasting feilet: ${res.status} ${await res.text()}`);
  const item = await res.json();
  return { ok: true, webUrl: item.webUrl, fileName };
}

module.exports = { archiveSignedContract };
