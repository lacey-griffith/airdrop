/**
 * sharepoint.js
 * -----------------------------------------------------------------------------
 * Microsoft Graph helpers for SharePoint folder listing + file download,
 * and Excel parsing for Convert preview links.
 *
 * Requires application permissions:
 *   - Files.Read.All
 *   - Sites.Read.All
 * with admin consent on your App Registration.
 */

import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// Per docs: /shares/{encodedUrl} where encoded = base64(url) with URL-safe chars
function encodeSharingUrl(url) {
  const b64 = Buffer.from(url, 'utf8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
  return `u!${b64}`;
}

// OAuth2 Client Credentials → access token for Graph
export async function getGraphToken() {
  const { MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET } = process.env;
  if (!MS_TENANT_ID || !MS_CLIENT_ID || !MS_CLIENT_SECRET) return null;

  const url = `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: MS_CLIENT_ID,
    client_secret: MS_CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });

  const res = await fetch(url, { method: 'POST', body });
  if (!res.ok) return null;

  const json = await res.json();
  return json.access_token || null;
}

/**
 * List children in a shared folder link.
 * Normalizes a subset of fields we care about.
 * Includes driveId so we can download with /drives/{driveId}/items/{id}/content
 */
export async function listFolderItems({ folderUrl, token }) {
  if (!token) return [];

  const encoded = encodeSharingUrl(folderUrl);
  const res = await fetch(`${GRAPH_BASE}/shares/${encoded}/driveItem/children?$top=999`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph list failed: ${res.status}`);

  const json = await res.json();
  const items = (json.value || []).map((it) => ({
    id: it.id,
    driveId: it.parentReference?.driveId || '',
    name: it.name,
    size: it.size,
    mime: it.file?.mimeType || '',
    webUrl: it.webUrl,
    isFolder: !!it.folder,
  }));
  return items;
}

/**
 * Download a file's bytes using the driveId + itemId form.
 * (Works for application-permission scenarios across sites/drives.)
 */
export async function downloadItemBuffer({ driveId, itemId, token }) {
  if (!token) return null;
  if (!driveId || !itemId) throw new Error('downloadItemBuffer missing driveId or itemId');

  const res = await fetch(`${GRAPH_BASE}/drives/${driveId}/items/${itemId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Graph download failed: ${res.status}`);

  const buf = Buffer.from(await res.arrayBuffer());
  return buf;
}

/**
 * Excel selection strategy based on the ClickUp task's title.
 * Order:
 *   1) exact "<taskTitle>.xlsx|xls"
 *   2) startsWith "<taskTitle> "
 *   3) contains "preview"
 *   4) any .xlsx (fallback)
 */
export function findExcelForTask(items, taskTitle) {
  const isExcel = (n = '') => /\.xlsx?$/i.test(n);

  const norm = (s = '') =>
    s
      .toLowerCase()
      .replace(/[\u2013\u2014]/g, '-') // en/em dash → hyphen
      .replace(/[_\s]+/g, ' ')         // collapse underscores/whitespace
      .replace(/[^\w.\- ]+/g, '')      // keep letters/numbers/._-
      .trim();

  const base = norm(taskTitle);

  // exact match
  const exact = items.find(
    (f) => !f.isFolder && isExcel(f.name) && norm(f.name).replace(/\.xlsx?$/, '') === base,
  );
  if (exact) return exact;

  // starts with "<title> "
  const starts = items.find(
    (f) => !f.isFolder && isExcel(f.name) && norm(f.name).startsWith(base + ' '),
  );
  if (starts) return starts;

  // contains "preview"
  const preview = items.find((f) => !f.isFolder && isExcel(f.name) && /preview/i.test(f.name));
  if (preview) return preview;

  // any excel
  return items.find((f) => !f.isFolder && isExcel(f.name)) || null;
}

/**
 * Return image items in the folder (png/jpg/webp/gif).
 * Accepts a custom extension regex.
 */
export function getImageFiles(items, extRegex = /\.(png|jpe?g|webp|gif)$/i) {
  return items.filter((f) => !f.isFolder && extRegex.test(f.name || ''));
}

/**
 * Extract Convert preview URLs from all cells in all sheets (pattern-based).
 * We flatten to CSV per sheet for a quick text scan.
 */
export function extractPreviewLinksFromXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const links = new Set();

  // keep the same pattern as in post-qa.js
  const RE = /\bhttps?:\/\/[^\s)]+?(?:convert_action=convert_vpreview|convert_e=\d{6,}|convert_v=\d{6,})[^\s)]*/gi;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws, { FS: '\t' });
    (csv.match(RE) || []).forEach((u) => links.add(u.trim()));
  }
  return [...links];
}
