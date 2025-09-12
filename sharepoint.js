// sharepoint.js
// ----------------------------------------------------------------------------
// Microsoft Graph helpers for reading a SharePoint folder, optionally drilling
// into a single child subfolder named like the ClickUp task. Also extracts
// preview links from an .xlsx and filters image files.
//
// Dependencies (already in your package.json):
//   "node-fetch": "^3", "xlsx": "^0.18"
//
// Environment (for private sites):
//   MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
//
// Exports:
//   getGraphToken()
//   listFolderItems({ folderUrl, token })
//   listFolderChildren({ driveId, itemId, token })
//   findTaskSubfolder(items, taskTitle)
//   findExcelForTask(items, taskTitle)
//   getImageFiles(items, imageRegex)
//   downloadItemBuffer({ driveId, itemId, token })
//   extractPreviewLinksFromXlsx(buffer)

import fetch from 'node-fetch';
import * as XLSX from 'xlsx';

// -----------------------------
// Small utils
// -----------------------------
const normalize = (s = '') =>
  String(s).toLowerCase().replace(/\s+/g, ' ').replace(/[^\w\s]/g, '').trim();

const isFile = (it) => !!it?.file;
const isFolder = (it) => !!it?.folder;

function urlToGraphShareId(url) {
  // Graph expects the URL-safe base64 of the full URL, prefixed with "u!"
  // Replace '+' -> '-', '/' -> '_', strip '='
  const b64 = Buffer.from(String(url)).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  return `u!${b64}`;
}

// -----------------------------
// Auth
// -----------------------------
export async function getGraphToken() {
  const tenant = process.env.MS_TENANT_ID;
  const client = process.env.MS_CLIENT_ID;
  const secret = process.env.MS_CLIENT_SECRET;
  if (!tenant || !client || !secret) return null;

  const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: client,
    client_secret: secret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });

  const res = await fetch(url, { method: 'POST', body });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Graph token failed: ${res.status} ${txt}`);
  }
  const json = await res.json();
  return json.access_token;
}

// -----------------------------
// Folder → driveItem helpers
// -----------------------------
async function getDriveItemFromFolderUrl(folderUrl, token) {
  // Robust way: /shares/{shareId}/driveItem
  const shareId = urlToGraphShareId(folderUrl);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`shares/driveItem failed: ${res.status} ${txt}`);
  }
  const json = await res.json();
  // json has .parentReference.driveId and .id
  return json;
}

export async function listFolderItems({ folderUrl, token }) {
  const di = await getDriveItemFromFolderUrl(folderUrl, token);
  const driveId = di?.parentReference?.driveId || di?.parentReference?.id;
  const itemId = di?.id;
  if (!driveId || !itemId) throw new Error('Could not resolve driveId/itemId from folder URL');

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=999`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`children list failed: ${res.status} ${txt}`);
  }
  const json = await res.json();
  // Normalize items to include driveId for later calls
  const items = (json.value || []).map(i => ({
    ...i,
    driveId,
  }));
  return items;
}

export async function listFolderChildren({ driveId, itemId, token }) {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/children?$top=999`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`children list failed: ${res.status} ${txt}`);
  }
  const json = await res.json();
  return (json.value || []).map(i => ({ ...i, driveId }));
}

// -----------------------------
// Selection helpers
// -----------------------------
export function findTaskSubfolder(items, taskTitle) {
  // Look for a single child folder that best matches the task title
  const want = normalize(taskTitle);
  const folders = items.filter(isFolder);

  // 1) exact normalized match
  let hit = folders.find(f => normalize(f.name) === want);
  if (hit) return hit;

  // 2) startsWith (common pattern: "<Task Title> - v2")
  hit = folders.find(f => normalize(f.name).startsWith(`${want} `));
  if (hit) return hit;

  // 3) contains (as last resort)
  return folders.find(f => normalize(f.name).includes(want)) || null;
}

export function findExcelForTask(items, taskTitle) {
  const want = normalize(taskTitle);

  const excels = items.filter(it =>
    isFile(it) &&
    /\.xlsx?$/i.test(it.name || '')
  );

  if (!excels.length) return null;

  // Strip extension helper
  const base = (n) => normalize(n.replace(/\.(xlsx?|XLSX?)$/, ''));

  // 1) exact match "<Task Title>.xlsx"
  let hit = excels.find(f => base(f.name) === want);
  if (hit) return hit;

  // 2) startsWith "<Task Title> "
  hit = excels.find(f => base(f.name).startsWith(`${want} `));
  if (hit) return hit;

  // 3) filename contains "preview"
  hit = excels.find(f => /preview/i.test(f.name));
  if (hit) return hit;

  // 4) first xlsx as last resort
  return excels[0];
}

export function getImageFiles(items, imageRegex = /\.(png|jpe?g|webp|gif)$/i) {
  return items.filter(
    (it) =>
      isFile(it) &&
      (imageRegex.test(it.name || '') ||
        String(it?.file?.mimeType || '').toLowerCase().startsWith('image/'))
  );
}

// -----------------------------
// Download
// -----------------------------
export async function downloadItemBuffer({ driveId, itemId, token }) {
  // Use /content to stream bytes
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`download failed: ${res.status} ${txt}`);
  }
  const arr = await res.arrayBuffer();
  return Buffer.from(arr);
}

// -----------------------------
// XLSX: extract preview links
// -----------------------------
export function extractPreviewLinksFromXlsx(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });

  // Find "Preview Links" sheet (case-insensitive, trimmed)
  const sheetName =
    (wb.SheetNames || []).find(
      (n) => normalize(n) === normalize('Preview Links')
    ) || wb.SheetNames?.[0];

  if (!sheetName) return [];

  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });

  // Regex you already use in CONFIG
  const urlRe = /\bhttps?:\/\/[^\s)]+?(?:convert_action=convert_vpreview|convert_e=\d{6,}|convert_v=\d{6,})[^\s)]*/gi;

  const links = new Set();

  for (const row of json) {
    for (const cell of row) {
      const txt = String(cell ?? '');
      const matches = txt.match(urlRe);
      if (matches) {
        matches.forEach((m) => links.add(m.trim()));
      }
    }
  }

  return Array.from(links);
}
