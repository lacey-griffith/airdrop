/**
 * post-qa.js
 * -----------------------------------------------------------------------------
 * Click-to-run (GitHub Actions or local) helper that:
 *   1) Validates a ClickUp task's status + "Passed QA" checkbox
 *   2) Reads the "QA Doc" custom field (SharePoint folder URL)
 *   3) Finds the Excel named like the ClickUp task title, extracts preview links
 *   4) Downloads images from the folder and re-uploads them to the ClickUp task
 *   5) Posts a formatted comment with the links + images (+ optional @mentions)
 *
 * Usage:
 *   node post-qa.js <ClickUp task URL or ID>
 *
 * Env (GitHub Actions → repo Secrets or local .env):
 *   CLICKUP_TOKEN=<clickup personal token>
 *   MS_TENANT_ID=<azure tenant id>           (needed for private SharePoint)
 *   MS_CLIENT_ID=<app registration client id>
 *   MS_CLIENT_SECRET=<app registration secret>
 */

// =========================
// 🔧 CONFIG (edit here)
// =========================
const CONFIG = {
  clickup: {
    requiredStatus: 'Needs Approval Dev',          // gate: must match (case-insensitive)
    fieldNames: {
      passedQA: 'Passed QA',                        // checkbox
      qaDocUrl: 'QA Doc',                           // SharePoint folder URL
      clientMentions: 'Client Mentions',            // optional, comma-separated names
    },
    // Friendly name → ClickUp userId (add later when you create the field)
    mentionMap: {
      // "Acme PM": "12345678"
    },
  },
  notifications: {
  postGateFailureComment: true, // posts a comment when gates fail
},
  sharepoint: {
    // Excel selection policy:
    //   1) exact "<taskTitle>.xlsx|xls"
    //   2) startsWith "<taskTitle> " (e.g., " - v2.xlsx")
    //   3) filename contains "preview"
    //   4) any .xlsx as last resort
    imageExtensions: /\.(png|jpe?g|webp|gif)$/i,
  },

  // Convert preview URLs we care about (adjust when needed)
  linkPatterns: {
    previewUrl: /\bhttps?:\/\/[^\s)]+?(?:convert_action=convert_vpreview|convert_e=\d{6,}|convert_v=\d{6,})[^\s)]*/gi,
  },

  // Attachments already on the task (fallback if SharePoint images aren’t accessible)
  taskAttachmentFilter: {
    restrictToQAishNames: false,                    // set true to only include filenames containing "qa"
    qaNamePattern: /(^|\/|_|-|\s)qa($|\.|\s|_|-)/i, // only used if restrictToQAishNames = true
  },

  // Toggle logging verbosity
  logging: {
    verbose: true,
  },
};
// =========================
// end CONFIG
// =========================

import 'dotenv/config';
import fetch from 'node-fetch';
import FormData from 'form-data';
import {
  getGraphToken,
  listFolderItems,
  downloadItemBuffer,
  findExcelForTask,
  getImageFiles,
  extractPreviewLinksFromXlsx,
} from './sharepoint.js';

// -------------------------
// ClickUp REST helpers
// -------------------------
const CLICKUP_API = 'https://api.clickup.com/api/v2';
const CLICKUP_TOKEN = process.env.CLICKUP_TOKEN;

if (!CLICKUP_TOKEN) {
  console.error('Missing CLICKUP_TOKEN env. Add it to repo Secrets or .env');
  process.exit(1);
}

function log(...args) {
  if (CONFIG.logging.verbose) console.log(...args);
}

async function cuGet(path) {
  const res = await fetch(`${CLICKUP_API}${path}`, {
    headers: { Authorization: CLICKUP_TOKEN },
  });
  if (!res.ok) throw new Error(`ClickUp GET ${path} failed: ${res.status}`);
  return res.json();
}

async function cuPost(path, body) {
  const res = await fetch(`${CLICKUP_API}${path}`, {
    method: 'POST',
    headers: {
      Authorization: CLICKUP_TOKEN,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`ClickUp POST ${path} failed: ${res.status} ${txt}`);
  }
  return res.json();
}

async function cuUploadAttachment(taskId, filename, buffer) {
  const form = new FormData();
  form.append('attachment', buffer, { filename });

  const res = await fetch(`${CLICKUP_API}/task/${taskId}/attachment`, {
    method: 'POST',
    headers: { Authorization: CLICKUP_TOKEN },
    body: form,
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`ClickUp attachment upload failed: ${res.status} ${txt}`);
  }
  return res.json();
}

// -------------------------
// ClickUp task helpers
// -------------------------
function parseTaskId(input) {
  const m = String(input).match(/\/t\/([^/?#]+)/i);
  return m ? m[1] : String(input);
}

function getCustomField(task, fieldName) {
  return (task.custom_fields || []).find(
    (cf) => (cf.name || '').trim() === fieldName.trim(),
  ) || null;
}

function isCheckboxChecked(cf) {
  return cf && (cf.value === true || cf.value === 1 || cf.value === '1');
}

function getTextFieldValue(cf) {
  return cf && typeof cf.value === 'string' ? cf.value : '';
}

function extractUrlsFromText(text = '') {
  return text.match(/\bhttps?:\/\/[^\s)]+/gi) || [];
}

function extractPreviewLinksFromText(txt) {
  return (txt.match(CONFIG.linkPatterns.previewUrl) || []).map((u) => u.trim());
}

function selectTaskImageAttachments(task) {
  const atts = task.attachments || [];
  return atts.filter((att) => {
    const name = (att.title || att.name || '').toLowerCase();
    const mime = (att.mime_type || '').toLowerCase();
    const isImage =
      mime.startsWith('image/') || /\.(png|jpg|jpeg|gif|webp)$/i.test(name);

    if (!isImage) return false;
    if (!CONFIG.taskAttachmentFilter.restrictToQAishNames) return true;

    return (
      CONFIG.taskAttachmentFilter.qaNamePattern.test(name) ||
      CONFIG.taskAttachmentFilter.qaNamePattern.test(att.path || '')
    );
  });
}

function resolveMentionIdsFromTask(task) {
  const fieldName = CONFIG.clickup.fieldNames.clientMentions;
  if (!fieldName) return [];
  const cf = getCustomField(task, fieldName);
  const raw = getTextFieldValue(cf);
  if (!raw) return [];

  return raw
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((label) => CONFIG.clickup.mentionMap[label])
    .filter(Boolean);
}

function formatMentions(userIds) {
  if (!userIds || !userIds.length) return '';
  return userIds.map((id) => `<@${id}>`).join(' ');
}

function buildComment({ taskName, previewLinks, images, mentionIds }) {
  const lines = [];

  if (mentionIds?.length) {
    lines.push(`${formatMentions(mentionIds)}\n`);
  }

  lines.push(`**QA Passed → Preview Links for _${taskName}_**\n`);

  if (previewLinks.length) {
    lines.push('**Preview Links**');
    previewLinks.forEach((url, i) => lines.push(`- [Link ${i + 1}](${url})`));
    lines.push('');
  } else {
    lines.push('_No preview links found._\n');
  }

  if (images.length) {
    lines.push('**QA Images**');
    images.forEach((img) => lines.push(`- ${img.name || 'image'} → ${img.url}`));
  } else {
    lines.push('_No QA images found._');
  }

  return lines.join('\n');
}

// -------------------------
// Main
// -------------------------
const INPUT = process.argv[2] || '';
if (!INPUT) {
  console.error('Usage: node post-qa.js <ClickUp task URL or ID>');
  process.exit(1);
}
const TASK_ID = parseTaskId(INPUT);

(async function run() {
  try {
   // 1) Load task + validate gate
const task = await cuGet(`/task/${TASK_ID}`);
const taskName = task.name || TASK_ID;

const statusName =
  (task.status && (task.status.status || task.status.name)) || '';
const requiredStatus = CONFIG.clickup.requiredStatus;
const statusOk = new RegExp(`^${requiredStatus}$`, 'i').test(statusName);

const cfPassed = getCustomField(task, CONFIG.clickup.fieldNames.passedQA);
const passedQA = isCheckboxChecked(cfPassed);

// 🔔 If gates fail, optionally post a human-friendly comment on the task and exit
if (!statusOk || !passedQA) {
  console.log('AirDrop gate failed; not posting.');

  if (CONFIG.notifications.postGateFailureComment) {
    const failMsg = `🪂 AirDrop Status: Fail. Status must be [Needs Approval (Dev)] and Passed QA must be checked. Current Status: [${statusName || 'Unknown'}].`;
    try {
      await cuPost(`/task/${TASK_ID}/comment`, { comment_text: failMsg });
      console.log('Posted AirDrop failure status comment.');
    } catch (e) {
      console.warn('Could not post failure comment:', e.message || e);
    }
  }

  process.exit(0);
}


    // 2) Resolve SharePoint folder URL from "QA Doc"
    const cfDoc = getCustomField(task, CONFIG.clickup.fieldNames.qaDocUrl);
    const qaFolderUrl = getTextFieldValue(cfDoc);
    if (!qaFolderUrl) {
      console.log(`⛔ No "${CONFIG.clickup.fieldNames.qaDocUrl}" URL present on task. Nothing to do.`);
      process.exit(0);
    }

    // 3) List folder items via Graph (preferred)
    const graphToken = await getGraphToken();
    let previewLinks = [];
    let uploadedImages = [];

    if (graphToken) {
      try {
        const items = await listFolderItems({ folderUrl: qaFolderUrl, token: graphToken });

        // 3a) Excel selection (based on task title), then extract links
        const excelItem = findExcelForTask(items, taskName);
        if (excelItem) {
          log(`Excel selected: ${excelItem.name}`);
          const buf = await downloadItemBuffer({
            driveId: excelItem.driveId,
            itemId: excelItem.id,
            token: graphToken,
          });
          previewLinks = extractPreviewLinksFromXlsx(buf);
        } else {
          log('No Excel found in folder.');
        }

        // 3b) Images: download from SharePoint → re-upload to ClickUp
        const imageItems = getImageFiles(items, CONFIG.sharepoint.imageExtensions);
        for (const img of imageItems) {
          const buf = await downloadItemBuffer({
            driveId: img.driveId,
            itemId: img.id,
            token: graphToken,
          });
          const uploaded = await cuUploadAttachment(TASK_ID, img.name, buf);
          const uploadedUrl = uploaded?.data?.url || uploaded?.url || '';
          uploadedImages.push({ name: img.name, url: uploadedUrl });
        }
      } catch (err) {
        console.warn('SharePoint parse warning:', err.message || err);
      }
    } else {
      console.warn('No Graph token available. Skipping SharePoint folder read (private folders require MS_* secrets).');
    }

    // 4) Fallbacks if Excel or Graph was unavailable
    if (!previewLinks.length) {
      // Try task description for preview URLs (safety net)
      const descUrls = extractUrlsFromText(task.description || '');
      previewLinks = extractPreviewLinksFromText(descUrls.join('\n'));
    }

    const taskAttachments = selectTaskImageAttachments(task).map((a) => ({
      name: a.title || a.name || 'image',
      url: a.url,
    }));

    const finalImages = uploadedImages.length ? uploadedImages : taskAttachments;

    // 5) Mentions (optional)
    const mentionIds = resolveMentionIdsFromTask(task);

    // 6) Build + post comment
    const uniqueLinks = Array.from(new Set(previewLinks));
    const commentText = buildComment({
      taskName,
      previewLinks: uniqueLinks,
      images: finalImages,
      mentionIds,
    });

    await cuPost(`/task/${TASK_ID}/comment`, { comment_text: commentText });
    console.log('✅ Posted QA preview comment to ClickUp.');
  } catch (err) {
    console.error('Error:', err.message || err);
    process.exit(1);
  }
})();
