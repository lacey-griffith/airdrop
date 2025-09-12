/**
 * post-qa.js
 * -----------------------------------------------------------------------------
 * Click-to-run (GitHub Actions or local) helper that:
 *   1) Validates a ClickUp task's status + "Passed QA" checkbox
 *   2) Reads the "QA Doc" custom field (SharePoint folder URL)
 *   3) Finds the Excel named like the ClickUp task title, extracts preview links
 *   4) Downloads images from the folder and re-uploads them to the ClickUp task
 *   5) Posts a DRAFT comment with the links + images (no notifications)
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
    // Use the human label shown in your UI; we normalize for comparison.
    requiredStatus: "Needs Approval (Dev)",

    fieldNames: {
      passedQA: "Passed QA",  // checkbox
      qaDocUrl: "QA Doc",     // SharePoint folder URL (folder that contains images + Excel)
      clientMentions: "Client Mentions", // optional, comma-separated names mapped via mentionMap
    },

    // Friendly name → ClickUp userId (only used if you later enable mentions in FINAL mode)
    mentionMap: {
      // "Acme PM": "12345678"
    },
  },

  notifications: {
    postGateFailureComment: true, // post a failure comment when gates aren't met
  },

  sharepoint: {
    // Excel selection policy is implemented in sharepoint.js (findExcelForTask)
    imageExtensions: /\.(png|jpe?g|webp|gif)$/i,
  },

  // Convert preview URLs we care about (adjust when needed)
  linkPatterns: {
    previewUrl:
      /\bhttps?:\/\/[^\s)]+?(?:convert_action=convert_vpreview|convert_e=\d{6,}|convert_v=\d{6,})[^\s)]*/gi,
  },

  // Attachments already on the task (fallback if SharePoint images aren’t accessible)
  taskAttachmentFilter: {
    restrictToQAishNames: false,                    // true → only include names containing "qa"
    qaNamePattern: /(^|\/|_|-|\s)qa($|\.|\s|_|-)/i, // only used if restrictToQAishNames = true
  },

  // Logging
  logging: { verbose: true },

  // Always create a draft (no notifications). You will manually send a client-ready message after review.
  commentMode: "draft", // 'draft' | 'final'

  draft: {
    addBanner: true,
    bannerText: "📝 DRAFT — Review before sending",
    notifyAll: false,       // do NOT ping watchers on draft
    saveToTextField: null,  // disabled by request: do not write to any Text CF
  },

  final: {
    notifyAll: true,        // notify watchers (only used if you ever switch to 'final')
    includeMentions: true,  // include @mentions (only in 'final')
  },
};
// =========================
// end CONFIG
// =========================

import "dotenv/config";
import fetch from "node-fetch";
import FormData from "form-data";
import {
  getGraphToken,
  listFolderItems,
  listFolderChildren,
  downloadItemBuffer,
  findTaskSubfolder,
  findExcelForTask,
  getImageFiles,
  extractPreviewLinksFromXlsx,
} from './sharepoint.js';


// -------------------------
// Small utils
// -------------------------
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const normalize = (s = "") =>
  String(s)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\w\s]/g, "")
    .trim();

function parseTaskId(input) {
  const m = String(input).match(/\/t\/([^/?#]+)/i);
  return m ? m[1] : String(input);
}

function log(...args) {
  if (CONFIG.logging.verbose) console.log(...args);
}

// -------------------------
// ClickUp REST helpers
// -------------------------
const CLICKUP_API = "https://api.clickup.com/api/v2";
const CLICKUP_TOKEN = process.env.CLICKUP_TOKEN;

if (!CLICKUP_TOKEN) {
  console.error("Missing CLICKUP_TOKEN env. Add it to repo Secrets or .env");
  process.exit(1);
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
    method: "POST",
    headers: {
      Authorization: CLICKUP_TOKEN,
      "Content-Type": "application/json",
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
  form.append("attachment", buffer, { filename });

  const res = await fetch(`${CLICKUP_API}/task/${taskId}/attachment`, {
    method: "POST",
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
function getCustomField(task, fieldName) {
  return (
    (task.custom_fields || []).find(
      (cf) => (cf.name || "").trim() === fieldName.trim()
    ) || null
  );
}

/** Find a custom field by ID or name (case-insensitive) */
function findCustomField(task, target) {
  const list = task?.custom_fields || [];
  const needle = String(target || "").trim().toLowerCase();
  return list.find(
    (cf) =>
      (cf.id && String(cf.id).toLowerCase() === needle) ||
      (cf.name && String(cf.name).trim().toLowerCase() === needle)
  );
}

/** ClickUp checkboxes can be true/false, 1/0, and string encodings */
function isCheckboxChecked(cfObj) {
  const v = cfObj?.value;
  if (v === true || v === 1) return true;
  if (typeof v === "string") {
    const s = v.trim().toLowerCase();
    return s === "true" || s === "1" || s === "yes" || s === "checked" || s === "on";
  }
  return false;
}

function getTextFieldValue(cf) {
  return cf && typeof cf.value === "string" ? cf.value : "";
}

function extractUrlsFromText(text = "") {
  return text.match(/\bhttps?:\/\/[^\s)]+/gi) || [];
}

function extractPreviewLinksFromText(txt) {
  return (txt.match(CONFIG.linkPatterns.previewUrl) || []).map((u) => u.trim());
}

function selectTaskImageAttachments(task) {
  const atts = task.attachments || [];
  return atts.filter((att) => {
    const name = (att.title || att.name || "").toLowerCase();
    const mime = (att.mime_type || "").toLowerCase();
    const isImage =
      mime.startsWith("image/") || /\.(png|jpg|jpeg|gif|webp)$/i.test(name);

    if (!isImage) return false;
    if (!CONFIG.taskAttachmentFilter.restrictToQAishNames) return true;

    return (
      CONFIG.taskAttachmentFilter.qaNamePattern.test(name) ||
      CONFIG.taskAttachmentFilter.qaNamePattern.test(att.path || "")
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
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((label) => CONFIG.clickup.mentionMap[label])
    .filter(Boolean);
}

function formatMentions(userIds) {
  if (!userIds || !userIds.length) return "";
  return userIds.map((id) => `<@${id}>`).join(" ");
}

/** Build the comment. Mentions are included only if explicitly allowed. */
function buildComment({ taskName, previewLinks, images, mentionIds, includeMentions }) {
  const lines = [];

  if (includeMentions && mentionIds?.length) {
    lines.push(`${formatMentions(mentionIds)}\n`);
  }

  lines.push(`**QA Passed → Preview Links for _${taskName}_**\n`);

  if (previewLinks.length) {
    lines.push("**Preview Links**");
    previewLinks.forEach((url, i) => lines.push(`- [Link ${i + 1}](${url})`));
    lines.push("");
  } else {
    lines.push("_No preview links found._\n");
  }

  if (images.length) {
    lines.push("**QA Images**");
    images.forEach((img) => lines.push(`- ${img.name || "image"} → ${img.url}`));
  } else {
    lines.push("_No QA images found._");
  }

  return lines.join("\n");
}

// -------------------------
// Main
// -------------------------
const INPUT = process.argv[2] || "";
if (!INPUT) {
  console.error("Usage: node post-qa.js <ClickUp task URL or ID>");
  process.exit(1);
}
const TASK_ID = parseTaskId(INPUT);

(async function run() {
  try {
    console.log("[AirDrop] Start", {
      taskId: TASK_ID,
      mode: CONFIG.commentMode,
      requiredStatus: CONFIG.clickup.requiredStatus,
    });

    // 1) Load task + validate gates
    const task = await cuGet(`/task/${TASK_ID}`);
    const taskName = task.name || TASK_ID;

    log("[AirDrop] Task basics", {
      name: taskName,
      statusRaw: task?.status?.status || task?.status?.name || "",
    });

    // Status (robust normalization on both sides)
    const requiredStatus = CONFIG.clickup.requiredStatus;
    const requiredStatusNorm = normalize(requiredStatus);

    let statusNow = (task.status && (task.status.status || task.status.name)) || "";
    let statusNowNorm = normalize(statusNow);

    let statusOk = statusNowNorm === requiredStatusNorm;

    // Passed QA
    const passedFieldNameOrId = CONFIG.clickup.fieldNames.passedQA;
    const cfPassed = findCustomField(task, passedFieldNameOrId);
    let passedQA = isCheckboxChecked(cfPassed);

    log("[AirDrop] Gates", {
      requiredStatus,
      requiredStatusNorm,
      statusNow,
      statusNowNorm,
      statusOk,
      passedField: cfPassed
        ? { id: cfPassed.id, name: cfPassed.name, type: cfPassed.type, raw: cfPassed.value }
        : null,
      passedQA,
    });

    // Safety-net: if status is still QA (or QA Dev) but Passed QA is true, wait once and re-check
    const qaNorm = normalize("QA");
    const qaDevNorm = normalize("QA (Dev)");
    if (!statusOk && passedQA && (statusNowNorm === qaNorm || statusNowNorm === qaDevNorm)) {
      console.log("[AirDrop] Status still QA while Passed QA is true; waiting 1.5s for automation to land…");
      await sleep(1500);
      const t2 = await cuGet(`/task/${TASK_ID}`);
      statusNow = (t2.status && (t2.status.status || t2.status.name)) || "";
      statusNowNorm = normalize(statusNow);
      statusOk = statusNowNorm === requiredStatusNorm;
      log("[AirDrop] Re-check after wait", { statusNow, statusNowNorm, statusOk });
    }

    // Final decision
    if (!statusOk || !passedQA) {
      if (CONFIG.notifications?.postGateFailureComment) {
        const failMsg =
          `🪂 AirDrop Status: Fail. Status must be [${requiredStatus}] and Passed QA must be checked. ` +
          `Current Status: [${statusNow || "Unknown"}].`;
        try {
          await cuPost(`/task/${TASK_ID}/comment`, { comment_text: failMsg });
          console.log("[AirDrop] Posted gate-failure comment.");
        } catch (e) {
          console.warn("[AirDrop] Could not post failure comment:", e?.message || e);
        }
      }
      return; // stop early
    }

    // 2) Resolve SharePoint folder URL from "QA Doc"
    // Use findCustomField (case-insensitive / id or name)
    const cfDoc = findCustomField(task, CONFIG.clickup.fieldNames.qaDocUrl);
    const qaFolderUrl = getTextFieldValue(cfDoc);

    log("[AirDrop] QA Doc field", {
      fieldFound: !!cfDoc,
      fieldId: cfDoc?.id || null,
      qaFolderUrl: qaFolderUrl || "(none)",
    });

    if (!qaFolderUrl) {
      console.log(`⛔ No "${CONFIG.clickup.fieldNames.qaDocUrl}" URL present on task. Nothing to do.`);
      return;
    }

    // 3) List folder items via Graph (preferred)
    const graphToken = await getGraphToken();
    let previewLinks = [];
    let uploadedImages = [];

    log("[AirDrop] Graph token present:", !!graphToken);

    if (graphToken) {
      try {
        // Root children
        const rootItems = await listFolderItems({ folderUrl: qaFolderUrl, token: graphToken });

        // If a child folder matches the task name, drill in one level
        let items = rootItems;
        const taskFolder = findTaskSubfolder(rootItems, taskName);
        if (taskFolder) {
          log(`[SP] Drilling into task subfolder: ${taskFolder.name}`);
          items = await listFolderChildren({ driveId: taskFolder.driveId, itemId: taskFolder.id, token: graphToken });
        } else {
          log('[SP] No matching task subfolder found; using root folder');
        }

        // 3a) Excel selection (based on task title), then extract links
        const excelItem = findExcelForTask(items, taskName);

        if (excelItem) {
          log("[AirDrop] Excel chosen:", { name: excelItem.name, id: excelItem.id });
          const buf = await downloadItemBuffer({
            driveId: excelItem.driveId,
            itemId: excelItem.id,
            token: graphToken,
          });
          previewLinks = extractPreviewLinksFromXlsx(buf);
          log("[AirDrop] Preview links from Excel:", previewLinks.length);
          if (previewLinks.length) log("[AirDrop] Sample link:", previewLinks[0]);
        } else {
          log("[AirDrop] No Excel found in folder.");
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
          const uploadedUrl = uploaded?.data?.url || uploaded?.url || "";
          uploadedImages.push({ name: img.name, url: uploadedUrl });
          log("[AirDrop] Uploaded image", { name: img.name, uploadedUrl });
        }
      } catch (err) {
        console.warn("SharePoint parse warning:", err.message || err);
      }
    } else {
      console.warn("No Graph token available. Skipping SharePoint folder read (private folders require MS_* secrets).");
    }

    // 4) Fallbacks if Excel/Graph unavailable
    if (!previewLinks.length) {
      log("[AirDrop] No Excel links; trying description…");
      const descUrls = extractUrlsFromText(task.description || "");
      previewLinks = extractPreviewLinksFromText(descUrls.join("\n"));
    }
    log("[AirDrop] Preview links (final set):", previewLinks.length);

    const taskAttachments = selectTaskImageAttachments(task).map((a) => ({
      name: a.title || a.name || "image",
      url: a.url,
    }));
    const finalImages = uploadedImages.length ? uploadedImages : taskAttachments;

    // 5) Mentions (optional)
    const mentionIds = resolveMentionIdsFromTask(task);

    // 6) Build + post (always as DRAFT preview)
    const isDraft = CONFIG.commentMode === "draft";
    const includeMentions = !isDraft && CONFIG.final.includeMentions;
    const uniqueLinks = Array.from(new Set(previewLinks));

    let commentText = buildComment({
      taskName,
      previewLinks: uniqueLinks,
      images: finalImages,
      mentionIds,
      includeMentions,
    });

    if (isDraft && CONFIG.draft.addBanner) {
      commentText = `${CONFIG.draft.bannerText}\n\n${commentText}`;
    }

    log("[AirDrop] Posting comment (draft mode)", {
      notify: isDraft ? CONFIG.draft.notifyAll : CONFIG.final.notifyAll,
      links: uniqueLinks.length,
      images: finalImages.length,
    });

    await cuPost(`/task/${TASK_ID}/comment`, {
      comment_text: commentText,
      notify_all: isDraft ? CONFIG.draft.notifyAll : CONFIG.final.notifyAll,
    });

    console.log(`✅ Posted ${isDraft ? "DRAFT" : "FINAL"} QA preview comment to ClickUp.`);
  } catch (err) {
    console.error("Error:", err?.message || err);
    if (err?.stack) console.error(err.stack);
    process.exit(1);
  }
})();
