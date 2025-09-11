/**
 * api/dispatch.js
 * ---------------------------------------------------------------------------
 * Vercel Serverless Function: ClickUp Button → GitHub workflow_dispatch
 *
 * ClickUp Automation (Call webhook) URL example:
 *   https://<your-app>.vercel.app/api/dispatch?task_id={{task.id}}&token=<SHARED_DISPATCH_TOKEN>
 *
 * Vercel Project → Settings → Environment Variables (add and redeploy):
 *   GH_REPO               e.g. "your-org/your-repo"
 *   GH_WORKFLOW           e.g. "post-qa.yml"  (the workflow file name)
 *   GH_REF                e.g. "main"
 *   GH_TOKEN              GitHub PAT with repo + actions:write
 *   SHARED_DISPATCH_TOKEN long random string (same one you put in the ClickUp webhook URL)
 */

export default async function handler(req, res) {
  try {
    const {
      GH_REPO,
      GH_WORKFLOW,
      GH_REF = 'main',
      GH_TOKEN,
      SHARED_DISPATCH_TOKEN,
    } = process.env;

    if (!GH_REPO || !GH_WORKFLOW || !GH_TOKEN || !SHARED_DISPATCH_TOKEN) {
      return res.status(500).json({ error: 'Missing required env vars' });
    }

    // Accept GET or POST
    const isJson = (req.headers['content-type'] || '').includes('application/json');
    const body = isJson ? await readJson(req) : {};
    const q = req.query || {};

    // Shared secret check (query, body, or X-Auth header)
    const providedToken = q.token || body.token || req.headers['x-auth'] || '';
    if (providedToken !== SHARED_DISPATCH_TOKEN) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    // Task id from query/body
    const taskId = (q.task_id || body.task_id || '').trim();
    if (!taskId) return res.status(400).json({ error: 'Missing task_id' });

    // GitHub workflow_dispatch
    const ghUrl = `https://api.github.com/repos/${GH_REPO}/actions/workflows/${GH_WORKFLOW}/dispatches`;
    const payload = { ref: GH_REF, inputs: { task: taskId } };

    const ghRes = await fetch(ghUrl, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${GH_TOKEN}`,
        Accept: 'application/vnd.github+json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    if (!ghRes.ok) {
      return res.status(502).json({ error: 'GitHub dispatch failed', detail: await safeText(ghRes) });
    }

    return res.status(200).json({ ok: true, taskId });
  } catch (err) {
    return res.status(500).json({ error: err?.message || 'Server error' });
  }
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => (data += c));
    req.on('end', () => {
      try { resolve(data ? JSON.parse(data) : {}); } catch (e) { reject(e); }
    });
    req.on('error', reject);
  });
}

async function safeText(res) {
  try { return await res.text(); } catch { return '<no body>'; }
}
