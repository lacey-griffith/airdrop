// api/dispatch.js
// Accepts GET or POST from ClickUp. Auth via shared token.
// Dispatches a GitHub workflow with the ClickUp task id as input.

export default async function handler(req, res) {
  try {
    // --- Read inputs (supports GET query or POST JSON) ---
    const method = req.method || 'GET';
    const q = req.query || {};
    const body = typeof req.body === 'object' ? req.body : {};

    const taskId = (q.task_id || body.task_id || q.task || body.task || '').toString().trim();
    const tokenQP = (q.token || '').toString().trim();
    const tokenHdr = (req.headers['x-auth'] || req.headers['X-Auth'] || '').toString().trim();

    // --- Auth check against shared token ---
    const expected = (process.env.SHARED_DISPATCH_TOKEN || '').toString().trim();
    const authOk = !!expected && (tokenQP === expected || tokenHdr === expected);

    console.log('[Dispatch] Incoming', {
      method,
      hasTaskId: !!taskId,
      hasQueryToken: !!tokenQP,
      hasHeaderToken: !!tokenHdr,
      expectedSet: !!expected,
      authOk,
    });

    if (!authOk) {
      return res.status(401).json({ error: 'Unauthorized (bad or missing token)' });
    }
    if (!taskId) {
      return res.status(400).json({ error: 'Missing task_id' });
    }

    // --- GitHub envs ---
    const GH_REPO = (process.env.GH_REPO || '').trim();          // e.g., "lacey-griffith/airdrop"
    const GH_WORKFLOW = (process.env.GH_WORKFLOW || '').trim();  // e.g., "post-qa.yml" OR numeric ID string
    const GH_REF = (process.env.GH_REF || 'main').trim();
    const GH_TOKEN = (process.env.GH_TOKEN || '').trim();

    // Minimal sanity logs (never print secrets)
    console.log('[Dispatch] Target', {
      repo: GH_REPO,
      workflow: GH_WORKFLOW,
      ref: GH_REF,
      hasToken: !!GH_TOKEN,
    });

    if (!GH_REPO || !GH_WORKFLOW || !GH_TOKEN) {
      return res.status(500).json({ error: 'Server misconfigured (missing GH envs)' });
    }

    // --- Call GitHub workflow_dispatch ---
    const url = `https://api.github.com/repos/${GH_REPO}/actions/workflows/${GH_WORKFLOW}/dispatches`;

    const ghRes = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${GH_TOKEN}`,
        'Accept': 'application/vnd.github+json',
      },
      body: JSON.stringify({
        ref: GH_REF,
        inputs: { task: taskId }, // maps to workflow_dispatch.inputs.task
      }),
    });

    const text = await ghRes.text();
    console.log('[Dispatch] GitHub response', { status: ghRes.status, text: text?.slice(0, 400) });

    // GitHub returns 204 No Content on success
    if (ghRes.status === 204) {
      return res.status(200).json({ ok: true, taskId });
    }

    // Bubble up useful detail
    return res.status(502).json({
      error: 'GitHub dispatch failed',
      status: ghRes.status,
      detail: text,
    });
  } catch (err) {
    console.error('[Dispatch] Error', err?.message || err);
    return res.status(500).json({ error: 'Internal error' });
  }
}