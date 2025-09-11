# 🪂 AIRDROP

AirDrop turns a **ClickUp button** into a one-click client handoff: **Send to Client** → Vercel `/api/dispatch` → GitHub Action → posts preview links + QA images to the ClickUp task.

---

## TL;DR
- In ClickUp, click **Send to Client** on the task.
- AirDrop validates gates: **Status = “Needs Approval (Dev)”** and **Passed QA** is checked.
- It reads the **QA Doc** field (a SharePoint folder URL). Inside that folder, it locates the **Excel named after the task title** (exact match or starts-with), extracts preview links using smart rules (prefer **Sheet 2** column headed **“Preview Links”**; otherwise scan configured columns on **Sheet 1**), uploads QA images, and posts a formatted comment to the task.
- If gates fail, it posts a clear failure comment on the task:
`🪂 AirDrop Status: Fail. Status must be [Needs Approval (Dev)] and Passed QA must be checked. Current Status: [Strategy].`
---
## Repo layout
    <repo-root>/
    ├─ api/
    │  └─ dispatch.js
    ├─ .github/
    │  └─ workflows/
    │     └─ post-qa.yml
    ├─ post-qa.js
    ├─ sharepoint.js
    ├─ package.json
    ├─ .gitignore
    └─ README.md

---

## Prereqs
- **ClickUp** personal token with access to your workspace.
- **GitHub** repo for this code.
- **Vercel** project connected to the repo.
- **(If SharePoint is private)** Azure App Registration with Graph *application* permissions:
  - `Files.Read.All`
  - `Sites.Read.All`
  - (Grant admin consent)

---

## Configure secrets

### GitHub (Repo → Settings → Secrets and variables → Actions)
Required:
- `CLICKUP_TOKEN`

If using private SharePoint:
- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`

### Vercel (Project → Settings → Environment Variables)
- `GH_REPO` = `owner/repo`
- `GH_WORKFLOW` = `post-qa.yml`
- `GH_REF` = `main`
- `GH_TOKEN` = GitHub PAT with `repo` and `workflow` (actions:write)
- `SHARED_DISPATCH_TOKEN` = long random string (used in ClickUp webhook)

> After adding Vercel envs, redeploy.

---

## ClickUp setup

### Custom fields (task level)
- **Passed QA** (Checkbox)
- **QA Doc** (Text/URL) → SharePoint **folder** URL (contains Excel + images)
- **Send to Client** (Button)

### Automation (per List/Folder/Space)
- **Trigger:** When Button **Send to Client** is clicked  
- **Conditions:**
  - *Status is* `Needs Approval (Dev)`
  - *Passed QA* is checked
- **Action:** **Call webhook**  
      https://<your-vercel-app>.vercel.app/api/dispatch?task_id={{task.id}}&token=<YOUR_SHARED_DISPATCH_TOKEN>

---

## Excel extraction rules (smart)
1. If a **second sheet** exists and its **first row** contains header **`Preview Links`** (case-insensitive), AirDrop grabs links from **that column** (rows below the header).
2. Otherwise AirDrop falls back to the **first sheet**, scanning configured **columns** (e.g., `A,B`) from a configured **start row** (e.g., row 2).

> Configure in `CONFIG.excelExtraction` at the top of `post-qa.js`.

---

## SharePoint folder expectations
- Contains **one Excel** named like the **ClickUp task title** (e.g., `CF-123 Add Sticky Header.xlsx`). `"Title - v2.xlsx"` also matches.
- Contains **QA images** (`.png/.jpg/.webp/.gif`) that should be included in the comment.

---

## How to use (happy path)
1. On the task, set:
   - **Status** = `Needs Approval (Dev)`
   - **Passed QA** = checked
   - **QA Doc** = SharePoint folder with Excel + images
2. Click **Send to Client**.
3. Watch GitHub **Actions** run **“AirDrop — Post QA Preview”**.
4. The task gets a **formatted comment** with preview links + images.

---

## How to test with a demo task
Example: `https://app.clickup.com/t/868fj80zu`

**Fail test (gates not met)**
- Leave Status as something else (e.g., `Strategy`) and/or uncheck **Passed QA**.
- Click **Send to Client**.
- Expected **comment** on the task:
      🪂 AirDrop Status: Fail. Status must be [Needs Approval (Dev)] and Passed QA must be checked. Current Status: [Strategy].

**Pass test**
- Change Status to **Needs Approval (Dev)** and check **Passed QA**.
- Ensure **QA Doc** points to a folder with:
  - Excel named like the task title (Sheet2 header “Preview Links” if using the header path, else fallback columns on Sheet1)
  - QA images
- Click **Send to Client**.
- Expected: “AirDrop — QA Passed …” comment with deduped links + image list.

---

## Local debug (optional)
    # .env (DO NOT COMMIT)
    CLICKUP_TOKEN=...
    MS_TENANT_ID=...
    MS_CLIENT_ID=...
    MS_CLIENT_SECRET=...

    npm ci
    node post-qa.js <taskId or https://app.clickup.com/t/...>

Check console output; the script logs what it found and posts to the task.

---

## Troubleshooting
- **GitHub Action fails at secrets** → ensure `CLICKUP_TOKEN` exists; remove `MS_*` env lines from workflow if not using private SharePoint yet.
- **Vercel dispatch 401** → ClickUp webhook `token` must match `SHARED_DISPATCH_TOKEN`.
- **No Excel found** → Make sure filename matches the task title (or contains “preview” for fallback).
- **No links extracted** → Confirm header cell is exactly `Preview Links` on Sheet2 row 1; otherwise verify fallback columns in `CONFIG.excelExtraction`.
- **SharePoint 403/401** → Confirm Graph App has `Files.Read.All` + `Sites.Read.All` with admin consent; secrets set correctly.

---

## Branding
- Comment header: `🪂 AirDrop — QA Passed for _<Task Title>_`
- ClickUp Button label: **Send to Client**
- Workflow name: **AirDrop — Post QA Preview**

---

## Security
- Tokens live only in **GitHub/Vercel envs** (never in code).
- Vercel endpoint requires a **shared token**; ClickUp calls include it in the URL.
- GitHub Action reads repo **secrets**; no plaintext tokens in logs or code.
