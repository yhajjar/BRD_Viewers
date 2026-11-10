# Interactive BRD Library — zero-install reviewer experience

A small self-hosted static HTML library for business requirements documents (BRDs). Reviewers open the unique HTTPS link from a normal Word .docx file and interact with the illustration in their browser. **No Word add-in, centralized deployment, or reviewer installation is required.**

## Features

- Multiple BRDs, multiple illustrations each, and versioned HTML directories.
- One direct link per illustration and a searchable landing page.
- Original DEA-2026-238 developer flow preserved, including sequence steps S1–S15, dynamic rules, code copying and acceptance checklist.
- Docker/Nginx ready for Coolify (internal port 8080, external HTTPS via Coolify).
- No external application runtime, database, or API required.

## Folder structure

```text
web/
  index.html
  viewer.js
  viewer.css
  catalog.json
  illustrations/
    DEA-2026-238/
      v1.0/
        developer-flow.html
    DEA-2026-236/
      v1.2/
        project-setup.html        # illustrative only: add when available
```

## Deploy with Coolify

1. Connect the GitHub repository `yhajjar/BRD_Viewers` to a new Coolify Application.
2. Choose the **Dockerfile** build pack. Dockerfile path `./Dockerfile`, expose internal port **8080**, and configure an HTTPS hostname such as `https://brd.example.org`.
3. Do not map container port 80 or 443 to your host; Coolify handles reverse proxy and TLS.
4. Deploy, then open `https://brd.example.org/healthz` (expected `ok`), `https://brd.example.org/catalog.json`, and the homepage.
5. Commit future HTML files and catalog changes to GitHub, then redeploy (or configure Coolify GitHub webhook auto-deploy).

## Add another BRD

1. Create `web/illustrations/<BRD-ID>/v<VERSION>/` and place the HTML plus its assets in that directory. Use relative references for CSS/JS/images.
2. Add a document entry in `web/catalog.json` with `id`, `version`, `title`, and `illustrations`. Each illustration has `id`, `title`, and a **same-origin** `path` like `/illustrations/DEA-2026-238/v1.0/developer-flow.html`.
3. Run `python3 scripts/check_project.py`, commit and redeploy.
4. Select the BRD and illustration on the landing page, click **Copy shareable link**, and paste it as a hyperlink in the Word document (for example after Functional Requirements).
5. For direct full-page access, hyperlink to `https://YOUR-DOMAIN/illustrations/BRD-ID/v1.0/illustration.html`. Both the landing page's shareable URL and direct HTML URL remain independent of the Word add-in.

### Example catalog

```json
{
  "schemaVersion": 1,
  "documents": [
    {
      "id": "DEA-2026-238",
      "version": "1.0",
      "title": "Scopus Integration for Researcher Renewal and Recruitment Forms",
      "illustrations": [
        {
          "id": "developer-flow",
          "title": "Developer Flow (S1-S15)",
          "path": "/illustrations/DEA-2026-238/v1.0/developer-flow.html",
          "description": "Interactive sequence, API notes, design checks and acceptance checklist"
        }
      ]
    }
  ]
}
```

The catalog references only files that actually exist. Use the checker to catch missing paths.

## Link to HTML from Microsoft Word

Insert a heading such as **Interactive Workflow Illustration** after the Functional Requirements. Select a short description, press Ctrl+K, and paste the unique HTTPS illustration URL. Readers click it to open the browser; there is no way to execute arbitrary HTML directly inside a standard Word desktop document without an add-in.

## Security and approvals

**Important:** The repository is currently public and the static site serves files to whoever can reach its URL. Internal workflows, API integration details, and BRD contents may be confidential. Before sharing, review the material and restrict publication appropriately (private network, managed authentication or another suitable access-control layer). Making the GitHub repository private does not automatically protect a deployed public Coolify URL. The preview iframe uses both `allow-scripts` and `allow-same-origin`, so only publish trusted HTML, not unreviewed uploads from other parties.

The example HTML uses browser localStorage for acceptance checkboxes; that status is not shared across reviewers and does not constitute formal BRD sign-off.

## Validation

Run `python3 scripts/check_project.py` before pushing. This validates catalog IDs, illustration IDs, and target HTML paths. Check the actual site after Coolify deployment. No Office.js installation is involved.
