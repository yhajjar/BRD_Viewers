# Universal Interactive BRD Viewer (Word task-pane add-in)

A **self-hosted private-test starter**, not a turnkey organization-wide deployment. One Office XML task-pane add-in hosts multiple versioned interactive HTML diagrams. **The final BRDs remain `.docx` files**. The interactive content appears **in a Word task pane**, NOT embedded between document paragraphs.

## Contents

- `web/index.html`, `viewer.js`, `viewer.css` — universal Office.js Word task pane with an `<iframe>` for the currently selected illustration.
- `web/catalog.json` — BRD-to-illustration mappings. Do not put credentials or sensitive tokens in this public catalog.
- `web/illustrations/DEA-2026-238/v1.0/developer-flow.html` — the **original supplied interactive HTML**, unmodified, including S1–S15, API details, navigation, and acceptance checklist.
- `Dockerfile`, `docker-compose.yml`, `nginx.conf` — HTTPS-ready web app behind Coolify's reverse proxy.
- `scripts/generate_manifest.py` — generate the Word XML manifest for YOUR real HTTPS domain.
- `scripts/check_project.py` — static catalog/XML validation.

## Requirements

- An accessible Coolify instance with a publicly or internally routable hostname and a trusted HTTPS certificate.
- Microsoft Word Desktop that supports Office web add-ins (on Windows for the network-share sideloading instructions).
- Permission to configure **Trusted Add-in Catalogs** in Word; some corporate policies block this. For organization-wide deployment, ask Microsoft 365 administration to deploy the add-in centrally.
- Python 3 on the machine on which you generate the manifest; no runtime Python or Node.js is required in the container.

## 1. Deploy web service in Coolify

1. Put this folder in your own Git repository. (The project runs without a separate backend.)
2. In Coolify: **New Resource → Application → Git repository**, select this repo, choose **Dockerfile** as the build pack. Dockerfile path: `./Dockerfile`.
3. Set **container port = 8080**, assign a domain such as `https://brd.example.org`, and enable HTTPS. **Do not publish a host port 80/443 mapping**; let Coolify provide routing and TLS.
4. Deploy. Test `https://brd.example.org/healthz` (should say `ok`), `https://brd.example.org/catalog.json`, and `https://brd.example.org/index.html` in a browser.
5. Word will only show interactive diagrams once its manifest has been installed. Browser preview mode is useful for checking HTML but can't save Word document settings.

**Security warning:** the starter serves all HTML files to anyone who can reach this hostname. For sensitive internal BRDs, put authentication or an organization-managed access-control layer in front of the service before wider use. A shared password via basic auth might prevent Office from loading; test SSO and Office WebView compatibility. The iframe in this starter uses `allow-scripts allow-same-origin` specifically because the user-supplied HTML needs JavaScript, so **publish only trusted/reviewed HTML**, and isolate untrusted contributions on a separate domain with restrictive CSP before production.

## 2. Generate the Office manifest

Replace the sample domain with your actual reachable HTTPS origin:

```bash
python3 scripts/generate_manifest.py --url https://brd.example.org --out deploy/manifest.xml
python3 scripts/check_project.py
```

The manifest uses a stable add-in GUID; keep it unchanged for all BRDs. Changing only `catalog.json` or HTML illustrations does NOT require a new Word add-in installation. If you change ribbon commands or manifest metadata, users may need to reinstall/refresh the manifest.

## 3. Sideload locally for a private Windows Word Desktop test

Official instructions: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins

1. Create a Windows SMB shared folder, e.g. `C:\OfficeAddinCatalog` shared as `\\YOUR-PC\OfficeAddinCatalog` (UNC share). Place only `deploy/manifest.xml` there; the HTML stays hosted on Coolify.
2. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Enter that UNC network share path, click **Add catalog**, select **Show in Menu**, confirm and restart Word.
3. In Word: **Home → Add-ins → Advanced** (wording may vary) → **SHARED FOLDER**, select **Interactive BRD Viewer** and add it.
4. Open the correct DEA-2026-238 BRD `.docx`, then use **Home → Interactive BRD → Open BRD Viewer**.
5. Select `DEA-2026-238` in the BRD dropdown. The original interactive HTML loads within the task pane. Select S1–S15 inside the HTML.
6. Select **Open this viewer automatically with this document**, then click **Associate with this Word document**. Read the success message. **Save the Word document (Ctrl+S)**. Close and reopen the document to test automatic task-pane opening.

**Caveat:** auto-open depends on the exact Office version and trusted/add-in installation state. When auto-open is unavailable, the user can manually click the ribbon button, and the saved BRD association will still be used.

**No website navigation:** clicking through the diagram and using its checklists happens within the Word add-in task pane. Some links within your original illustration (e.g. a Scopus profile link) intentionally navigate externally; this starter does not attempt to intercept third-party links.

## 4. Add another BRD or illustration

Example: Add `web/illustrations/DEA-2026-236/v1.2/approval-flow.html`. Add a corresponding entry to `web/catalog.json`:

```json
{
  "id": "DEA-2026-236",
  "version": "1.2",
  "title": "Research Project Set-Up and Task Requests",
  "illustrations": [
    {
      "id": "approval-flow",
      "title": "Approval Flow",
      "path": "/illustrations/DEA-2026-236/v1.2/approval-flow.html",
      "description": "Interactive approval routing"
    }
  ]
}
```

Put the new object inside `documents` (separated by a comma from the previous object). Then commit and redeploy in Coolify. Existing deployed Word add-ins pick up the new catalog on reopening/reloading; the version in the catalog and URL must be maintained separately for approved historic BRDs. For *two versions of the same BRD to coexist*, assign distinct `id` values (e.g. `DEA-2026-238@1.0` and `DEA-2026-238@1.1`); the provided example uses an ID per current BRD.

## 5. How document mapping works

The UI stores `InteractiveBRD.DocumentId.v1` in **Office.context.document.settings** and calls `.saveAsync()`; those settings travel with the saved Word document. It also sets `Office.AutoShowTaskpaneWithDocument` as a per-document Boolean when requested. The add-in manifest uses `TaskpaneId=Office.AutoShowTaskpaneWithDocument` to enable supported auto-opening. These changes **do not attach or embed the HTML** inside `.docx`.

Official references:
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/persisting-add-in-state-and-settings
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview

## 6. Troubleshooting

- **Static image in Word:** the image is a preview, not an iframe. Use the add-in task pane. Images embedded on the Word page cannot execute arbitrary HTML/JS.
- **Task pane blank:** confirm HTTPS certificate, correct manifest domain, Office add-ins enabled, firewall reachability from the computer, and network/tenant policy.
- **Iframe blank:** open the illustration URL in a browser, verify it is in catalog.json and that CSP `frame-ancestors` permits it. Only same-origin illustrations are loaded in this starter.
- **Auto-open fails:** first open pane manually, associate document, check auto-open, click Save Association, then Ctrl+S. Certain Office environments/policies can disable auto-open.
- **Checklist state disappears:** the supplied HTML uses `localStorage`; it may be isolated, cleared, or partitioned within Office's webview. Shared or durable UAT checklist status requires a server-side authenticated storage feature; NOT implemented in this MVP.
- **Layout narrow:** the original HTML uses a wide sequence diagram with horizontal scrolling. Word task panes are narrow. The MVP preserves original content; a step-focused responsive view is a future enhancement.
- **Not intended as a full access-controlled multiuser platform:** no SSO, catalog write API, web upload portal, centralized progress database, or collaboration system is implemented.

## Test checklist

- [ ] HTTPS server loads `/index.html` and `/catalog.json`.
- [ ] `developer-flow.html` opens in browser and diagram S1–S15 is clickable.
- [ ] Manifest passes XML parse and Word accepts sideload installation.
- [ ] Ribbon opens panel without leaving Word.
- [ ] Choosing DEA-2026-238 displays the diagram and dynamic details.
- [ ] Document association persists after closing and reopening `.docx`.
- [ ] Auto-open works on your Word build; if not, ribbon manual-open works.
- [ ] Test on a second approved document to confirm per-document isolation.

## Optional Docker Compose alternative

A `docker-compose.yml` is included. In Coolify, Dockerfile application deployment is simplest; use the Compose file only if you prefer that mode. Do not bind a public host port. The app listens on port **8080** within the container.
