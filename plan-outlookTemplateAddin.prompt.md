## Plan: Outlook template add-in hosted on GitHub Pages

This is an Office web add-in with a task pane that opens while composing an email. It runs in classic Outlook for Windows, new Outlook and Outlook on the web. The add-in code and the templates live in one GitHub repository and are served from GitHub Pages on the same domain. Templates are HTML files in one folder. On every push to `main`, a GitHub Actions workflow tests everything, runs a PowerShell script that builds an `index.json` list of templates, builds the add-in, and deploys the result to Pages. The add-in shows the template list with search and a category filter, previews the selected template, and inserts it where the cursor is. You'll deploy it through Integrated Apps in the Microsoft 365 admin center, with no Entra app registration and no server to maintain.

**Working constraint:** prepare the complete, portable source tree on the current corporate PC without installing software, packages, extensions, PowerShell modules, or project generators, and without running local builds or tests. Create configuration and source files directly. Keep generated dependencies and build output out of the project. GitHub Actions can install dependencies on its hosted runners and perform CI/build/deployment; any interactive Outlook sideloading or manual verification is deferred until the files are moved to an unrestricted computer. Do not make local development a prerequisite for completing this plan.

**Why not SharePoint:** any way an add-in can read SharePoint today needs an Entra app registration. Since you want to avoid one, SharePoint and per-group permission trimming are out of scope.

**Steps**

*Phase 1: Scaffold*
1. Create the project source and configuration directly in `addin/` (do not install or run Yeoman): task pane, React, TypeScript, Outlook. Use Fluent UI v9. Include the package manifest and lockfile/configuration needed for a later `npm ci` on GitHub Actions or the destination computer, but do not create `node_modules/` or build output on this PC.
2. Manifest: use the unified JSON manifest, with the older XML manifest as a fallback if needed. Activate in compose mode only, with a ribbon button on new message, reply and forward. Make the task pane pinnable. Require Mailbox requirement set 1.5 as the minimum. The source location points to `https://<org>.github.io/<repo>/taskpane.html`, the default GitHub Pages address (no DNS changes needed). All URLs (manifest, base URL, webpack `publicPath`) come from one build setting, so switching to a custom domain later is a config change.

*Phase 2: Template format and publish script (can run in parallel with Phase 3)*

3. Template format: one `*.html` file per template in a flat `templates/` folder.
   - Metadata goes in `<meta>` tags in `<head>`: `template-name` and `template-category` are required; `template-description` and `template-keywords` are optional.
   - The content that gets inserted is everything inside `<body>`.
   - Images go in `templates/images/` and can be referenced with relative paths.
4. Add 3–4 sample templates with a logo image.
5. `scripts/Publish-TemplateIndex.ps1` takes `-TemplateRoot`, `-OutputPath` and `-BaseUrl`. It runs in GitHub Actions and can also run on the destination computer later. It:
   - parses the meta tags in each template;
   - checks that required fields exist, that there are no `<script>` tags or `on*` event attributes, and that every referenced image file exists;
   - writes `index.json` into the build output (it is generated, not committed) with id, name, category, description, keywords, file URL and last-modified date (from the last git commit of each file);
   - returns a non-zero exit code if any template fails these checks, which stops the deployment.
6. Pester 5+ tests in `tests/Publish-TemplateIndex.Tests.ps1`, using `TestDrive` fixtures for valid, missing-meta, script-tag and missing-image cases.

*Phase 3: Task pane UI (depends on Phase 1)*

7. Load `index.json` and template files with a `?v=<build id>` query string. GitHub Pages caches files for about 10 minutes and its headers can't be changed, so the build id (set at build time) makes sure each deployment loads fresh content. The templates base URL is also a build-time setting.
8. Layout: a search box (matches name, description and keywords), a category dropdown, the template list, and a preview area. Show a loading state, an empty state and an error state.
9. Preview: fetch the template HTML, clean it with DOMPurify, and show it in a sandboxed `iframe` (`srcdoc`, scripts not allowed).

*Phase 4: Insert logic (depends on Phase 3)*

10. Take the `<body>` content, clean it with DOMPurify, and turn relative `img src` paths into full URLs.
11. Check whether the message is HTML or plain text with `item.body.getTypeAsync`.
    - HTML: insert with `item.body.setSelectedDataAsync(html, { coercionType: Html })`.
    - Plain text: insert a text-only version and show a notice.
12. Show success or failure as a toast message.
13. Unit tests (Jest) for: template-list parsing, filtering, URL conversion, body extraction and cleaning.

*Phase 5: GitHub Actions and Pages (depends on Phases 2–4)*

14. `.github/workflows/ci.yml` runs on every pull request using GitHub-hosted runners: set up the required Node.js and PowerShell environments, install dependencies there, then run Pester, PSScriptAnalyzer, the publish script in check-only mode, `npm run lint`, `npm test` and `npm run validate`. No runner or tool installation is required on the corporate PC.
15. `.github/workflows/deploy.yml` runs on push to `main`:
    - sets up Node.js and PowerShell and installs dependencies on a GitHub-hosted runner;
    - repeats the checks above;
    - builds the add-in with the base URL and build id;
    - copies `templates/` into the build output and generates `index.json` there;
    - deploys with `actions/upload-pages-artifact` and `actions/deploy-pages`.
16. Site layout under `https://<org>.github.io/<repo>/`: the add-in at the top level (`taskpane.html`, assets) and `templates/` for the HTML, images and `index.json`.
17. Protect `main` (require pull requests and passing CI) so a broken template can't reach users.

*Phase 6: Documentation for users and maintainers*

18. The root `README.md` is a detailed, beginner-friendly guide for someone unfamiliar with GitHub or the codebase. Use plain language, define GitHub terms the first time they appear, and include the following:
    - what the add-in does, what it does not do, how templates are hosted, and a simple overview of the request/data flow from Outlook to GitHub Pages;
    - a map of the repository, explaining the purpose of every important file and folder, including the manifest, task pane source, template files, publish script, tests, and GitHub Actions workflows;
    - the corporate-PC workflow: which source files can be prepared or edited without installing tools, which actions happen on GitHub's hosted runners, and which optional steps require the later unrestricted computer;
    - step-by-step GitHub instructions for creating or receiving the repository, understanding the repository page, editing through the GitHub web interface when appropriate, creating a branch, committing changes, opening a pull request, reviewing CI results, fixing failures, merging, and locating a deployment;
    - the first-time setup checklist for repository visibility and its privacy implications, enabling GitHub Pages through Actions, required Actions permissions, and the admin steps needed to upload and assign the manifest;
    - how to add or edit a template, including metadata, body content, images, validation rules, and the exact review/publish workflow;
    - a file-by-file explanation of how search, category filtering, preview, sanitization, and insertion into Outlook work, plus where to make common changes;
    - troubleshooting for failed Actions, missing templates/images, stale Pages content, manifest/deployment delays, and Outlook client differences;
    - optional setup and test commands for the unrestricted computer, clearly marked as optional and never prerequisites for working on this corporate PC.
19. `docs/deployment.md` is the concise admin runbook and covers:
    - enabling Pages with the source set to GitHub Actions (the site appears at `https://<org>.github.io/<repo>/`);
    - how IT adds or edits a template: branch, pull request, review, merge, then live within about 10 minutes;
    - uploading the manifest in Integrated Apps and assigning it to users or groups;
    - that Outlook can take up to 24 hours to pick up a new deployment or manifest change.
20. Keep the README and runbook consistent with the implemented repository structure and workflow. Include concrete paths, UI labels, and commands only after those names and steps are established; avoid assuming GitHub permissions or organization settings that the reader may not have.

**Files and folders to create**
- `README.md`: beginner-friendly GitHub, workflow, setup, troubleshooting, and codebase guide
- `addin/`: the scaffold, `manifest.json`, `src/taskpane/*` and Jest tests
- `templates/`: sample `*.html` files and `images/`
- `scripts/Publish-TemplateIndex.ps1`
- `tests/Publish-TemplateIndex.Tests.ps1`
- `.github/workflows/ci.yml` and `.github/workflows/deploy.yml`
- `docs/deployment.md`

**Verification**
1. On this corporate PC, verify the source/configuration files are present and review the diff; do not run package installation, builds, tests, or sideloading here.
2. A GitHub Actions pull-request run executes `npm run validate` on the manifest, `npm run lint`, `npm test` (Jest), `Invoke-Pester tests/`, and PSScriptAnalyzer on `scripts/` using hosted runners.
3. After moving the repository to an unrestricted computer, optionally install dependencies with `npm ci` and run the same checks locally. Sideload with `npm start` there and test this matrix in classic Outlook, new Outlook and Outlook on the web:
   - new message, reply and forward;
   - HTML and plain-text messages;
   - inserting at the cursor in the middle of existing text;
   - images showing correctly for the recipient;
   - the task pane staying pinned in classic Outlook.
4. A test pull request runs CI green. A merge to `main` deploys, and the site loads over HTTPS at `https://<org>.github.io/<repo>/`.
5. Once an administrator is ready to deploy, assign the add-in to a pilot group through Integrated Apps and confirm the button appears in all three clients.

**Decisions**
- No SharePoint, no Entra app registration and no sign-in. Templates and images are readable by anyone who has the URL.
- Hosting is GitHub Pages, so there's no server to maintain. If the repo is public, its source and templates are public too; never put internal-only content in templates.
- Publishing happens only through Git: every change is reviewed, tracked and can be rolled back.
- Users see templates by category; there's no per-group visibility.
- The add-in and templates share one domain, so no cross-origin (CORS) setup is needed.
- Images use public HTTPS URLs. Some recipients' email clients block remote images until the reader allows them.
- Templates insert at the cursor only. Not included: subject, recipients, attachments, placeholders, Mac and mobile.
- Not included in this repo: template authoring tools (IT edits the HTML directly).

**Further considerations**
1. **Placeholders later:** auto-filling things like `{{DisplayName}}` from Outlook's user profile needs no app registration and would be easy to add in a later version. Leave it out of v1.
2. **Public or private repo:** GitHub Team or Enterprise can publish a public Pages site from a private repo, which keeps the source and history private. Don't use Enterprise's "private Pages" (sign-in required), because Outlook can't sign in to load the add-in. On GitHub Free, the repo must be public.
3. **GitHub's terms:** Pages isn't meant as free hosting for a commercial service. An internal company tool is usually fine, but confirm with whoever handles licensing or compliance.
4. **Custom domain later:** once you can add a DNS CNAME record (e.g. `templates.company.com` pointing to `<org>.github.io`), switch to it. That means changing the build setting, re-uploading the manifest in Integrated Apps and keeping the `github.io` address working for a while. Images in emails already sent use the old `github.io` URLs, and they break if the repo is renamed or moved, so avoid renaming the repo or org.
