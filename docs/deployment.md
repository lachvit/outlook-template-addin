# Deployment runbook (admins)

Short checklist for whoever runs this add-in. The [README](../README.md) explains each step in more detail, and [production-setup.md](production-setup.md) is the complete, ordered walkthrough for a new tenant.

Throughout, `<org>` is your GitHub organization (or user) name, `<repo>` the repository name, and `<site-url>` the SharePoint site for the templates (for example `https://contoso.sharepoint.com/sites/EmailTemplates`).

Nothing in the code is tied to one tenant: every tenant-specific value is a build setting. To move to another tenant, repeat steps 2-5 there and update the repository variables.

## 1. One-time GitHub setup

1. **Repository visibility.** On GitHub Free the repository must be **public** to use Pages, which makes the source code public. On GitHub Team/Enterprise a private repository can publish a **public** Pages site. Do **not** turn on Enterprise "private Pages" (sign-in required), because Outlook can't sign in to load the add-in. The templates themselves are in SharePoint and are never published.
2. **Enable Pages.** In the repository, go to **Settings > Pages > Build and deployment > Source** and choose **GitHub Actions**.
3. **Allow Actions.** In **Settings > Actions > General**, make sure Actions are allowed for this repository. If your organization restricts actions, allow `actions/*` (checkout, setup-node, configure-pages, upload-artifact, upload-pages-artifact, deploy-pages).
4. **Protect `main`.** In **Settings > Branches** (or **Settings > Rules > Rulesets**), add a rule for `main` that:
   - requires a pull request before merging (with at least 1 approval);
   - requires these status checks to pass: **PowerShell** and **Add-in** (they appear after CI has run once).

Note the add-in URL: `https://<org>.github.io/<repo>/` (or your custom domain). You need it for the app registration.

## 2. SharePoint site and library

1. Create (or choose) a SharePoint site for the templates, for example a **Team site** or **Communication site** called *Email Templates*.
2. Create the library and its columns. Either:
   - run the script (site owner or SharePoint admin, needs the `Microsoft.Graph.Authentication` module):

     ```powershell
     ./scripts/Initialize-TemplateLibrary.ps1 -SiteUrl <site-url>
     # optional: also upload the sample templates
     ./scripts/Initialize-TemplateLibrary.ps1 -SiteUrl <site-url> -SampleTemplatesPath samples/templates
     ```

   - or create them by hand: a document library **Email Templates** with these columns (the *internal* name matters, so create each column with that name first and rename the display name afterwards if you want):

     | Internal name | Type | Notes |
     | --- | --- | --- |
     | `TemplateCategory` | Choice (drop-down) | Your categories, e.g. General, HR, Sales |
     | `TemplateDescription` | Multiple lines of text, plain text | Optional |
     | `TemplateKeywords` | Single line of text | Optional, comma-separated |

     The built-in **Title** column is the template name.
3. **Permissions:** template **authors** need *Edit* (site Members), **everyone who uses the add-in** needs *Read* (site Visitors, e.g. *Everyone except external users*).
4. Recommended library settings (**Library settings > Versioning settings**): keep **version history** on, so you can restore a template. Turn on **Require content approval** if new or edited templates should be approved before users see them: readers only see approved (major) versions.
5. Graph-created columns aren't added to the default view. In the library, use **Add column > Show/hide columns** to show *Template category*, *Template description* and *Template keywords*.

## 3. Entra ID app registration

Follow [app-registration.md](app-registration.md). In short:

1. Register a single-tenant app.
2. Add SPA redirect URIs `brk-multihub://<org>.github.io` and `https://<org>.github.io/<repo>/auth.html`.
3. Add the delegated Microsoft Graph permission **Sites.Selected** and **grant admin consent**.
4. Grant the app **read** access to the site:

   ```powershell
   ./scripts/Grant-TemplateSiteAccess.ps1 -SiteUrl <site-url> -ClientId <client-id>
   ```

5. In **Enterprise applications**, decide whether **Assignment required** is on and set **Visible to users** to **No**.

## 4. Build settings and first deployment

1. In **Settings > Secrets and variables > Actions > Variables**, add repository variables:

   | Variable | Value |
   | --- | --- |
   | `AUTH_CLIENT_ID` | Application (client) ID |
   | `AUTH_TENANT_ID` | Directory (tenant) ID |
   | `SHAREPOINT_SITE_URL` | `<site-url>` |
   | `SHAREPOINT_LIBRARY` | Only if the library isn't called `Email Templates` (e.g. `Documents`) |
   | `SHAREPOINT_FOLDER` | Only if the templates are in a folder of the library (e.g. `Email`) |
   | `ADDIN_BASE_URL` | Only for a custom domain |

2. Open **Actions > Deploy > Run workflow** (or merge any pull request to `main`). When the run is green, the task pane is at `https://<org>.github.io/<repo>/taskpane.html` and the sign-in redirect page at `.../auth.html`.

The build fails with a clear message if a variable is missing or malformed.

## 5. Deploying the add-in to users (Microsoft 365 admin center)

1. Open the latest green **Deploy** run under **Actions** and download the **addin-manifests** artifact. It contains:
   - `manifest-unified.zip`: the unified (JSON) manifest package. Use this one by default.
   - `manifest.xml`: the older XML manifest. Use it only if the admin center or a client won't accept the unified manifest.
   Both already contain your real `https://<org>.github.io/<repo>/` URLs.
2. Go to the [Microsoft 365 admin center](https://admin.microsoft.com) > **Settings > Integrated apps > Upload custom apps**.
3. Choose the app type **Office Add-in**, then upload the manifest file you picked.
4. Assign the add-in to a **pilot group** first, then to more users or groups once it works. If the enterprise app has **Assignment required = Yes**, assign the same group there.
5. Finish the wizard and accept the permissions ("read and write the current item").

Deploy only **one** manifest. They share the same add-in ID.

**Timing:** Outlook can take **up to 24 hours** to show a new deployment or a manifest change. Users may need to restart Outlook.

## 6. Adding or editing a template

Template changes happen **only in SharePoint**: no pull request, build or deployment.

1. Upload or edit an `.html` file in the library.
2. Fill in **Title**, **Template category** and optionally **Template description** and **Template keywords**.
3. If content approval is on, approve it.

Users see the change the next time they open the task pane. Restore an older version with **Version history**.

## 7. When you need a new manifest

Template changes **never** require a new manifest. Upload a new manifest only when:

- you change names, icons, buttons or permissions in `addin/manifest.json` / `addin/manifest.xml` (also raise the `version`);
- the site URL changes (new repository name, organization or custom domain).

To update, upload the new file over the existing app in **Integrated apps**, then allow up to 24 hours.

Changing a **repository variable** (for example a new SharePoint site) only needs a new **Deploy** run, not a new manifest.

## 8. Moving to a custom domain later

1. Add a DNS CNAME record, for example `templates.contoso.com` pointing to `<org>.github.io`.
2. In **Settings > Pages**, enter the custom domain and turn on **Enforce HTTPS**.
3. Add a repository variable `ADDIN_BASE_URL` = `https://templates.contoso.com/`.
4. In the app registration, add the SPA redirect URIs `brk-multihub://templates.contoso.com` and `https://templates.contoso.com/auth.html`.
5. Run **Deploy**, download the new manifests and update the app in Integrated Apps.
