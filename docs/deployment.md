# Deployment runbook (admins)

Short checklist for whoever runs this add-in. The [README](../README.md) explains each step in more detail.

Throughout, `<org>` is your GitHub organization (or user) name and `<repo>` is the repository name.

## 1. One-time GitHub setup

1. **Repository visibility.** On GitHub Free the repository must be **public** to use Pages, which makes the source code and templates public. On GitHub Team/Enterprise a private repository can publish a **public** Pages site. Do **not** turn on Enterprise "private Pages" (sign-in required), because Outlook can't sign in to load the add-in.
2. **Enable Pages.** In the repository, go to **Settings > Pages > Build and deployment > Source** and choose **GitHub Actions**.
3. **Allow Actions.** In **Settings > Actions > General**, make sure Actions are allowed for this repository. If your organization restricts actions, allow `actions/*` (checkout, setup-node, configure-pages, upload-artifact, upload-pages-artifact, deploy-pages).
4. **Protect `main`.** In **Settings > Branches** (or **Settings > Rules > Rulesets**), add a rule for `main` that:
   - requires a pull request before merging (with at least 1 approval);
   - requires these status checks to pass: **Templates and PowerShell** and **Add-in** (they appear after CI has run once).
5. **First deployment.** Open **Actions > Deploy > Run workflow** (or merge any pull request to `main`). When the run is green, the site is at `https://<org>.github.io/<repo>/` and the task pane at `https://<org>.github.io/<repo>/taskpane.html`.

## 2. Adding or editing a template

1. Create a branch, then add or edit an `.html` file in `templates/` (images go in `templates/images/`).
2. Open a pull request. The **CI** workflow validates every template.
3. When CI is green and someone has reviewed the change, merge.
4. **Deploy** runs automatically. Users see the change within about **10 minutes** of the deployment finishing. Each deployment loads fresh files, but GitHub Pages can keep serving the old site for a few minutes.

To roll back, revert the pull request on GitHub (**Revert** button on the merged pull request), then merge the revert.

## 3. Deploying the add-in to users (Microsoft 365 admin center)

1. Open the latest green **Deploy** run under **Actions** and download the **addin-manifests** artifact. It contains:
   - `manifest-unified.zip`: the unified (JSON) manifest package. Use this one by default.
   - `manifest.xml`: the older XML manifest. Use it only if the admin center or a client won't accept the unified manifest.
   Both already contain your real `https://<org>.github.io/<repo>/` URLs.
2. Go to the [Microsoft 365 admin center](https://admin.microsoft.com) > **Settings > Integrated apps > Upload custom apps**.
3. Choose the app type **Office Add-in**, then upload the manifest file you picked.
4. Assign the add-in to a **pilot group** first, then to more users or groups once it works.
5. Finish the wizard and accept the permissions ("read and write the current item").

Deploy only **one** manifest. They share the same add-in ID.

**Timing:** Outlook can take **up to 24 hours** to show a new deployment or a manifest change. Users may need to restart Outlook.

## 4. When you need a new manifest

Template changes **never** require a new manifest. Upload a new manifest only when:

- you change names, icons, buttons or permissions in `addin/manifest.json` / `addin/manifest.xml` (also raise the `version`);
- the site URL changes (new repository name, organization or custom domain).

To update, upload the new file over the existing app in **Integrated apps**, then allow up to 24 hours.

## 5. Moving to a custom domain later

1. Add a DNS CNAME record, for example `templates.contoso.com` pointing to `<org>.github.io`.
2. In **Settings > Pages**, enter the custom domain and turn on **Enforce HTTPS**.
3. In **Settings > Secrets and variables > Actions > Variables**, add a repository variable `ADDIN_BASE_URL` = `https://templates.contoso.com/`.
4. Run **Deploy**, download the new manifests and update the app in Integrated Apps.
5. Images in emails that were already sent still use the old `github.io` URLs. Keep that address working, and don't rename the repository or organization.
