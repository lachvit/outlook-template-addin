# Publishing this project to public GitHub

A step-by-step guide for someone who has not used public GitHub before. It takes you from a folder on your PC to a live GitHub Pages site that the Outlook add-in can load.

Throughout, `<username>` is your GitHub user name (or organization name) and `<repo>` is the repository name you choose, for example `outlook-template-addin`.

## Contents

1. [Before you start: decide and check](#1-before-you-start-decide-and-check)
2. [Clean up the project before it becomes public](#2-clean-up-the-project-before-it-becomes-public)
3. [Create your GitHub account](#3-create-your-github-account)
4. [Create an empty repository on GitHub](#4-create-an-empty-repository-on-github)
5. [Push the project to GitHub](#5-push-the-project-to-github)
6. [Turn on GitHub Pages](#6-turn-on-github-pages)
7. [Check Actions settings and run the first deployment](#7-check-actions-settings-and-run-the-first-deployment)
8. [Verify the published site](#8-verify-the-published-site)
9. [Protect the main branch](#9-protect-the-main-branch)
10. [Next steps](#10-next-steps)
11. [Troubleshooting](#11-troubleshooting)

---

## 1. Before you start: decide and check

**Public means public.** Once the repository is public, anyone on the internet can read every file **and the full history** of every commit. Deleting something later does not remove it from history, and it may already have been copied.

Check these first:

- **Company policy.** If this is work you did for your employer, confirm that you are allowed to publish it. This is usually a quick question to your manager, security or compliance team. Also check which account to use: a company-managed GitHub organization is often required for work code, and a personal account may be forbidden.
- **No confidential content.** Templates, images, names, e-mail addresses and links in the project must all be safe for the public to see. The published templates are readable by anyone with the URL anyway.
- **Why public?** On the free GitHub plan, GitHub Pages only works for public repositories. The add-in needs Pages (it is the web hosting), so public is the simple route. Paid plans (Team, Enterprise) can keep the source private and still publish a public Pages site. See [Part 6 of the README](../README.md#part-6-first-time-setup-checklist).
- **Network access.** Your PC must be able to reach `github.com`. Locked-down corporate PCs sometimes block it or require a proxy. If `git push` later fails to connect, ask IT, or do the upload from an unrestricted computer.

Install **Git for Windows** from <https://git-scm.com/download/win> if `git --version` does not work in PowerShell. Use the default options.

## 2. Clean up the project before it becomes public

Do this **before** the first push. It is much easier than fixing it afterwards.

Open PowerShell in the project folder:

```powershell
cd C:\Development\Templates\Add-in
git status
```

### 2.1 Leave out the internal package lock file

`addin/package-lock.json` currently contains addresses of your company's internal package server (`artifactory.danskenet.net`). Publishing it would:

- reveal an internal host name;
- **break the GitHub build**, because GitHub's servers cannot reach that internal server and `npm ci` would fail.

Pick one option:

- **Simplest:** do not commit it. The workflows already handle this: if there is no `package-lock.json`, they run `npm install` instead.

  ```powershell
  Remove-Item addin\package-lock.json
  ```

- **Better for repeatable builds:** regenerate it against the public npm registry, then check it is clean.

  ```powershell
  cd addin
  Remove-Item package-lock.json
  npm install --package-lock-only --registry https://registry.npmjs.org/
  cd ..
  Select-String -Path addin\package-lock.json -Pattern "danskenet"
  ```

  The last command must print **nothing**. If it prints lines, delete the file and use the simplest option.

### 2.2 Remove stray files

These are untracked files that are probably not meant for the repository:

| File | What to do |
|---|---|
| `package-lock.json` (project root) | It is empty (`"packages": {}`). Delete it. |
| `Local testing.md` (project root) | A duplicate of `docs/local-testing.md`. Delete it, or keep only one of the two. |
| `icon.png` (project root) | Keep only if you want it. The add-in uses the icons in `addin/assets/`. |
| `plan-outlookTemplateAddin.prompt.md` | A planning note that is already committed. Delete it if you do not want it public. |

```powershell
Remove-Item package-lock.json
Remove-Item "Local testing.md"
```

### 2.3 Make the add-in yours

As noted in the README, in `addin/manifest.json` and `addin/manifest.xml`, replace the `Contoso IT` developer/provider name. Replace the sample icons and logo if they are not meant to be public.

### 2.4 Search for secrets and internal details

Look for anything you would not put on a billboard: passwords, tokens, API keys, internal URLs, colleague e-mail addresses.

```powershell
Select-String -Path (git ls-files) -Pattern "password|secret|token|apikey|danskenet|@danskebank" -List
```

Review each match. No output means nothing was found. (Text like `token` in source code can be harmless; judge each hit.)

### 2.5 Choose which e-mail address appears in your commits

Every Git commit stores the author name and e-mail, and **they become public**. Your Git is currently configured with your work e-mail address.

GitHub offers a private "no-reply" address so you do not expose your real one:

1. Sign in on GitHub (see section 3) and open **Settings > Emails**.
2. Tick **Keep my email addresses private** and **Block command line pushes that expose my email**.
3. Copy the no-reply address shown at the top. It looks like `12345678+<username>@users.noreply.github.com`.
4. Set it for this repository only:

   ```powershell
   git config user.name "Your Name"
   git config user.email "12345678+<username>@users.noreply.github.com"
   ```

The existing commit (`First`) was made with your work e-mail. Since the history is a single commit, the easiest fix is to rewrite it after you finish this section:

```powershell
git add -A
git commit --amend --reset-author -m "Initial commit"
```

`--reset-author` replaces the author with the configured no-reply address. Only do this **before** you push for the first time.

If you instead want a separate commit for the cleanup, skip `--amend`, but then the first commit still shows your work e-mail publicly.

### 2.6 Add a license (strongly recommended)

Without a license file, nobody else is legally allowed to use or copy the code, even though they can see it. `addin/package.json` currently says `"license": "UNLICENSED"`.

- To allow reuse, choose a license (the MIT license is short and common) at <https://choosealicense.com/>, then add a `LICENSE` file in the project root and update `"license"` in `addin/package.json`.
- If the company wants to keep all rights, leave it as is and mention it in the README.

Confirm the choice with your company if this is work code.

### 2.7 Commit the cleaned-up state

```powershell
git status
git add -A
git commit -m "Prepare repository for publishing"
```

(Skip this commit if you already did `git commit --amend` in 2.5 and nothing changed since.)

Check the result with `git status`. It should say `nothing to commit, working tree clean`. The `.gitignore` already keeps `node_modules/`, `addin/dist/` and other generated files out.

### 2.8 Moving the project to another computer

If you will push from a different PC, copy the project there after sections 2.1 to 2.6 are done. Check that your company allows the transfer method you use (USB stick, personal cloud storage, e-mail). Work files often may only travel by approved means.

**What to copy** (everything inside `C:\Development\Templates\Add-in`):

| Copy | Notes |
|---|---|
| `.github/` | **Essential.** Holds the CI and Deploy workflows. It starts with a dot, so some tools skip it. Without it nothing gets built or published. |
| `addin/` | The add-in source, **without** `addin/node_modules/` and `addin/dist/`. Includes `src/`, `tests/`, `assets/`, `manifest.json`, `manifest.xml`, `package.json`, `webpack.config.js`, `tsconfig*.json`, `jest.config.js`, `eslint.config.mjs`. Include `package-lock.json` only if you regenerated it in section 2.1. |
| `templates/` | The email templates and `images/`. `templates/index.json` is generated, so it is optional. |
| `scripts/` | `Publish-TemplateIndex.ps1`. |
| `tests/` | `Publish-TemplateIndex.Tests.ps1`. |
| `docs/` | Documentation, including this guide. |
| `README.md` | |
| `.gitignore`, `.gitattributes` | Dot files, easy to miss. `.gitignore` keeps generated files out of the repository. |
| `LICENSE` | If you added one in 2.6. |

**What not to copy:**

- `node_modules/` (in `addin/`): large and rebuilt by `npm install`. It is not needed for publishing, because GitHub builds the project itself.
- `addin/dist/`, `coverage/`, `*.zip`: generated output.
- `.git/`: see below.
- Any files you deleted in 2.2 (`Local testing.md`, root `package-lock.json`, and so on).

**With or without `.git/`?**

- **Recommended: without `.git/`.** The new PC starts a fresh history, so your work e-mail in the old commit (2.5) is never published. Skip 2.5's `--amend` step in that case.
- With `.git/`: the full history moves along, including the commit with your work e-mail. Only choose this if you have fixed that as described in 2.5.

**Copy safely with PowerShell.** `robocopy` copies everything, including dot folders, and skips what you list. Run it on the old PC (change the destination to your USB drive or share):

```powershell
robocopy C:\Development\Templates\Add-in D:\Add-in-copy /E /XD .git node_modules dist coverage /XF *.zip "Local testing.md"
```

(`/E` includes subfolders, `/XD` excludes folders, `/XF` excludes files. Robocopy prints exit codes below 8 for success.)

Alternatively, with Git only (no `.git` folder), export tracked files as a zip. This leaves out ignored and untracked files, so first make sure everything wanted is committed (`git status` must be clean):

```powershell
git archive --format=zip -o C:\Temp\add-in.zip HEAD
```

**On the new computer:**

1. Install Git for Windows. Node.js is **not** needed just to publish.
2. Put the files in a folder, for example `C:\Development\Add-in`, and check that `.github` is there (in File Explorer, turn on **View > Show > Hidden items** if you cannot see it).
3. Create the repository history and set your identity (use the no-reply address from 2.5):

   ```powershell
   cd C:\Development\Add-in
   git init -b main
   git config user.name "Your Name"
   git config user.email "12345678+<username>@users.noreply.github.com"
   git add -A
   git status
   git commit -m "Initial commit"
   ```

   Read the `git status` list. It should show your project files and no `node_modules`.
4. Continue with section 3, or section 4 if you already have an account. In section 5, skip `git branch -M main` (already done) and start at `git remote add origin ...`.

## 3. Create your GitHub account

Skip this if you already have one.

1. Go to <https://github.com/signup>.
2. Enter an e-mail address, a password and a user name. The user name becomes part of the web address of your site (`https://<username>.github.io/<repo>/`), and **changing it later breaks the add-in's URLs**. Choose it carefully. If this is a company project, consider creating an **organization** (below) so the address does not depend on one person.
3. Verify the e-mail address with the code GitHub sends you.
4. **Turn on two-factor authentication:** **profile picture > Settings > Password and authentication > Two-factor authentication**. Use an authenticator app. This is important for any public account.

### Optional: create an organization

An organization is a shared owner for repositories, useful if several people will maintain the project.

1. Click the **+** at the top right of GitHub and choose **New organization**.
2. Choose the **Free** plan, give it a name (this becomes the `<username>` in URLs) and a contact e-mail, and say whether it belongs to a personal account or a business.
3. Add the other maintainers later under **People**.

If you do this, use the organization name wherever this guide says `<username>`.

## 4. Create an empty repository on GitHub

1. Sign in and click the **+** at the top right, then **New repository**.
2. Fill in the form:
   - **Owner:** your account or organization.
   - **Repository name:** for example `outlook-template-addin`. Use lowercase letters, digits and hyphens. **Do not rename it later**: the add-in URL contains it, and emails already sent contain image links with it.
   - **Description:** one sentence, optional.
   - **Public** (selected).
   - **Do not** tick **Add a README file**, **Add .gitignore** or **Choose a license**. The project already has these, and extra files would make your first push conflict.
3. Click **Create repository**.

GitHub now shows a page with setup instructions and a URL such as `https://github.com/<username>/<repo>.git`. Copy that URL.

## 5. Push the project to GitHub

"Push" means uploading your local commits to GitHub.

1. In PowerShell, in the project folder, make sure the branch is called `main` and tell Git where GitHub is:

   ```powershell
   cd C:\Development\Templates\Add-in
   git branch -M main
   git remote add origin https://github.com/<username>/<repo>.git
   git remote -v
   ```

   `git remote -v` should list `origin` twice (fetch and push) with your URL.

2. Upload:

   ```powershell
   git push -u origin main
   ```

3. **Sign in when asked.** Git for Windows opens a browser window (Git Credential Manager). Choose **Sign in with your browser** and approve. If it asks for a password in the terminal instead, **your GitHub password does not work**. Use a personal access token:
   1. GitHub: **profile picture > Settings > Developer settings > Personal access tokens > Fine-grained tokens > Generate new token**.
   2. Name it, set a short expiry, under **Repository access** choose **Only select repositories** and pick this repo, and under **Permissions > Repository permissions** set **Contents** to **Read and write**.
   3. Generate it, copy it immediately (you cannot see it again) and paste it as the password.
   4. Never put a token in a file or commit it.

4. Refresh the repository page on GitHub. You should see all your files and the README displayed under them.

If you see `error: remote origin already exists`, run `git remote set-url origin https://github.com/<username>/<repo>.git` instead of `git remote add`.

> **Alternative (when the code lives in another Git host):** GitHub can import from a URL at <https://github.com/new/import>. For a project that is already on your PC, the push above is simpler.

## 6. Turn on GitHub Pages

Pages is GitHub's free web hosting for static files. The Deploy workflow publishes the add-in and templates there.

1. In your repository, click **Settings** (top tab; you need admin rights on the repo, which you have if you created it).
2. In the left menu, click **Pages**.
3. Under **Build and deployment**, set **Source** to **GitHub Actions**. (Not "Deploy from a branch".)
4. Nothing else to save. The page shows the future address `https://<username>.github.io/<repo>/`.

## 7. Check Actions settings and run the first deployment

**Actions** are GitHub's automatic jobs. This project has two workflows in [.github/workflows/](../.github/workflows/): **CI** (checks every change) and **Deploy** (publishes after changes reach `main`).

1. **Settings > Actions > General**. Under **Actions permissions**, make sure **Allow all actions and reusable workflows** is selected, or that the GitHub-made actions listed in [README Part 6](../README.md#part-6-first-time-setup-checklist) are allowed. Under **Workflow permissions**, the default **Read repository contents** is enough.
2. Open the **Actions** tab. Because you pushed to `main`, a **Deploy** run may already be running or finished. If you see a message asking you to enable workflows, click the green button to enable them.
3. If nothing ran, click **Deploy** in the left list, then **Run workflow > Run workflow**.
4. Click the running entry to watch it. It goes through **Checks**, **Build site** and **Deploy**. It takes a few minutes. All green ticks means success.
5. The first time, GitHub may ask for approval of the `github-pages` environment, or may report that Pages is not enabled. If so, finish section 6 and click **Re-run all jobs**.

## 8. Verify the published site

Open each of these in a browser (replace the placeholders):

- `https://<username>.github.io/<repo>/templates/index.json`: should show a list of templates as JSON text.
- `https://<username>.github.io/<repo>/taskpane.html`: should show the add-in task pane (it may look empty or show a message outside Outlook, but it must not be a "404" page).

The first publish can take a minute or two before the site responds.

To get the manifest files for Outlook, open the finished **Deploy** run in the **Actions** tab and download the **addin-manifests** artifact at the bottom of the page. Uploading it to Microsoft 365 is described in [README Part 6, step 6](../README.md#part-6-first-time-setup-checklist) and [deployment.md](./deployment.md).

## 9. Protect the main branch

This stops accidental or unreviewed changes from being published, because every change to `main` goes live.

1. Run CI once so GitHub knows its check names: create a branch and a pull request with any small change (for example a typo fix), or open **Actions > CI** and run it if it is offered.
2. **Settings > Branches > Add branch protection rule** (or **Settings > Rules > Rulesets > New branch ruleset**).
3. Branch name pattern: `main`.
4. Tick **Require a pull request before merging** and require at least **1 approval**. If you are the only maintainer, you can leave approvals at 0, but keep the pull request requirement so CI runs first.
5. Tick **Require status checks to pass before merging** and add **Templates and PowerShell** and **Add-in**.
6. Save.

Also useful under **Settings > Code security** for a public repository: turn on **Dependabot alerts** and **Secret scanning** (and **Push protection**, which blocks accidental commits of passwords and tokens).

## 10. Next steps

- Add or edit templates by following [README Part 5](../README.md#part-5-adding-or-editing-a-template). The usual flow is: create a branch, edit, open a pull request, wait for green checks, merge. The site updates itself.
- Finish the Microsoft 365 steps (uploading the manifest and assigning users) from [README Part 6](../README.md#part-6-first-time-setup-checklist).
- If you later want your own web address instead of `github.io`, see [deployment.md](./deployment.md). Don't rename the repository or the account afterwards.

## 11. Troubleshooting

| Problem | Likely cause and fix |
|---|---|
| `git: command not found` / not recognized | Install Git for Windows and reopen PowerShell. |
| `Could not resolve host: github.com` or timeouts | Network or proxy blocks GitHub. Ask IT, or push from another network. |
| `Authentication failed` or `Permission denied` | You used your password. Use the browser sign-in or a personal access token (section 5). Check you are signed in as the owner of the repo. |
| `remote: error: GH007: Your push would publish a private email address` | You enabled "Block command line pushes that expose my email" and a commit uses your work e-mail. Redo section 2.5 (`git commit --amend --reset-author`), then push. |
| `failed to push some refs ... fetch first` | The new repo is not empty (a README or license was added on GitHub). Either delete the repo and recreate it empty, or run `git pull origin main --allow-unrelated-histories`, resolve conflicts, then push. |
| Deploy fails at **npm ci** | The lock file points to an internal registry. Fix as in section 2.1, commit and push. |
| Deploy fails: "Get Pages site failed" / "Pages not enabled" | Do section 6 (Source = GitHub Actions), then **Re-run all jobs**. |
| Pages shows 404 | Wait a few minutes. Check the Deploy run was green and the URL is exactly `https://<username>.github.io/<repo>/`, with the right capitalization of the repo name. |
| Actions tab shows no workflows | Check `.github/workflows/` was pushed. Enable workflows if GitHub asks. |
| Pages option is missing or says it needs a paid plan | Private repositories need a paid plan for Pages. Make the repository public (**Settings > General > Danger Zone > Change visibility**) or upgrade. |
| I pushed something secret | Treat it as leaked. Revoke or change the secret first. Removing it from history is a separate, harder job (`git filter-repo`) and does not undo any copies already made. |
