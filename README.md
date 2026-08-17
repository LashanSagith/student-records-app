# Students Master Database

A standalone web app (not a Claude artifact) with:
- **Sign in with Microsoft** — only people in your own Microsoft 365 organisation can get in.
- **Automatic OneDrive backup** — every change is saved to a dedicated app folder in the signed-in user's own OneDrive (`Apps/Student Records App`), so there's no shared server to secure or pay for.
- **Edit history log** — every add / edit / delete is timestamped with who did it.
- **Excel import/export** — unchanged from before.

You need to do two one-time setup steps before this works: **(1) Azure App Registration** and **(2) GitHub Pages hosting**. Both are below, in order. You said you have both Microsoft 365/Azure admin access and a GitHub account, so you can do all of this yourself.

---

## 1. Azure App Registration (do this first)

This tells Microsoft "this exact app, hosted at this exact address, is allowed to ask our staff to sign in and use their OneDrive." It's a one-time, ~10 minute setup.

1. Go to **[portal.azure.com](https://portal.azure.com)** and sign in with your admin account.
2. Search for **"App registrations"** → click **+ New registration**.
3. Fill in:
   - **Name**: `Student Records App` (this name becomes the OneDrive folder name each user gets, so keep it recognisable)
   - **Supported account types**: choose **"Accounts in this organizational directory only (Single tenant)"** — this is what restricts sign-in to your organisation only.
   - **Redirect URI**: platform = **Single-page application (SPA)**, URI = `https://<your-github-username>.github.io/student-records-app/` (see step 2 for the exact URL — you can come back and fix this after you know it).
4. Click **Register**.
5. On the app's **Overview** page, copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**
   Paste both into `src/authConfig.js` in this project (`CLIENT_ID` and `TENANT_ID`).
6. Go to **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions** → search for and add:
   - `User.Read` (usually already there by default)
   - `Files.ReadWrite.AppFolder`
   Neither of these needs Global Admin "grant consent for organisation" — each staff member can consent for themselves the first time they sign in. (If your tenant has a policy blocking user consent, click **Grant admin consent** on this page once, as the admin.)
7. Go to **Authentication** → confirm the SPA redirect URI is there → under **Implicit grant and hybrid flows**, leave both checkboxes **unchecked** (MSAL.js uses the more secure auth code flow, not implicit flow) → **Save**.

That's it on the Azure side. Because you chose "single tenant," anyone outside your organisation who tries to open the app will be rejected by Microsoft at sign-in — the app code never has to enforce that itself.

---

## 2. Host it on GitHub Pages

There are two ways to publish — pick whichever suits you. Option A needs no terminal at all.

### Option A — Upload via the GitHub website (no terminal, no Git, no npm needed on your PC)

This project already includes a GitHub Actions workflow (`.github/workflows/deploy.yml`) that automatically builds and publishes the app every time you upload/change files — GitHub's own servers run `npm install` and `npm run build` for you.

1. Create a new GitHub repository — name it anything you like, e.g. `student-records-app`. You no longer need to match it to anything in `vite.config.js` (the build uses relative paths, so it works under any repo name).
2. On the repo's page, click **Add file → Upload files**.
3. Drag the entire contents of the extracted `student-records-app` folder into the browser window (all the files and folders — `src/`, `.github/`, `package.json`, `index.html`, etc. — but not a `node_modules` or `dist` folder, since those shouldn't exist yet anyway). Most browsers preserve folder structure when you drag a folder in; if yours doesn't, upload folder-by-folder.
4. Scroll down, click **Commit changes**.
5. Go to **Settings → Pages**, and under **Build and deployment → Source**, choose **GitHub Actions** (not "Deploy from a branch").
6. Go to the **Actions** tab — you'll see a workflow run start automatically. Wait for it to finish (green checkmark, usually 1-2 minutes).
7. Your app is now live at `https://<your-github-username>.github.io/student-records-app/`.
8. Go back to Azure Portal → your app registration → **Authentication**, and make sure that exact URL (with the trailing slash) is listed as a redirect URI.

Every time you edit a file and re-upload it the same way (**Add file → Upload files**, then Commit), the Actions workflow re-runs automatically and republishes the update — no local build step ever needed.

### Option B — Push from your own computer with Git + npm

1. Create the GitHub repo as in step 1 above.
2. Push all the files in this project to that repo (`git init`, `git add .`, `git commit`, `git remote add origin ...`, `git push`).
3. Same as Option A: **Settings → Pages → Source → GitHub Actions**. The included workflow will build and deploy automatically on every push, so you don't even need `npm run deploy` locally — though it's still available if you'd rather deploy straight from your own machine:
   ```bash
   npm install
   npm run build
   npm run deploy
   ```
4. Confirm the redirect URI in Azure Portal as in Option A step 8.

---

## How the OneDrive backup actually works

- The app asks each signed-in user for permission to use **one folder only**: a hidden folder Microsoft creates automatically at `OneDrive → Apps → Student Records App`. The app can never see or touch anything else in anyone's OneDrive — this is the safest permission scope Microsoft Graph offers for this kind of app.
- Every add, edit, delete, or import saves the current data to a single file in that folder: `student-records-data.json`. Saves are automatic (debounced ~1 second after your last change) — there's no manual "backup" button to remember.
- **Important**: because the data lives in *each signed-in user's own* OneDrive app folder, different staff members signing in will each get their **own separate copy** unless you all sign in as the same shared account, or you designate one person's OneDrive as the "master" and everyone else uses read-only access. If you want one single shared dataset that the whole team edits together (most likely what you want for a shared master database), the cleanest fix is to point the app at a **SharePoint site / Team's shared document library** instead of personal OneDrive — that's a small change to `graphService.js` (swapping the `/me/drive/special/approot` endpoint for a `/sites/{site-id}/drive` endpoint) and needs a `Sites.ReadWrite.All` or `Files.ReadWrite.All` permission instead of `Files.ReadWrite.AppFolder`. Happy to make that change if that's what you actually need — let me know.
- **Edit history**: shown in-app via the "Edit history" button (who changed what, and when). Separately, OneDrive itself automatically keeps its own version history for the JSON file (right-click the file in OneDrive's web UI → **Version history**) — so even if the in-app log were ever lost, you could still roll back to an earlier saved version of the data itself.

---

## Local development

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173`. Sign-in will work locally too, as long as you also add `http://localhost:5173` as a redirect URI in the Azure app registration (Authentication → + Add a platform → Single-page application).

---

## Project structure

```
index.html
vite.config.js          ← builds with relative asset paths (works under any repo name)
src/
  authConfig.js          ← paste your Azure CLIENT_ID / TENANT_ID here
  graphService.js         ← OneDrive load/save via Microsoft Graph
  main.jsx                ← MSAL provider setup
  App.jsx                 ← login screen / signed-in app switch
  StudentRecords.jsx      ← the actual student database UI
```
