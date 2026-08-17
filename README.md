# Students Master Database — SharePoint Edition

A Microsoft 365 / GitHub Pages student master database with the same My Tasks-style UI, Microsoft sign-in, Excel import/export, edit history, and **one shared dataset stored in a SharePoint document library**.

## Why this version fixes the "Couldn't save" problem

The previous version stored `student-records-data.json` in each user's personal OneDrive App Folder. That meant different staff members could end up with different copies.

This version stores one JSON file in the configured SharePoint site's default Documents library:

`Documents/Student Records App/student-records-data.json`

All authorised staff members therefore read and write the same database.

## 1. Azure App Registration

1. Open https://portal.azure.com/
2. Microsoft Entra ID → App registrations → open your existing Student Records App registration (or create one).
3. Authentication → Single-page application (SPA).
4. Add this redirect URI exactly:

`https://LashanSagith.github.io/student-records-app/`

5. API permissions → Microsoft Graph → Delegated permissions:
   - `User.Read`
   - `Sites.ReadWrite.All`
6. If your organisation blocks user consent, an admin must click **Grant admin consent**.
7. Do **not** create a client secret. This is a browser SPA.

## 2. Configure the SharePoint site

Create or choose the Microsoft 365 SharePoint site that should own the database. The signed-in users must have permission to edit the site's Documents library.

Open `src/authConfig.js` and change:

```js
export const SHAREPOINT_SITE_URL = "https://YOURTENANT.sharepoint.com/sites/StudentRecords";
```

Example:

```js
export const SHAREPOINT_SITE_URL = "https://icbtcampus.sharepoint.com/sites/Engineering";
```

The app uses the site's **default Documents library**. On first save it automatically creates:

`Student Records App/student-records-data.json`

## 3. GitHub Pages

The repository already contains `.github/workflows/deploy.yml`.

1. Push/commit the project files to the `main` branch.
2. GitHub → Settings → Pages → Source → **GitHub Actions**.
3. Open Actions and wait for the deployment workflow to finish.
4. Open:

`https://LashanSagith.github.io/student-records-app/`

5. Make sure the same URL with the trailing `/` exists in Azure → Authentication → SPA redirect URIs.

## 4. Important permissions note

`Sites.ReadWrite.All` is a delegated Microsoft Graph permission. It acts as the signed-in user's access, so a user still needs permission to the SharePoint site/library.

For stricter enterprise security, this can later be changed to `Sites.Selected` with an explicit site-level permission assignment.

## 5. Local development

```bash
npm install
npm run dev
```

For local sign-in, add:

`http://localhost:5173/`

to Azure → Authentication → SPA redirect URIs.

## Data and audit behaviour

- Add/edit/delete/import changes are saved to the shared SharePoint JSON file.
- The app keeps the audit history in the same JSON file.
- SharePoint/OneDrive version history can provide an additional rollback mechanism depending on your Microsoft 365 library settings.
- The browser updates immediately and retries a failed save from the sync-status button.
