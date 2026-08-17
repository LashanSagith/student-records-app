/**
 * Microsoft sign-in configuration.
 *
 * You MUST fill in CLIENT_ID and TENANT_ID after you register this app in
 * Azure Portal. See README.md → "Azure App Registration" for exact steps.
 *
 * Using `TENANT_ID` (instead of "common") means ONLY people in your own
 * Microsoft 365 organisation can sign in — nobody outside it, automatically.
 * That is your "admin control" requirement: it's enforced by Azure itself,
 * not by app code, so it can't be bypassed from the browser.
 */

// Paste the "Application (client) ID" from your App Registration's Overview page.
export const CLIENT_ID = "PASTE_YOUR_AZURE_APPLICATION_CLIENT_ID_HERE";

// Paste the "Directory (tenant) ID" from the same page.
export const TENANT_ID = "PASTE_YOUR_AZURE_TENANT_ID_HERE";

// Must exactly match a "Single-page application" redirect URI you add in
// Azure Portal → App registration → Authentication. For GitHub Pages this
// is your Pages URL, e.g. https://yourorg.github.io/student-records-app/
export const REDIRECT_URI = window.location.origin + window.location.pathname;

export const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
    postLogoutRedirectUri: REDIRECT_URI,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    // localStorage (not sessionStorage) so a sign-in survives closing the
    // tab — same "stay signed in" feel as other Microsoft 365 apps.
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

/**
 * Scopes requested at sign-in.
 * - User.Read            → basic profile (name/email), so we can show who's signed in
 *                          and stamp their name on the edit history log.
 * - Files.ReadWrite.AppFolder
 *                         → read/write access to ONE dedicated, hidden folder in the
 *                           signed-in user's OneDrive: "Apps/Student Records App".
 *                           The app can NEVER see or touch anything else in anyone's
 *                           OneDrive. This is the safest possible scope for this job
 *                           and normal users can consent to it themselves — no
 *                           Global Admin needed to click "Grant consent" (though your
 *                           tenant may still require it, see README).
 */
export const loginRequest = {
  scopes: ["User.Read", "Files.ReadWrite.AppFolder"],
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};
