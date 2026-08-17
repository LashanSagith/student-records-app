/**
 * Microsoft sign-in + SharePoint configuration.
 *
 * CLIENT_ID and TENANT_ID are the Azure App Registration identifiers.
 * SHAREPOINT_SITE_URL must be the URL of the Microsoft 365 SharePoint site
 * whose default Documents library will hold the shared student database.
 */
export const CLIENT_ID = "8f3b1254-bc33-44fa-abc9-d40a5d6f6908";
export const TENANT_ID = "58bcd04b-6ab3-4a9e-93b9-2eb08627c630";

// Example: https://yourtenant.sharepoint.com/sites/StudentRecords
export const SHAREPOINT_SITE_URL = "https://icbt-my.sharepoint.com/";

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
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

// Delegated permissions requested at sign-in.
// User.Read = signed-in user's name/email for the audit log.
// Sites.ReadWrite.All = read/write the shared SharePoint document library.
export const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"],
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};
