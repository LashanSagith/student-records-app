import { loginRequest, SHAREPOINT_SITE_URL } from "./authConfig";

const DATA_FILE_NAME = "student-records-data.json";
const DATA_FOLDER_NAME = "Student Records App";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function assertSharePointConfigured() {
  if (!SHAREPOINT_SITE_URL || SHAREPOINT_SITE_URL.includes("YOURTENANT")) {
    throw new Error("SharePoint site URL is not configured in src/authConfig.js");
  }
}

function siteLookupPath() {
  const url = new URL(SHAREPOINT_SITE_URL);
  const path = url.pathname.replace(/\/+$/, "");
  return `/sites/${url.host}:${path}`;
}

async function graphFetch(path, token, options = {}) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });
  return res;
}

export async function getGraphToken(msalInstance, account) {
  const request = { ...loginRequest, account };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (silentError) {
    const result = await msalInstance.acquireTokenPopup(request);
    return result.accessToken;
  }
}

async function getSharePointDrive(token) {
  assertSharePointConfigured();

  const siteRes = await graphFetch(siteLookupPath(), token);
  if (!siteRes.ok) {
    const detail = await siteRes.text();
    throw new Error(`SharePoint site lookup failed (${siteRes.status}) ${detail}`);
  }
  const site = await siteRes.json();

  const driveRes = await graphFetch(`/sites/${encodeURIComponent(site.id)}/drive`, token);
  if (!driveRes.ok) {
    const detail = await driveRes.text();
    throw new Error(`SharePoint document library lookup failed (${driveRes.status}) ${detail}`);
  }
  const drive = await driveRes.json();
  return { site, drive };
}

async function ensureDataFolder(token, driveId) {
  const folderPath = `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(DATA_FOLDER_NAME)}`;
  const existing = await graphFetch(`${folderPath}:`, token);
  if (existing.ok) return;
  if (existing.status !== 404) {
    const detail = await existing.text();
    throw new Error(`SharePoint folder check failed (${existing.status}) ${detail}`);
  }

  const createRes = await graphFetch(`/drives/${encodeURIComponent(driveId)}/root/children`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: DATA_FOLDER_NAME,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    }),
  });

  if (!createRes.ok && createRes.status !== 409) {
    const detail = await createRes.text();
    throw new Error(`SharePoint folder creation failed (${createRes.status}) ${detail}`);
  }
}

function dataFilePath(driveId) {
  return `/drives/${encodeURIComponent(driveId)}/root:/${encodeURIComponent(DATA_FOLDER_NAME)}/${encodeURIComponent(DATA_FILE_NAME)}:/content`;
}

/** Load the shared student database from the configured SharePoint site. */
export async function loadFromSharePoint(msalInstance, account) {
  const token = await getGraphToken(msalInstance, account);
  const { drive } = await getSharePointDrive(token);
  const res = await graphFetch(dataFilePath(drive.id), token);
  if (res.status === 404) return null;
  if (!res.ok) {
    const detail = await res.text();
    throw new Error(`SharePoint load failed (${res.status}) ${detail}`);
  }
  return res.json();
}

/** Save the shared student database to the team's SharePoint document library. */
export async function saveToSharePoint(msalInstance, account, data) {
  const token = await getGraphToken(msalInstance, account);
  const { drive } = await getSharePointDrive(token);
  await ensureDataFolder(token, drive.id);

  const res = await graphFetch(dataFilePath(drive.id), token, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data, null, 2),
  });

  if (!res.ok) {
    const detail = await res.text();
    throw new Error(`SharePoint save failed (${res.status}) ${detail}`);
  }
  return res.json();
}

export async function getMyProfile(msalInstance, account) {
  const token = await getGraphToken(msalInstance, account);
  const res = await graphFetch("/me", token);
  if (!res.ok) throw new Error(`Profile fetch failed (${res.status})`);
  return res.json();
}
