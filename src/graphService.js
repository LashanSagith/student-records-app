import { loginRequest } from "./authConfig";

// Special "approot" folder = the hidden per-app folder Microsoft Graph
// creates automatically inside each user's OneDrive at
// "Apps/Student Records App" (folder name comes from the app's display
// name in Azure). Only this app (with this exact registration) can read
// or write inside it — not other apps, and not the user browsing OneDrive
// normally sees it mixed in with their other files.
const DATA_FILE_NAME = "student-records-data.json";
const DATA_FILE_PATH = `/me/drive/special/approot:/${DATA_FILE_NAME}:/content`;
const EXCEL_FILE_NAME = "student-records-backup.xlsx";
const EXCEL_FILE_PATH = `/me/drive/special/approot:/${EXCEL_FILE_NAME}:/content`;
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Get a fresh access token for Graph calls. Tries silently first (no popup)
 * and only falls back to an interactive popup if the silent attempt fails
 * (e.g. the token expired and needs a fresh consent/login prompt).
 */
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

/**
 * Load { students, auditLog } from the user's OneDrive app folder.
 * Returns null if the file doesn't exist yet (first-ever run for this user).
 */
export async function loadFromOneDrive(msalInstance, account) {
  const token = await getGraphToken(msalInstance, account);
  const res = await fetch(`${GRAPH_BASE}${DATA_FILE_PATH}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`OneDrive load failed (${res.status})`);
  return res.json();
}

/**
 * Save { students, auditLog } to the user's OneDrive app folder, overwriting
 * the previous version. OneDrive itself keeps automatic version history for
 * every save (visible in OneDrive's web UI → right-click file → Version
 * history), which is the "editing history" / rollback safety net.
 */
export async function saveToOneDrive(msalInstance, account, data) {
  const token = await getGraphToken(msalInstance, account);
  const res = await fetch(`${GRAPH_BASE}${DATA_FILE_PATH}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(data, null, 2),
  });
  if (!res.ok) throw new Error(`OneDrive save failed (${res.status})`);
  return res.json();
}

/**
 * Overwrite the Excel backup file in the user's OneDrive app folder with the
 * given workbook bytes. This is a best-effort human-readable copy of the
 * data — the JSON file remains the source of truth, so callers should treat
 * failures here as non-fatal (e.g. just a quiet toast, not a blocking error).
 */
export async function saveExcelBackupToOneDrive(msalInstance, account, workbookArrayBuffer) {
  const token = await getGraphToken(msalInstance, account);
  const res = await fetch(`${GRAPH_BASE}${EXCEL_FILE_PATH}`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    body: workbookArrayBuffer,
  });
  if (!res.ok) throw new Error(`OneDrive Excel backup failed (${res.status})`);
  return res.json();
}

/**
 * Fetch the signed-in user's basic profile (used to label edit-history
 * entries with a real name instead of an email address).
 */
export async function getMyProfile(msalInstance, account) {
  const token = await getGraphToken(msalInstance, account);
  const res = await fetch(`${GRAPH_BASE}/me`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Profile fetch failed (${res.status})`);
  return res.json();
}
