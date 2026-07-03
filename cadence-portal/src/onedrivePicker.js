// Picks a file from the signed-in user's OneDrive.
//
// We deliberately do NOT use Microsoft's embedded js.live.net picker: for
// work/school accounts that picker queries the SharePoint REST endpoint
// (https://{tenant}-my.sharepoint.com/...) and demands a *SharePoint-audience*
// access token. We only hold a Microsoft Graph token (audience
// graph.microsoft.com), so SharePoint answered every request with 401 and the
// picker failed before it ever returned a file.
//
// Instead we browse the drive through Microsoft Graph ourselves — the same
// audience our token is already valid for — which works identically for
// personal and business accounts. The UI is a small in-app modal rendered by
// <OneDrivePickerHost/> (see App.jsx); this module is the data + bridge layer.
import { MSAL_CONFIGURED, getAccount, signIn, getGraphToken } from "./msal";

// Bridge: pickOneDriveFile() can be called from any component, but the actual
// browser modal is mounted once at the app root. The host registers its opener
// here; pickOneDriveFile() invokes it and awaits the user's choice.
let opener = null;
export function registerOneDriveOpener(fn) { opener = fn; }

export const b64FromArrayBuffer = (buf) => {
  let binary = "";
  const bytes = new Uint8Array(buf);
  const CHUNK = 0x8000; // chunk to avoid "too many arguments" on big files
  for (let i = 0; i < bytes.length; i += CHUNK) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  return btoa(binary);
};

async function graph(path, token) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json())?.error?.message || ""; } catch { /* ignore */ }
    throw new Error(`Microsoft Graph ${res.status}${detail ? " — " + detail : ""}`);
  }
  return res;
}

// Children of a folder (the drive root when itemId is null), folders first.
export async function odListChildren(token, itemId) {
  const base = itemId ? `/me/drive/items/${itemId}/children` : `/me/drive/root/children`;
  const res = await graph(`${base}?$select=id,name,folder,file,size&$top=200`, token);
  const j = await res.json();
  return (j.value || []).sort((a, b) => (!!b.folder - !!a.folder) || a.name.localeCompare(b.name));
}

// Bytes of a picked file, fetched through Graph /content (personal + business).
export async function odDownload(token, item) {
  const res = await graph(`/me/drive/items/${item.id}/content`, token);
  return res.arrayBuffer();
}

// Ensures a connection + Graph token, then opens the in-app browser modal.
// Resolves { name, dataB64 } for the chosen file, or null if cancelled.
export async function pickOneDriveFile() {
  if (!MSAL_CONFIGURED) throw new Error("OneDrive connect isn't configured.");
  if (!getAccount()) await signIn(); // first-time only — persists for future picks
  const token = await getGraphToken();
  if (!opener) throw new Error("OneDrive picker isn't ready — reload and try again.");
  return opener(token);
}
