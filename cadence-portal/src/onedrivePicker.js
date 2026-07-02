// Embeds Microsoft's official OneDrive File Picker (v8, https://js.live.net)
// so "Attach a document" can pull a file straight from OneDrive instead of
// the local disk. Reuses the same MSAL connection as the header's "Connect to
// OneDrive" button — if the user is already connected, this never re-prompts;
// if not, it connects them inline the first time only (see pickOneDriveFile).
import { MSAL_CONFIGURED, getAccount, signIn, getGraphToken } from "./msal";

const PICKER_SCRIPT_URL = "https://js.live.net/v7.2/OneDrive.js";
let scriptPromise = null;
function loadPickerScript() {
  if (window.OneDrive) return Promise.resolve();
  if (!scriptPromise) {
    scriptPromise = new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = PICKER_SCRIPT_URL;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error("Couldn't load the OneDrive picker."));
      document.head.appendChild(s);
    });
  }
  return scriptPromise;
}

async function graphGet(path, token) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Microsoft Graph ${res.status} on ${path}`);
  return res.json();
}

const b64FromArrayBuffer = (buf) => {
  let binary = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
};

// Resolves with { name, dataB64 } for the picked file, or null if the user
// cancelled the picker. Throws on connection/picker/download failures.
export async function pickOneDriveFile() {
  if (!MSAL_CONFIGURED) throw new Error("OneDrive connect isn't configured.");
  if (!getAccount()) await signIn(); // first-time only — persists for every future pick
  const token = await getGraphToken();
  await loadPickerScript();

  // The v8 picker needs the connected drive's own webUrl as its endpointHint
  // for work/school accounts (personal accounts fall back to api.onedrive.com).
  const drive = await graphGet("/me/drive", token);
  const endpointHint = drive.driveType === "personal" ? "api.onedrive.com" : drive.webUrl;

  const picked = await new Promise((resolve, reject) => {
    window.OneDrive.open({
      clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
      action: "query",
      multiSelect: false,
      advanced: { accessToken: token, endpointHint, queryParameters: "select=id,name,file,@microsoft.graph.downloadUrl" },
      success: (res) => resolve((res.value || [])[0] || null),
      cancel: () => resolve(null),
      error: (e) => reject(new Error(e?.message || "OneDrive picker failed.")),
    });
  });
  if (!picked) return null;

  const downloadUrl = picked["@microsoft.graph.downloadUrl"];
  const fileRes = await fetch(downloadUrl);
  if (!fileRes.ok) throw new Error(`Couldn't download "${picked.name}" from OneDrive.`);
  const buf = await fileRes.arrayBuffer();
  return { name: picked.name, dataB64: b64FromArrayBuffer(buf) };
}
