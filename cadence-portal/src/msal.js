// Microsoft sign-in for "Connect to OneDrive" — a public SPA client (PKCE auth
// code flow), no client secret involved. Configured entirely via VITE_MSAL_*
// env vars; if VITE_MSAL_CLIENT_ID is unset the connect button hides itself
// and none of this ever touches the network (see MSAL_CONFIGURED).
import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";

export const MSAL_CONFIGURED = !!import.meta.env.VITE_MSAL_CLIENT_ID;

// Files.Read (not Files.Read.All) — the broader scope needs tenant admin
// consent in many orgs, and this pass only needs the signed-in user's own
// OneDrive. Request the wider scope later, incrementally, once an actual
// file picker needs shared/SharePoint files.
const SCOPES = ["User.Read", "Files.Read", "offline_access"];

export const msalInstance = MSAL_CONFIGURED
  ? new PublicClientApplication({
      auth: {
        clientId: import.meta.env.VITE_MSAL_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${import.meta.env.VITE_MSAL_TENANT_ID || "common"}`,
        redirectUri: import.meta.env.VITE_MSAL_REDIRECT_URI || window.location.origin,
      },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
    })
  : null;

// msal-browser v3 requires an async initialize() before any other call.
let initPromise = null;
function ready() {
  if (!msalInstance) return Promise.reject(new Error("OneDrive connect isn't configured."));
  if (!initPromise) initPromise = msalInstance.initialize();
  return initPromise;
}

export function getAccount() {
  if (!msalInstance) return null;
  return msalInstance.getAllAccounts()[0] || null;
}

export async function signIn() {
  await ready();
  const res = await msalInstance.loginPopup({ scopes: SCOPES });
  msalInstance.setActiveAccount(res.account);
  return res.account;
}

export async function signOut() {
  await ready();
  const account = getAccount();
  if (account) await msalInstance.logoutPopup({ account });
}

// For a future OneDrive file picker — silent refresh first, falls back to a
// popup only if consent needs to be renewed.
export async function getGraphToken(scopes = SCOPES) {
  await ready();
  const account = getAccount();
  if (!account) throw new Error("Not connected to OneDrive.");
  try {
    const res = await msalInstance.acquireTokenSilent({ scopes, account });
    return res.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const res = await msalInstance.acquireTokenPopup({ scopes, account });
      return res.accessToken;
    }
    throw e;
  }
}
