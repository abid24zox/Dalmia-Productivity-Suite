import { useEffect, useState } from "react";
import { Cloud, Loader2, LogOut } from "lucide-react";
import { MSAL_CONFIGURED, getAccount, signIn, signOut } from "./msal";

// `compact` renders a small header pill (used in the main app bar, post-login,
// optional — never blocks anything); the default renders a fuller panel.
export default function OneDriveConnect({ compact = false }) {
  const [account, setAccount] = useState(() => getAccount());
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => { if (MSAL_CONFIGURED) setAccount(getAccount()); }, []);

  if (!MSAL_CONFIGURED) {
    if (compact) return null; // don't clutter the header when it isn't set up
    return <div className="mt-4 rounded-xl border border-dashed border-slate-300 bg-white p-4 text-xs text-slate-400">OneDrive connect isn't configured yet — set VITE_MSAL_CLIENT_ID in .env.</div>;
  }

  const connect = async () => {
    setErr(""); setBusy(true);
    try { const acc = await signIn(); setAccount(acc); }
    catch (e) {
      if (e.errorCode === "user_cancelled") { /* quiet — they backed out on purpose */ }
      else if (e.errorCode === "popup_window_error" || e.errorCode === "empty_window_error") setErr("Popup blocked — allow popups for this site and try again.");
      else setErr(e.message || "Couldn't connect to OneDrive.");
    }
    setBusy(false);
  };
  const disconnect = async () => {
    setErr(""); setBusy(true);
    try { await signOut(); setAccount(null); }
    catch (e) { setErr(e.message || "Couldn't disconnect."); }
    setBusy(false);
  };

  if (compact) {
    return (
      <div className="relative">
        {account ? (
          <button onClick={disconnect} disabled={busy} title={`Connected to OneDrive as ${account.username} — click to disconnect`} className="inline-flex items-center justify-center gap-1.5 rounded-md border border-emerald-200 bg-emerald-50 px-2.5 py-2 text-sm font-medium text-emerald-700 hover:bg-emerald-100 disabled:opacity-50">{busy ? <Loader2 size={14} className="animate-spin" /> : <Cloud size={14} />}<span className="hidden sm:inline">OneDrive</span></button>
        ) : (
          <button onClick={connect} disabled={busy} title="Connect to OneDrive" className="inline-flex items-center justify-center gap-1.5 rounded-md border border-slate-200 bg-white px-2.5 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50">{busy ? <Loader2 size={14} className="animate-spin" /> : <Cloud size={14} />}<span className="hidden sm:inline">OneDrive</span></button>
        )}
        {err && <div className="absolute right-0 top-full z-10 mt-1 w-56 rounded-md border border-amber-200 bg-amber-50 px-2 py-1.5 text-xs text-amber-700">{err}</div>}
      </div>
    );
  }

  return (
    <div className="mt-4 rounded-xl border border-slate-200 bg-white p-4">
      <div className="mb-1 flex items-center gap-1.5 text-xs font-medium text-slate-500"><Cloud size={14} /> OneDrive</div>
      {account ? (
        <div className="flex items-center justify-between gap-2">
          <div className="min-w-0 flex-1 truncate text-sm text-slate-700">Connected as <span className="font-medium">{account.username}</span></div>
          <button onClick={disconnect} disabled={busy} className="inline-flex shrink-0 items-center gap-1 rounded-md border border-slate-200 px-2.5 py-1.5 text-xs font-medium text-slate-500 hover:bg-slate-50 disabled:opacity-50">{busy ? <Loader2 size={13} className="animate-spin" /> : <LogOut size={13} />} Disconnect</button>
        </div>
      ) : (
        <button onClick={connect} disabled={busy} className="inline-flex w-full items-center justify-center gap-1.5 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50">{busy ? <Loader2 size={14} className="animate-spin" /> : <Cloud size={14} />} Connect to OneDrive</button>
      )}
      {err && <div className="mt-2 text-xs text-amber-700">{err}</div>}
    </div>
  );
}
