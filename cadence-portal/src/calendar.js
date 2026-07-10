// One-way sync of the signed-in user's activities into their Outlook / Microsoft
// 365 calendar via Microsoft Graph. Reuses the same MSAL sign-in as OneDrive; the
// only extra requirement is the Calendars.ReadWrite scope (added to the Azure app
// registration), which the user consents to on first use.
//
// Each upcoming task becomes an ALL-DAY event on its due date. The returned Graph
// event id is stored back on the activity (outlookEventId) so a re-sync UPDATES
// the event instead of duplicating it, and completed/removed tasks delete theirs.
import { MSAL_CONFIGURED, getAccount, signIn, getGraphToken } from "./msal";

const TZ = (typeof Intl !== "undefined" && Intl.DateTimeFormat().resolvedOptions().timeZone) || "UTC";
const CAL_SCOPES = ["Calendars.ReadWrite"];

async function calToken() {
  if (!getAccount()) await signIn(); // first-time sign-in (persists afterwards)
  return getGraphToken(CAL_SCOPES); // triggers the one-time calendar consent popup if needed
}

async function graph(path, method, token, body, extraHeaders) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method, headers: { Authorization: `Bearer ${token}`, "content-type": "application/json", ...(extraHeaders || {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    if (method === "DELETE" && res.status === 404) return null; // already gone in Outlook
    let detail = ""; try { detail = (await res.json())?.error?.message || ""; } catch { /* ignore */ }
    throw new Error(`Microsoft Graph ${res.status}${detail ? " — " + detail : ""}`);
  }
  return method === "DELETE" ? null : res.json();
}

const nextDay = (isoDate) => { const d = new Date(isoDate + "T00:00:00"); d.setDate(d.getDate() + 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; };

// An all-day event on the task's due date, tagged so it's recognisable in Outlook.
function buildEvent(t) {
  return {
    subject: t.title,
    body: { contentType: "text", content: [t.description || null, t.workTitle ? `Work: ${t.workTitle}` : null, `Estimated effort: ${t.planned || 0}h${t.actual ? ` (logged ${t.actual}h so far)` : ""}`, "— added by Cadence"].filter(Boolean).join("\n\n") },
    isAllDay: true,
    start: { dateTime: `${t.date}T00:00:00.000`, timeZone: TZ },
    end: { dateTime: `${nextDay(t.date)}T00:00:00.000`, timeZone: TZ },
    isReminderOn: true,
    reminderMinutesBeforeStart: 1440, // 1 day before
    categories: ["Cadence"],
  };
}

// tasks: [{ id, title, description, date, planned, actual, status, workTitle, outlookEventId }]
// saveEventId(activityId, eventIdOrNull): persist the Graph event id on the activity.
// Returns { created, updated, removed }.
export async function syncTasksToOutlook(tasks, saveEventId) {
  if (!MSAL_CONFIGURED) throw new Error("Microsoft sign-in isn't configured.");
  const token = await calToken();
  const upcoming = tasks.filter((t) => t.date && t.status !== "executed" && t.status !== "cancelled");
  const upcomingIds = new Set(upcoming.map((t) => t.id));
  let created = 0, updated = 0, removed = 0;

  for (const t of upcoming) {
    const ev = buildEvent(t);
    if (t.outlookEventId) {
      try { await graph(`/me/events/${t.outlookEventId}`, "PATCH", token, ev); updated++; }
      catch { const made = await graph("/me/events", "POST", token, ev); await saveEventId(t.id, made.id); created++; } // event was deleted in Outlook — recreate
    } else {
      const made = await graph("/me/events", "POST", token, ev); await saveEventId(t.id, made.id); created++;
    }
  }
  // Tasks that are done/cancelled/removed but still have a calendar event → delete it.
  for (const t of tasks) {
    if (t.outlookEventId && !upcomingIds.has(t.id)) { await graph(`/me/events/${t.outlookEventId}`, "DELETE", token); await saveEventId(t.id, null); removed++; }
  }
  return { created, updated, removed };
}

// Reverse sync: read the signed-in user's meetings for the CURRENT week (Mon–Sun)
// straight from Outlook, minus the events Cadence itself created (tagged
// "Cadence"). The service then AI-files each under the right work. Returns a
// normalized list the /api/calendar/import endpoint understands.
export async function readWeekEvents() {
  if (!MSAL_CONFIGURED) throw new Error("Microsoft sign-in isn't configured.");
  const token = await calToken();
  const now = new Date();
  const monday = new Date(now); monday.setDate(now.getDate() - ((now.getDay() + 6) % 7)); monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday); sunday.setDate(monday.getDate() + 7);
  const q = `/me/calendarView?startDateTime=${monday.toISOString()}&endDateTime=${sunday.toISOString()}`
    + `&$select=id,subject,bodyPreview,start,end,isAllDay,isCancelled,showAs,attendees,categories&$top=100&$orderby=start/dateTime`;
  // Prefer header makes Graph return start/end in the user's local zone, so the
  // date we slice off matches their calendar (not UTC, which can be a day off).
  const data = await graph(q, "GET", token, null, { Prefer: `outlook.timezone="${TZ}"` });
  return (data.value || [])
    .filter((ev) => !(ev.categories || []).includes("Cadence")) // skip our own synced tasks
    .map((ev) => ({
      id: ev.id,
      subject: ev.subject || "(no title)",
      bodyPreview: (ev.bodyPreview || "").slice(0, 300),
      start: ev.start?.dateTime || null,
      end: ev.end?.dateTime || null,
      isAllDay: !!ev.isAllDay,
      isCancelled: !!ev.isCancelled,
      showAs: ev.showAs || null,
      attendees: (ev.attendees || []).map((a) => a.emailAddress?.name || a.emailAddress?.address).filter(Boolean).slice(0, 8),
    }));
}
