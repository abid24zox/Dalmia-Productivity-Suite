import { useState, useRef, useEffect, useCallback } from "react";
import {
  Lock, User, LogOut, LayoutGrid, Calendar, Users, ClipboardCheck, Plus,
  Sparkles, Play, Square, Check, X, ChevronLeft, ChevronRight, Target, Clock,
  AlertTriangle, Mic, FileText, Upload, Pencil, Trash2, Loader2, MessageSquare, Star,
  Gauge, Type as TypeIcon, Wand2, CalendarClock, ArrowRight, AlertCircle, ClipboardList, Cloud, Folder, Search, Mail, Send,
  FileSpreadsheet, Presentation, Paperclip, CheckCircle2, Circle, Package,
} from "lucide-react";
import { api } from "./api";
import OneDriveConnect from "./OneDriveConnect";
import { MSAL_CONFIGURED } from "./msal";
import { pickOneDriveFile, registerOneDriveOpener, odListChildren, odDownload, b64FromArrayBuffer } from "./onedrivePicker";
import { syncTasksToOutlook } from "./calendar";

/* =====================================================================
 * Cadence portal — LIVE. State is the shared Cadence service (single source
 * of truth); the Teams agent reads/writes the same store, so changes flow both
 * ways. The in-house AI runs through the service's Azure AI Foundry model (the
 * same model the Teams bot uses) — no model keys ever reach the browser.
 * =================================================================== */
async function aiComplete(system, user) {
  return api.aiComplete(system, user);
}
function parseJSON(t) { const c = t.replace(/```json/gi, "").replace(/```/g, "").trim(); return JSON.parse(c.slice(c.indexOf("{"), c.lastIndexOf("}") + 1)); }
const AI = {
  async decompose(title, type) { return parseJSON(await aiComplete('Enterprise delivery planner. Break the goal into 3-6 works (phases of execution), each with 1-4 concrete activities. Return ONLY JSON: {"works":[{"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Work: "${title}". Type: ${type}.`)); },
  async extractMom(text) { return parseJSON(await aiComplete('Read meeting minutes, extract work. Return ONLY JSON: {"works":[{"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Minutes:\n"""${text}"""`)); },
  async modifyPlan(planText, instruction) { return parseJSON(await aiComplete('Edit a project plan. Return ONLY JSON: {"ops":[{"op":"add_activity","work":string,"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}|{"op":"add_work","title":string}|{"op":"retype","match":string,"type":"self"|"meeting"|"call"|"site"}]}', `Current plan:\n${planText}\n\nInstruction: ${instruction}`)); },
  async insight(title, m) { return parseJSON(await aiComplete('Execution advisor to a CEO. Return ONLY JSON: {"read":string(2 sentences),"action":string}', `Work "${title}". Planning ${m.planning}%, execution ${m.execution}%. Behind: ${m.behind}. Stuck: ${m.stuck || "none"}.`)); },
  async score(work, activity, spec, content) { return parseJSON(await aiComplete('Delivery quality reviewer. Score how well the deliverable satisfies what the activity asked for, 0-100. Return ONLY JSON: {"score":number,"verdict":string(<=6 words),"feedback":string}', `Work: "${work}". Activity: "${activity}". What was asked (spec): "${spec || 'n/a'}". Deliverable submitted:\n"""${content}"""`)); },
  // Draft a short internal email nudging a colleague about a work item.
  async draftEmail(info) {
    const sys = 'Draft a short, warm, professional internal email from a leader to a colleague about ONE specific work item. 3-5 sentences: a friendly opener, the specific ask (unblock / status update / take ownership), and a courteous close. Sign off with the sender\'s first name only. No placeholders. Return ONLY JSON: {"subject":string,"body":string}.';
    const usr = `From: ${info.from}. To: ${info.to}. Work item: "${info.item}"${info.parent ? ` (under ${info.parent})` : ""}. Situation: ${info.context}. Purpose: ${info.purpose}. Today: ${iso(TODAY)}.`;
    return parseJSON(await aiComplete(sys, usr));
  },
  // Decide what kind of item the note describes and extract its fields + best parent.
  // Propose the concrete outputs (documents, spreadsheets, emails, decks) a work
  // should produce, given its title and the activities inside it.
  async suggestDeliverables(workTitle, activityTitles) {
    const sys = 'You define the concrete DELIVERABLES a work package must produce — the tangible outputs, not the tasks. Propose 2-5 items. Each has a short label and a kind. Return ONLY JSON: {"deliverables":[{"label":string,"kind":"document"|"spreadsheet"|"email"|"slides"|"other"}]}';
    const usr = `Work: "${workTitle}". Activities inside it: ${JSON.stringify(activityTitles || [])}.`;
    return parseJSON(await aiComplete(sys, usr));
  },
  async classify(note, ctx) {
    const sys = 'You are a Cadence planning assistant. Decide which single item the user wants to CREATE and extract its fields. Levels (pick the SMALLEST that fits): "objective" = a top-level enterprise outcome (no parent); "initiative" = a program under an objective with a measurable target; "work" = a phase / work-package under an initiative that groups tasks; "activity" = one schedulable task (owner + date + hours) under a work. Set parentTitle to the closest EXISTING item name from the context that this should live under; if none fits leave it "". Return ONLY JSON: {"level":"objective"|"initiative"|"work"|"activity","title":string,"description":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","parentTitle":string,"deadline":string|null,"assignee":string,"team":string,"estimateHrs":number|null,"actType":"self"|"meeting"|"call"|"site"|null,"metric":string,"target":number|null,"unit":string}';
    const usr = `Existing objectives: ${ctx.objectives}\nExisting initiatives: ${ctx.initiatives}\nExisting works: ${ctx.works}\nPeople: ${ctx.people}\nTeams: ${ctx.teams}\nToday: ${ctx.today}\n\nUser note:\n"""${note}"""`;
    return parseJSON(await aiComplete(sys, usr));
  },
};

/* ---------- dates ---------- */
const MSD = 86400000;
const sod = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const TODAY = sod(new Date());
const addDays = (d, n) => sod(new Date(sod(d).getTime() + n * MSD));
// Local (not UTC) YYYY-MM-DD — toISOString() shifts the date back a day in
// timezones ahead of UTC, which is why the calendar looked one day off.
const iso = (d) => { const x = sod(d); return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`; };
const parseISO = (s) => (s ? sod(new Date(s + "T00:00:00")) : null);
const startOfWeek = (d) => addDays(d, -((sod(d).getDay() + 6) % 7));
const fmtFull = (d) => d.toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
const fmtShort = (isoStr) => (isoStr ? parseISO(isoStr).toLocaleDateString("en-GB", { day: "numeric", month: "short" }) : "—");
const isOverdue = (a) => a.date && parseISO(a.date) < TODAY && a.status !== "executed";
// Days until a date (negative = past). Used to highlight due / overdue nodes.
const daysLeft = (dateStr) => (dateStr ? Math.round((parseISO(dateStr) - TODAY) / MSD) : null);
// Compact deadline chip for a work/initiative: "in 3d" / "due today" / "5d over".
function DueChip({ date, small }) {
  if (!date) return null;
  const d = daysLeft(date);
  const tone = d < 0 ? "bg-rose-50 text-rose-700" : d <= 3 ? "bg-amber-50 text-amber-700" : "bg-slate-100 text-slate-500";
  const txt = d < 0 ? `${-d}d over` : d === 0 ? "due today" : `in ${d}d`;
  return <span className={`inline-flex items-center gap-0.5 rounded-full px-1.5 ${small ? "text-[10px]" : "py-0.5 text-xs"} font-medium ${tone}`}><Clock size={small ? 9 : 10} /> {txt}</span>;
}
// Scopes an activity set to a due-date cutoff (e.g. for "as of this date" analysis) — no-op when no date is given.
const actsUpTo = (acts, dueISO) => (dueISO ? acts.filter((a) => a.date && parseISO(a.date) <= parseISO(dueISO)) : acts);
// Start → end timeline chip (activities have both; nodes span their descendants).
function TimelineChip({ start, end, small }) {
  if (!start && !end) return null;
  return <span className={`inline-flex items-center gap-1 text-slate-400 ${small ? "text-[10px]" : "text-xs"}`}><CalendarClock size={small ? 9 : 11} className="shrink-0" />{start ? fmtShort(start) : "?"} <ArrowRight size={small ? 8 : 10} className="shrink-0 text-slate-300" /> {end ? fmtShort(end) : "?"}</span>;
}
// actual / estimated effort in hours; flags an over-run.
function EffortBadge({ planned, actual, small }) {
  if (!planned && !actual) return null;
  const over = planned > 0 && actual > planned;
  return <span className={`inline-flex items-center gap-1 rounded-full bg-slate-100 px-1.5 font-medium text-slate-500 ${small ? "text-[10px]" : "py-0.5 text-xs"}`} title="actual / estimated effort (hours)"><Clock size={small ? 9 : 10} className="shrink-0" />{actual || 0}/{planned || 0}h{over ? <span className="text-rose-600">▲</span> : null}</span>;
}
// Effort roll-up (sum of planned/actual hours over a node's whole subtree).
function effortRollup(works, acts, id) {
  const ids = subtreeIds(works, id);
  const a = acts.filter((x) => ids.includes(x.workId) && x.status !== "cancelled");
  return { planned: a.reduce((s, x) => s + (x.plannedHrs || 0), 0), actual: a.reduce((s, x) => s + (x.actualHrs || 0), 0) };
}
// A node's timeline: earliest descendant start → its deadline (or latest due).
function nodeSpan(works, acts, node) {
  const ids = subtreeIds(works, node.id);
  const a = acts.filter((x) => ids.includes(x.workId) && x.status !== "cancelled");
  const start = a.map((x) => x.startDate).filter(Boolean).sort()[0] || null;
  const end = node.deadline || a.map((x) => x.date).filter(Boolean).sort().slice(-1)[0] || null;
  return { start, end };
}

/* ---------- data ---------- */
// Client id generator — timestamped so ids the portal creates never collide
// with ids the service creates (w5001…) as both write the shared store.
let _uid = 0; const nid = (p) => `${p}${Date.now().toString(36)}${(_uid++).toString(36)}`;
// Live user directory — populated from the service snapshot on login. The
// helpers below read it; it is set before any user data renders.
let USERS = [];
const setDirectory = (users) => { USERS = users || []; };
const fnOf = (id) => USERS.find((u) => u.id === id)?.fn;
const uName = (id) => USERS.find((u) => u.id === id)?.name || "Unassigned";
const uFirst = (id) => uName(id).split(" ")[0];
const initials = (id) => uName(id).split(" ").map((n) => n[0]).join("");
const METRIC_BY_TYPE = { procurement: { metric: "Units issued", unit: "count" }, cost: { metric: "Cost saved", unit: "₹/unit" }, onboarding: { metric: "Cycle time", unit: "days" }, compliance: { metric: "Coverage", unit: "%" }, general: { metric: "Completion", unit: "%" } };
const ACT_TYPES = ["self", "meeting", "call", "site"];

const SEED_WORKS = [
  // Objectives (level 1) — MD-owned
  { id: "o1", parentId: null, level: "objective", title: "Modernise field operations", type: "general", ownerId: "u_vik" },
  { id: "o2", parentId: null, level: "objective", title: "Cut supply-chain cost", type: "general", ownerId: "u_vik" },
  // Initiatives (level 2) — VP-owned, carry the result metric
  { id: "w1", parentId: "o1", level: "initiative", title: "Replace retired laptops (Oct-24 & older)", type: "procurement", ownerId: "u_pri", teamId: "t_it", scope: "group", result: { metric: "Laptops issued", unit: "count", baseline: 0, target: 120, current: 38 } },
  { id: "w1x", parentId: "o1", level: "initiative", title: "Roll out MFA to field staff", type: "compliance", ownerId: "u_pri", teamId: "t_it", scope: "group", result: { metric: "Coverage", unit: "%", baseline: 0, target: 100, current: 20 } },
  { id: "w2", parentId: "o2", level: "initiative", title: "Reduce logistics cost per bag by ₹0.50", type: "cost", ownerId: "u_mee", teamId: "t_sc", scope: "group", result: { metric: "Cost saved", unit: "₹/bag", baseline: 0, target: 0.5, current: 0.18 } },
  // Works (level 3)
  { id: "w1W1", parentId: "w1", level: "work", title: "Demand & selection", type: "procurement", ownerId: "u_roh" },
  { id: "w1W2", parentId: "w1", level: "work", title: "Procurement", type: "procurement", ownerId: "u_pri" },
  { id: "w1W3", parentId: "w1", level: "work", title: "Rollout", type: "procurement", ownerId: "u_roh" },
  { id: "w1xW1", parentId: "w1x", level: "work", title: "Assess & pilot", type: "compliance", ownerId: "u_roh" },
  { id: "w2W1", parentId: "w2", level: "work", title: "Freight optimisation", type: "cost", ownerId: "u_neh" },
  { id: "w2W2", parentId: "w2", level: "work", title: "Commercials", type: "cost", ownerId: "u_mee" },
  // Sub-works (level 4) — hold the activities
  { id: "w1a", parentId: "w1W1", level: "subwork", title: "Employee survey", type: "procurement", ownerId: "u_roh" },
  { id: "w1b", parentId: "w1W1", level: "subwork", title: "Shortlist & finalize models", type: "procurement", ownerId: "u_roh" },
  { id: "w1c", parentId: "w1W2", level: "subwork", title: "Budget approvals from HODs", type: "procurement", ownerId: "u_pri" },
  { id: "w1d", parentId: "w1W2", level: "subwork", title: "Raise PR–PO & procure", type: "procurement", ownerId: "u_pri" },
  { id: "w1e", parentId: "w1W3", level: "subwork", title: "Image, issue & hand over", type: "procurement", ownerId: "u_roh" },
  { id: "w1xa", parentId: "w1xW1", level: "subwork", title: "Assess current auth", type: "compliance", ownerId: "u_roh" },
  { id: "w2a", parentId: "w2W1", level: "subwork", title: "Optimize primary freight routes", type: "cost", ownerId: "u_neh" },
  { id: "w2b", parentId: "w2W1", level: "subwork", title: "Improve truck load factor", type: "cost", ownerId: "u_neh" },
  { id: "w2c", parentId: "w2W2", level: "subwork", title: "Renegotiate transporter contracts", type: "cost", ownerId: "u_mee" },
  // Operations objective — recurring plant routines (cement)
  { id: "o3", parentId: null, level: "objective", title: "Run reliable, compliant plant operations", type: "general", ownerId: "u_vik" },
  { id: "w3", parentId: "o3", level: "initiative", title: "Plant routines & statutory upkeep", type: "compliance", ownerId: "u_mee", teamId: "t_sc", scope: "group", result: { metric: "Plant uptime", unit: "%", baseline: 0, target: 95, current: 88 } },
  { id: "w3W1", parentId: "w3", level: "work", title: "Monthly payroll run", type: "general", ownerId: "u_ana", recurring: { cadence: "monthly" } },
  { id: "w3W2", parentId: "w3", level: "work", title: "Kiln & raw-mill preventive maintenance", type: "general", ownerId: "u_sam", recurring: { cadence: "monthly" } },
  { id: "w3W3", parentId: "w3", level: "work", title: "Statutory & pollution-control compliance", type: "compliance", ownerId: "u_div", recurring: { cadence: "quarterly" } },
  { id: "w3a", parentId: "w3W1", level: "subwork", title: "July payroll cycle", type: "general", ownerId: "u_ana" },
  { id: "w3b", parentId: "w3W2", level: "subwork", title: "July PM cycle — Kiln line 2", type: "general", ownerId: "u_sam" },
  { id: "w3c", parentId: "w3W3", level: "subwork", title: "Q2 emissions returns & consent renewal", type: "compliance", ownerId: "u_div" },
];
const D = (n) => iso(addDays(TODAY, n));
const SEED_ACTS = [
  { id: "a1", workId: "w1a", title: "Draft survey form", assigneeId: "u_roh", date: D(-2), status: "executed", plannedHrs: 2, actualHrs: 2, actType: "self" },
  { id: "a2", workId: "w1a", title: "Send survey to all staff", assigneeId: "u_roh", date: D(-1), status: "executed", plannedHrs: 1, actualHrs: 1, actType: "self" },
  { id: "a3", workId: "w1a", title: "Collate responses", assigneeId: "u_roh", date: D(-1), status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
  { id: "a3b", workId: "w1a", title: "Prepare survey summary", assigneeId: "u_roh", date: D(0), status: "planned", plannedHrs: 1, actualHrs: null, actType: "self" },
  { id: "a4", workId: "w1b", title: "Gather target specs", assigneeId: "u_roh", date: D(-1), status: "executed", plannedHrs: 2, actualHrs: 3, actType: "self" },
  { id: "a5", workId: "w1b", title: "Get 3 vendor quotes", assigneeId: "u_roh", date: D(2), status: "planned", plannedHrs: 3, actualHrs: null, actType: "call" },
  { id: "a6", workId: "w1c", title: "Prepare BOM + cost sheet", assigneeId: null, date: null, status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
  { id: "a7", workId: "w1c", title: "HOD sign-offs (IT, Fin, Ops)", assigneeId: null, date: null, status: "planned", plannedHrs: 2, actualHrs: null, actType: "meeting" },
  { id: "a8", workId: "w1d", title: "Raise purchase requisition", assigneeId: "u_roh", date: D(1), status: "planned", plannedHrs: 1, actualHrs: null, actType: "self" },
  { id: "a9", workId: "w1d", title: "Issue PO to vendor", assigneeId: "u_roh", date: D(3), status: "planned", plannedHrs: 1, actualHrs: null, actType: "self", blocked: true },
  { id: "ae1", workId: "w1e", title: "Image devices", assigneeId: "u_roh", date: D(4), status: "planned", plannedHrs: 3, actualHrs: null, actType: "self" },
  { id: "ae2", workId: "w1e", title: "Hand over to users", assigneeId: "u_roh", date: D(6), status: "planned", plannedHrs: 2, actualHrs: null, actType: "site" },
  { id: "ax1", workId: "w1xa", title: "Inventory apps & auth methods", assigneeId: "u_roh", date: D(0), status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
  { id: "ax2", workId: "w1xa", title: "Map risky logins", assigneeId: "u_roh", date: D(2), status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
  { id: "a10", workId: "w2a", title: "Map current lane costs", assigneeId: "u_neh", date: D(-3), status: "executed", plannedHrs: 3, actualHrs: 3, actType: "self" },
  { id: "a11", workId: "w2a", title: "Model route consolidation", assigneeId: "u_neh", date: D(0), status: "planned", plannedHrs: 4, actualHrs: null, actType: "self" },
  { id: "a12", workId: "w2b", title: "Pull load-factor data by plant", assigneeId: "u_neh", date: D(-1), status: "executed", plannedHrs: 2, actualHrs: 2, actType: "self" },
  { id: "ap1", workId: "w3a", title: "Validate attendance & inputs", assigneeId: "u_ana", date: D(-2), status: "executed", plannedHrs: 2, actualHrs: 2, actType: "self" },
  { id: "ap2", workId: "w3a", title: "Process & disburse salaries", assigneeId: "u_ana", date: D(1), status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
  { id: "am1", workId: "w3b", title: "Inspect refractory & mill bearings", assigneeId: "u_sam", date: D(0), status: "planned", plannedHrs: 4, actualHrs: null, actType: "site" },
  { id: "am2", workId: "w3b", title: "Lubrication & alignment", assigneeId: "u_sam", date: D(3), status: "planned", plannedHrs: 3, actualHrs: null, actType: "site" },
  { id: "ac1", workId: "w3c", title: "Compile CPCB emission data", assigneeId: "u_div", date: D(-2), status: "planned", plannedHrs: 3, actualHrs: null, actType: "self" },
  { id: "ac2", workId: "w3c", title: "File returns & consent renewal", assigneeId: "u_div", date: D(4), status: "planned", plannedHrs: 2, actualHrs: null, actType: "self" },
];
const SEED_CRS = [{ id: "cr1", workId: "w1", subworkId: "w1a", proposerId: "u_roh", kind: "add_activity", desc: "Found 3 depots missed in the first survey — need a quick re-run.", payload: { title: "Re-run survey for 3 missed depots", hrs: 2, type: "self" }, status: "pending" }];
const SEED_TEAMS = [
  { id: "t_it", name: "IT Ops", memberIds: ["u_pri", "u_roh", "u_arj", "u_kav"] },
  { id: "t_sc", name: "Supply Chain", memberIds: ["u_mee", "u_neh", "u_sam", "u_div"] },
  { id: "t_fin", name: "Finance & Controls", memberIds: ["u_ana", "u_raj"] },
  { id: "t_dig", name: "Digital Rollout (cross-fn)", memberIds: ["u_roh", "u_sam", "u_ana"] },
];

/* ---------- meters + stats + rag ---------- */
function subtreeIds(works, id) { const out = [id]; works.filter((w) => w.parentId === id).forEach((k) => out.push(...subtreeIds(works, k.id))); return out; }
function computeMeters(works, acts, topId) {
  const ids = subtreeIds(works, topId);
  const nodes = works.filter((w) => ids.includes(w.id));
  const leaves = nodes.filter((n) => !works.some((w) => w.parentId === n.id));
  const my = acts.filter((a) => ids.includes(a.workId) && a.status !== "cancelled");
  const top = works.find((w) => w.id === topId);
  const definition = ([top.title, top.ownerId, top.type, top.result].filter(Boolean).length / 4) * 100;
  const decomposition = leaves.length ? (leaves.filter((l) => my.some((a) => a.workId === l.id)).length / leaves.length) * 100 : 0;
  const scheduling = my.length ? (my.filter((a) => a.assigneeId && a.date).length / my.length) * 100 : 0;
  const executed = my.filter((a) => a.status === "executed");
  const plannedTot = my.reduce((s, a) => s + a.plannedHrs, 0);
  const completion = plannedTot ? (executed.reduce((s, a) => s + a.plannedHrs, 0) / plannedTot) * 100 : 0;
  const planning = Math.round((definition + decomposition + scheduling) / 3);
  const execution = Math.round(completion);
  const behind = my.some((a) => isOverdue(a));
  const childWorks = works.filter((w) => w.parentId === topId);
  let stuck = null, low = 101;
  childWorks.forEach((c) => { const m = computeMeters(works, acts, c.id); if (m.completion < 100 && m.completion < low) { low = m.completion; stuck = c.title; } });
  let resultPct = null; if (top.result && top.result.target !== top.result.baseline) resultPct = Math.max(0, Math.min(100, ((top.result.current - top.result.baseline) / (top.result.target - top.result.baseline)) * 100));
  return { definition, decomposition, scheduling, completion, planning, execution, behind, stuck, resultPct };
}
function workStats(works, acts, topId) {
  const ids = subtreeIds(works, topId);
  const a = acts.filter((x) => ids.includes(x.workId) && x.status !== "cancelled");
  const done = a.filter((x) => x.status === "executed").length;
  const scheduled = a.filter((x) => x.assigneeId && x.date).length;
  const overdue = a.filter((x) => isOverdue(x)).length;
  const unscheduled = a.length - scheduled;
  const team = [...new Set(a.map((x) => x.assigneeId).filter(Boolean))];
  const deliv = works.filter((w) => ids.includes(w.id)).reduce((s, w) => s + (w.deliverables || []).filter((d) => d.done).length, 0);
  const nextDue = a.filter((x) => x.date && x.status !== "executed").map((x) => x.date).sort()[0] || null;
  return { total: a.length, done, scheduled, unscheduled, overdue, team, deliv, nextDue };
}
function ragOf(m, st) { if (st.overdue > 0) return "red"; if (m.stuck || m.planning - m.execution > 40) return "amber"; return "green"; }
function subtreeAssignees(works, acts, topId) { const ids = subtreeIds(works, topId); return [...new Set(acts.filter((a) => ids.includes(a.workId) && a.assigneeId).map((a) => a.assigneeId))]; }
function visibleTops(works, acts, user) {
  const tops = works.filter((w) => !w.parentId);
  if (user.level === "md") return tops;
  if (user.level === "vp") return tops.filter((w) => w.ownerId === user.id || fnOf(w.ownerId) === user.fn || subtreeAssignees(works, acts, w.id).some((uid) => fnOf(uid) === user.fn));
  return tops.filter((w) => w.ownerId === user.id || subtreeIds(works, w.id).some((id) => acts.some((a) => a.workId === id && a.assigneeId === user.id)));
}
const loadOf = (acts, uid) => acts.filter((a) => a.assigneeId === uid && a.status !== "executed" && a.status !== "cancelled").reduce((s, a) => s + a.plannedHrs, 0);
const RAG = { green: ["bg-emerald-500", "text-emerald-700", "bg-emerald-50", "On track"], amber: ["bg-amber-500", "text-amber-700", "bg-amber-50", "At risk"], red: ["bg-rose-500", "text-rose-700", "bg-rose-50", "Behind"] };

/* ---------- tree / roll-up helpers ---------- */
const LEVEL_LABEL = { objective: "Objective", initiative: "Initiative", work: "Work", activity: "Activity" };
const CHILD_LABEL = { objective: "initiatives", initiative: "works", work: "activities" };
const CHILD_LEVEL = { objective: "initiative", initiative: "work", work: "activity" };
// One distinct colour per level so you always know your altitude. Kept clear of the RAG greens/ambers/reds.
const LEVEL_THEME = {
  objective: { name: "Objective", bar: "bg-blue-800", chip: "bg-blue-100 text-blue-800", dot: "bg-blue-800", ring: "border-blue-200", soft: "bg-blue-50", text: "text-blue-800" },
  initiative: { name: "Initiative", bar: "bg-blue-500", chip: "bg-blue-100 text-blue-700", dot: "bg-blue-500", ring: "border-blue-200", soft: "bg-blue-50", text: "text-blue-700" },
  work: { name: "Work", bar: "bg-teal-500", chip: "bg-teal-100 text-teal-700", dot: "bg-teal-500", ring: "border-teal-200", soft: "bg-teal-50", text: "text-teal-700" },
  activity: { name: "Activity", bar: "bg-slate-400", chip: "bg-slate-100 text-slate-600", dot: "bg-slate-400", ring: "border-slate-200", soft: "bg-slate-50", text: "text-slate-600" },
};
const HOME_LEVEL = { md: "objective", vp: "initiative", member: "work" };
function subtreeActs(works, acts, id) { const ids = subtreeIds(works, id); return acts.filter((a) => ids.includes(a.workId) && a.status !== "cancelled"); }
// Strict roll-up: any overdue/blocked descendant turns a node red, all the way up.
function nodeRag(works, acts, id) { const sa = subtreeActs(works, acts, id); if (sa.some((a) => isOverdue(a) || a.blocked)) return "red"; const m = computeMeters(works, acts, id); if (m.stuck || m.planning - m.execution > 40) return "amber"; return "green"; }
// The specific leaf causing the red — blockers first, then overdue.
function deepestIssue(works, acts, id) { const sa = subtreeActs(works, acts, id); return sa.find((a) => a.blocked) || sa.find((a) => isOverdue(a)) || null; }
function attentionCount(works, acts, id) { const sa = subtreeActs(works, acts, id); return { blocked: sa.filter((a) => a.blocked).length, overdue: sa.filter((a) => isOverdue(a)).length }; }
function homeNodes(works, acts, user) {
  const lvl = HOME_LEVEL[user.level]; const nodes = works.filter((w) => w.level === lvl);
  if (user.level === "md") return nodes;
  if (user.level === "vp") return nodes.filter((w) => fnOf(w.ownerId) === user.fn || w.ownerId === user.id || subtreeAssignees(works, acts, w.id).some((uid) => fnOf(uid) === user.fn));
  return nodes.filter((w) => w.ownerId === user.id || subtreeIds(works, w.id).some((id) => acts.some((a) => a.workId === id && a.assigneeId === user.id)));
}
function breadcrumbPath(works, user, nodeId) {
  const ceil = HOME_LEVEL[user.level]; const path = []; let cur = works.find((w) => w.id === nodeId); let g = 0;
  while (cur && g++ < 12) { path.unshift(cur); if (cur.level === ceil) break; const p = works.find((w) => w.id === cur.parentId); if (!p) break; cur = p; }
  return path;
}
function scopeInitiatives(works, acts, user) { const homes = homeNodes(works, acts, user); const ids = new Set(homes.flatMap((h) => subtreeIds(works, h.id))); return works.filter((w) => w.level === "initiative" && ids.has(w.id)); }
// Objectives a user can see — MD sees all; everyone else sees only objectives that parent one of their scoped initiatives.
function scopedObjectives(works, acts, user) { const inits = scopeInitiatives(works, acts, user); return works.filter((w) => w.level === "objective" && (user.level === "md" || inits.some((i) => i.parentId === w.id))); }

/* ---------- ui atoms ---------- */
const inputCls = "w-full rounded-md border border-slate-200 px-3 py-2 text-sm outline-none focus:border-slate-400";
const btnDark = "inline-flex items-center justify-center gap-1.5 rounded-md bg-slate-900 px-3 py-2 text-sm font-medium text-white hover:bg-slate-700 disabled:opacity-50";
const btnLight = "inline-flex items-center justify-center gap-1.5 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50";
const btnViolet = "inline-flex items-center justify-center gap-1.5 rounded-md bg-blue-700 px-3 py-2 text-sm font-medium text-white hover:bg-blue-800 disabled:opacity-50";
function Avatar({ id, size = 22 }) { return <span className="inline-flex items-center justify-center rounded-full bg-slate-200 font-medium text-slate-600" style={{ width: size, height: size, fontSize: size * 0.42 }}>{initials(id)}</span>; }
function Chip({ children, tone = "slate" }) { const t = { slate: "bg-slate-100 text-slate-600", rose: "bg-rose-50 text-rose-700", amber: "bg-amber-50 text-amber-800", emerald: "bg-emerald-50 text-emerald-700", violet: "bg-blue-100 text-blue-800", blue: "bg-blue-100 text-blue-700" }[tone]; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${t}`}>{children}</span>; }
function LevelChip({ level }) { const th = LEVEL_THEME[level] || LEVEL_THEME.activity; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${th.chip}`}><span className={`h-1.5 w-1.5 rounded-full ${th.bar}`} /> {th.name}</span>; }
function StatusPill({ rag }) { const rg = RAG[rag]; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${rg[2]} ${rg[1]}`}><span className={`h-1.5 w-1.5 rounded-full ${rg[0]}`} /> {rg[3]}</span>; }
function LevelLegend() { return <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-slate-400"><span className="font-medium text-slate-500">Levels:</span>{["objective", "initiative", "work", "activity"].map((l) => { const th = LEVEL_THEME[l]; return <span key={l} className="inline-flex items-center gap-1"><span className={`h-2 w-2 rounded-full ${th.bar}`} />{th.name}</span>; })}<span className="text-slate-300">→ each level nests inside the one before it</span></div>; }
function ProgressPair({ planning, execution, size = "sm" }) {
  const big = size === "lg";
  return (
    <div className="flex gap-4">
      {[["Planned", planning, "bg-teal-500", "text-teal-700", "bg-teal-100"], ["Done", execution, "bg-amber-500", "text-amber-700", "bg-amber-100"]].map(([l, pct, f, tx, tr]) => (
        <div key={l} className="flex-1">
          <div className="mb-1 flex items-baseline justify-between"><span className={`${big ? "text-sm" : "text-xs"} text-slate-500`}>{l}</span><span className={`font-mono ${big ? "text-base font-medium" : "text-xs"} ${tx}`}>{Math.round(pct)}%</span></div>
          <div className={`${big ? "h-2.5" : "h-2"} rounded-full ${tr} overflow-hidden`}><div className={`h-full rounded-full ${f}`} style={{ width: `${pct}%` }} /></div>
        </div>
      ))}
    </div>
  );
}
function Modal({ children, onClose, wide }) {
  return <div className="fixed inset-0 z-30 flex items-center justify-center bg-slate-900 bg-opacity-40 p-4" onClick={onClose}><div className={`w-full ${wide ? "max-w-2xl" : "max-w-md"} rounded-xl border border-slate-200 bg-white p-5 shadow-lg`} style={{ maxHeight: "88vh", overflowY: "auto" }} onClick={(e) => e.stopPropagation()}>{children}</div></div>;
}
const ACT_ICON = { meeting: Users, call: MessageSquare, self: User, site: Target };

/* ---------- OneDrive file browser (Microsoft Graph-backed) ---------- */
const fmtSize = (n) => !n ? "" : n < 1024 ? `${n} B` : n < 1048576 ? `${Math.round(n / 1024)} KB` : `${(n / 1048576).toFixed(1)} MB`;
// Mounted once at the app root. pickOneDriveFile() (called from anywhere)
// resolves through this single host, so both call sites keep the simple
// `const picked = await pickOneDriveFile()` contract.
function OneDrivePickerHost() {
  const [token, setToken] = useState(null); // non-null while the modal is open
  const resolveRef = useRef(null);
  useEffect(() => {
    registerOneDriveOpener((tok) => new Promise((resolve) => { resolveRef.current = resolve; setToken(tok); }));
    return () => registerOneDriveOpener(null);
  }, []);
  const finish = (val) => { const r = resolveRef.current; resolveRef.current = null; setToken(null); if (r) r(val); };
  if (!token) return null;
  return <OneDriveBrowser token={token} onCancel={() => finish(null)} onPick={finish} />;
}
function OneDriveBrowser({ token, onCancel, onPick }) {
  const [path, setPath] = useState([{ id: null, name: "OneDrive" }]);
  const [items, setItems] = useState(null); // null = loading
  const [err, setErr] = useState("");
  const [downloading, setDownloading] = useState(null);
  const here = path[path.length - 1];
  useEffect(() => {
    let alive = true; setItems(null); setErr("");
    odListChildren(token, here.id)
      .then((v) => { if (alive) setItems(v); })
      .catch((e) => { if (alive) { setErr(e.message || "Couldn't list OneDrive."); setItems([]); } });
    return () => { alive = false; };
  }, [here.id, token]);
  const pick = async (it) => {
    setDownloading(it.id); setErr("");
    try { const buf = await odDownload(token, it); onPick({ name: it.name, dataB64: b64FromArrayBuffer(buf) }); }
    catch (e) { setErr(e.message || "Couldn't download that file."); setDownloading(null); }
  };
  return (
    <Modal onClose={onCancel} wide>
      <div className="mb-3 flex items-center justify-between">
        <h3 className="inline-flex items-center gap-2 text-sm font-medium text-slate-900"><Cloud size={16} className="text-sky-600" /> Choose a file from OneDrive</h3>
        <button onClick={onCancel} className="text-slate-400 hover:text-slate-600"><X size={18} /></button>
      </div>
      <div className="mb-2 flex flex-wrap items-center gap-0.5 text-xs">
        {path.map((p, i) => (
          <span key={i} className="inline-flex items-center gap-0.5">
            {i > 0 && <ChevronRight size={12} className="text-slate-300" />}
            <button onClick={() => setPath((s) => s.slice(0, i + 1))} className={`rounded px-1 py-0.5 hover:bg-slate-100 ${i === path.length - 1 ? "font-medium text-slate-700" : "text-sky-600"}`}>{p.name}</button>
          </span>
        ))}
      </div>
      {err && <div className="mb-2 rounded-md border border-amber-200 bg-amber-50 px-2.5 py-1.5 text-xs text-amber-700">{err}</div>}
      <div className="max-h-80 divide-y divide-slate-100 overflow-y-auto rounded-lg border border-slate-200">
        {items === null ? (
          <div className="flex items-center justify-center gap-2 py-10 text-sm text-slate-400"><Loader2 size={16} className="animate-spin" /> Loading…</div>
        ) : items.length === 0 ? (
          <div className="py-10 text-center text-sm text-slate-400">This folder is empty.</div>
        ) : items.map((it) => it.folder ? (
          <button key={it.id} onClick={() => setPath((s) => [...s, { id: it.id, name: it.name }])} className="flex w-full items-center gap-2.5 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50">
            <Folder size={16} className="shrink-0 text-amber-500" /><span className="min-w-0 flex-1 truncate">{it.name}</span>
            <span className="text-xs text-slate-400">{it.folder.childCount || ""}</span><ChevronRight size={14} className="shrink-0 text-slate-300" />
          </button>
        ) : (
          <button key={it.id} onClick={() => pick(it)} disabled={!!downloading} className="flex w-full items-center gap-2.5 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50 disabled:opacity-50">
            <FileText size={16} className="shrink-0 text-slate-400" /><span className="min-w-0 flex-1 truncate">{it.name}</span>
            {downloading === it.id ? <Loader2 size={14} className="shrink-0 animate-spin text-slate-400" /> : <span className="text-xs text-slate-400">{fmtSize(it.size)}</span>}
          </button>
        ))}
      </div>
      <p className="mt-2 text-xs text-slate-400">PDF, Word, .md and .txt are read automatically once picked.</p>
    </Modal>
  );
}

/* ---------- multi-modal input ---------- */
// Runtime capability flags, set once from the service /api/health.
const CAP = { deepgram: false };
const fileToB64 = (f) => new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(String(r.result || "").split(",")[1] || ""); r.readAsDataURL(f); });

function MultiModalInput({ value, onChange, placeholder, onPdf }) {
  const [mode, setMode] = useState("type");
  const [rec, setRec] = useState(false);
  const [docName, setDocName] = useState(null);
  const [working, setWorking] = useState(null); // 'transcribing' | 'extracting'
  const [err, setErr] = useState("");
  const rr = useRef(null); const mr = useRef(null); const chunks = useRef([]);
  const append = (t) => onChange((value ? value + " " : "") + t);

  // Browser Web Speech API (used when Deepgram isn't configured, or as fallback).
  const startBrowserSR = () => {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SR) { setErr("This browser has no speech recognition — type instead."); return false; }
    const r = new SR(); r.continuous = true; r.interimResults = false;
    r.onresult = (e) => { let t = ""; for (let i = e.resultIndex; i < e.results.length; i++) t += e.results[i][0].transcript; append(t); };
    r.onerror = () => setRec(false); r.onend = () => setRec(false); rr.current = r; r.start(); setRec(true); return true;
  };
  // Record audio and transcribe via Deepgram (server-side), falling back to the
  // browser recognizer if the mic/recorder is unavailable.
  const startDeepgram = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const rec = new MediaRecorder(stream); chunks.current = [];
      rec.ondataavailable = (e) => e.data.size && chunks.current.push(e.data);
      rec.onstop = async () => {
        stream.getTracks().forEach((t) => t.stop());
        const blob = new Blob(chunks.current, { type: rec.mimeType || "audio/webm" });
        if (!blob.size) return;
        setWorking("transcribing");
        try { const b64 = await fileToB64(new File([blob], "audio")); const text = await api.aiTranscribe(b64, blob.type); if (text) append(text); else setErr("Didn't catch that — try again."); }
        catch (e) { setErr(e.message || "Transcription failed."); }
        setWorking(null);
      };
      mr.current = rec; rec.start(); setRec(true);
    } catch { startBrowserSR(); }
  };
  const start = () => { setErr(""); if (CAP.deepgram) startDeepgram(); else startBrowserSR(); };
  const stop = () => { if (mr.current && mr.current.state !== "inactive") mr.current.stop(); if (rr.current) rr.current.stop(); mr.current = null; setRec(false); };

  const readF = async (f) => {
    if (!f) return; setErr(""); const name = f.name.toLowerCase();
    if (name.endsWith(".txt") || name.endsWith(".md") || f.type.startsWith("text")) {
      const r = new FileReader(); r.onload = () => { setDocName(f.name); onChange(String(r.result || "").slice(0, 8000)); }; r.readAsText(f);
    } else {
      // pdf / docx / doc — extract text server-side (shared Foundry service).
      setDocName(f.name); setWorking("extracting");
      try { const b64 = await fileToB64(f); const text = await api.aiExtract(b64, f.name); onChange(text || ""); if (!text) setErr("No readable text found in that file."); }
      catch (e) { setErr(e.message || "Couldn't read that document."); setDocName(null); }
      setWorking(null);
    }
  };
  const pickOD = async () => {
    setErr("");
    try {
      const picked = await pickOneDriveFile();
      if (!picked) return; // cancelled
      setDocName(picked.name); setWorking("extracting");
      const text = await api.aiExtract(picked.dataB64, picked.name);
      onChange(text || ""); if (!text) setErr("No readable text found in that file.");
    } catch (e) { setErr(e.message || "Couldn't read that file from OneDrive."); setDocName(null); }
    setWorking(null);
  };
  const clearDoc = () => { setDocName(null); };
  return (
    <div>
      <div className="mb-2 flex gap-1">{[["type", "Type", TypeIcon], ["voice", "Voice", Mic], ["document", "Attach / doc", FileText]].map(([k, l, I]) => <button key={k} onClick={() => setMode(k)} className={`inline-flex items-center gap-1 rounded-md px-2.5 py-1 text-xs font-medium ${mode === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}><I size={12} /> {l}</button>)}</div>
      {mode === "voice" && <div className="mb-2 flex items-center gap-2">{working === "transcribing" ? <span className="inline-flex items-center gap-1.5 text-sm text-slate-500"><Loader2 size={14} className="animate-spin" /> Transcribing…</span> : !rec ? <button onClick={start} className={btnLight}><Mic size={14} /> Record</button> : <button onClick={stop} className="inline-flex items-center gap-1.5 rounded-md bg-rose-500 px-3 py-2 text-sm font-medium text-white"><Square size={13} /> Stop</button>}<span className="text-xs text-slate-400">{CAP.deepgram ? "Deepgram speech-to-text · transcript is editable below." : "Transcript is editable below."}</span></div>}
      {mode === "document" && <div className="mb-2 flex flex-col gap-1.5 sm:flex-row">
        <label className="flex flex-1 cursor-pointer items-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50"><Upload size={14} /> From this computer — PDF, Word, .md or .txt<input type="file" accept=".pdf,application/pdf,.docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,.txt,.md,text/plain" className="hidden" onChange={(e) => readF(e.target.files && e.target.files[0])} /></label>
        {MSAL_CONFIGURED && <button onClick={pickOD} disabled={working === "extracting"} className="flex flex-1 items-center justify-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50 disabled:opacity-50"><Cloud size={14} /> From OneDrive</button>}
      </div>}
      {docName && <div className="mb-2 flex items-center gap-2 rounded-md border border-blue-200 bg-blue-50 px-2.5 py-1.5 text-xs text-blue-700"><FileText size={13} /> <span className="min-w-0 flex-1 truncate">{docName}</span>{working === "extracting" ? <Loader2 size={13} className="animate-spin" /> : <button onClick={clearDoc} className="text-blue-400 hover:text-blue-700"><X size={13} /></button>}</div>}
      {err && <div className="mb-2 rounded-md border border-amber-200 bg-amber-50 px-2.5 py-1.5 text-xs text-amber-700">{err}</div>}
      <textarea value={value} onChange={(e) => onChange(e.target.value)} rows={mode === "type" ? 4 : 5} placeholder={placeholder} className={inputCls} />
    </div>
  );
}

/* ---------- session + hash routing ---------- */
// Persist the signed-in user so a refresh doesn't bounce back to Login, and mirror
// the current view in the URL hash so refresh restores it and the browser back/
// forward buttons navigate between pages. Hash routing needs no server rewrites,
// so it behaves identically locally, on a LAN IP, and over an ngrok tunnel.
const SESSION_KEY = "cadence_session_user";
const loadSession = () => { try { return JSON.parse(localStorage.getItem(SESSION_KEY) || "null"); } catch { return null; } };
const ROUTE_TABS = ["portfolio", "data", "myday", "team", "approvals"];
const routeToHash = (tab, openId) => (openId ? `#/n/${encodeURIComponent(openId)}` : `#/${tab}`);
const hashToRoute = () => {
  const h = (typeof window !== "undefined" ? window.location.hash : "").replace(/^#\/?/, "");
  if (h.startsWith("n/")) return { tab: "portfolio", openId: decodeURIComponent(h.slice(2)) };
  const t = h.split(/[/?]/)[0];
  return { tab: ROUTE_TABS.includes(t) ? t : "portfolio", openId: null };
};

/* ==================================================================== */
export default function App() {
  const initialRoute = hashToRoute();
  const [me, setMe] = useState(loadSession);
  const [hydrating, setHydrating] = useState(() => !!loadSession());
  const [works, setWorks] = useState([]);
  const [acts, setActs] = useState([]);
  const [crs, setCrs] = useState([]);
  const [teams, setTeams] = useState([]);
  const [remarks, setRemarks] = useState([]);
  const [tab, setTab] = useState(initialRoute.tab);
  const [openId, setOpenId] = useState(initialRoute.openId);
  const [dataFocus, setDataFocus] = useState(null); // { patch, key } — drives the Data tab's tracker filter
  const [busy, setBusy] = useState(null);
  const [note, setNote] = useState(null);
  const [capture, setCapture] = useState(false);
  const [smartAdd, setSmartAdd] = useState(false);
  const [followUpsOpen, setFollowUpsOpen] = useState(false);
  const [portView, setPortView] = useState("scorecard");
  const [reviewMode, setReviewMode] = useState(false);
  const [nudgeOpen, setNudgeOpen] = useState(false);
  const [objModal, setObjModal] = useState(false);
  const [remarkNode, setRemarkNode] = useState(null);
  const [loading, setLoading] = useState(false);
  const pendingWrites = useRef(0);

  const isOrg = me && (me.level === "md" || me.level === "vp");
  const flash = (m) => { setNote(m); setTimeout(() => setNote(null), 4000); };
  const pending = crs.filter((c) => c.status === "pending");

  /* ---- live sync with the shared Cadence service (single source of truth) ---- */
  const applySnap = useCallback((snap) => {
    if (!snap) return;
    setDirectory(snap.users || []);
    setWorks(snap.works || []); setActs(snap.acts || []); setCrs(snap.crs || []);
    setTeams(snap.teams || []); setRemarks(snap.remarks || []);
  }, []);
  const refresh = useCallback(async () => { try { applySnap(await api.snapshot()); } catch { /* transient */ } }, [applySnap]);
  // optimistic local update, persist to service, then reconcile to server truth
  const persist = useCallback(async (localUpdate, apiCall) => {
    pendingWrites.current++;
    try { if (localUpdate) localUpdate(); const r = await apiCall(); if (r && r.snapshot) applySnap(r.snapshot); }
    catch (e) { flash(`Sync failed: ${e.message || e}`); await refresh(); }
    finally { pendingWrites.current--; }
  }, [applySnap, refresh]);

  const store = {
    addWorks: (ws) => persist(() => setWorks((p) => [...p, ...ws]), () => api.addWorks(ws)),
    patchWork: (id, p) => persist(() => setWorks((prev) => prev.map((w) => (w.id === id ? { ...w, ...p } : w))), () => api.patchWork(id, p)),
    patchWorks: (patches) => persist(() => setWorks((prev) => prev.map((w) => (patches[w.id] ? { ...w, ...patches[w.id] } : w))), async () => { let last; for (const id of Object.keys(patches)) last = await api.patchWork(id, patches[id]); return last; }),
    deleteWork: (id) => persist(null, () => api.deleteWork(id)),
    addActs: (as) => persist(() => setActs((p) => [...p, ...as]), () => api.addActs(as)),
    patchAct: (id, p) => persist(() => setActs((prev) => prev.map((a) => (a.id === id ? { ...a, ...p } : a))), () => api.patchAct(id, p)),
    patchActs: (patches) => persist(() => setActs((prev) => prev.map((a) => (patches[a.id] ? { ...a, ...patches[a.id] } : a))), async () => { let last; for (const id of Object.keys(patches)) last = await api.patchAct(id, patches[id]); return last; }),
    deleteAct: (id) => persist(() => setActs((prev) => prev.filter((a) => a.id !== id)), () => api.deleteAct(id)),
    addCr: (cr) => persist(() => setCrs((p) => [...p, cr]), () => api.addCr(cr)),
    decideCr: (id, body) => persist(null, () => api.decideApproval(id, body)),
    addTeam: (t) => persist(() => setTeams((p) => [...p, t]), () => api.addTeam(t)),
  };
  // preserve the names used throughout the component tree
  const patchAct = store.patchAct;
  const patchWork = store.patchWork;

  const goApprovals = () => { setTab("approvals"); setOpenId(null); };
  // Node detail always lives under the portfolio tab, so opening a node from
  // anywhere (Data, Team, cards) switches there. Passing null closes the node.
  const openNode = (id) => { setOpenId(id); setTab("portfolio"); };
  // Drill-in from the dashboard KPIs / health chart: jump to the Data tab with
  // the tracker pre-filtered. The timestamped key re-fires DataPage's effect
  // even when the same filter is chosen twice.
  const goData = (patch) => { setOpenId(null); setDataFocus({ patch, key: Date.now() }); setTab("data"); };
  const unreadNudges = me ? remarks.filter((r) => r.toIds.includes(me.id) && !r.readBy.includes(me.id)).length : 0;
  const followUpCount = me && works.length ? followUps(works, acts, me).length : 0;

  const addRemark = ({ text, newOwnerId, shiftDays, ops }) => {
    const node = remarkNode; if (!node) return;
    const chain = []; { let cur = node, g = 0; while (cur && g++ < 12) { chain.push(cur); if (cur.level === "initiative") break; cur = works.find((w) => w.id === cur.parentId); } }
    const toIds = [...new Set(chain.map((n) => n.ownerId).filter((id) => id && id !== me.id))];
    const ini = chain.find((n) => n.level === "initiative");
    const workPatches = {}; const actPatches = {}; const nw = []; const na = [];
    if (newOwnerId) workPatches[node.id] = { ...(workPatches[node.id] || {}), ownerId: newOwnerId };
    if (shiftDays) {
      const ids = subtreeIds(works, node.id);
      acts.forEach((a) => { if (ids.includes(a.workId) && a.date) actPatches[a.id] = { ...(actPatches[a.id] || {}), date: iso(addDays(parseISO(a.date), shiftDays)) }; });
      // move deadlines too: this node, its descendants, and the initiative it belongs to
      const dlIds = new Set([...ids, ...(ini ? [ini.id] : [])]);
      works.forEach((w) => { if (dlIds.has(w.id) && w.deadline) workPatches[w.id] = { ...(workPatches[w.id] || {}), deadline: iso(addDays(parseISO(w.deadline), shiftDays)) }; });
    }
    // apply AI-suggested plan changes to the people beneath (within this node's subtree)
    let opsCount = 0;
    if (ops && ops.length) {
      opsCount = ops.length;
      const directSubs = works.filter((w) => w.parentId === node.id); const container = directSubs.length ? directSubs : [node];
      ops.forEach((op) => { if (op.op === "add_work" || op.op === "add_subwork") nw.push({ id: nid("w"), parentId: node.id, level: "work", title: op.title, type: node.type || "general", ownerId: node.ownerId }); });
      const liveSubs = container.concat(nw);
      ops.forEach((op) => { if (op.op === "add_activity") { const key = (op.work || op.subwork || "").toLowerCase(); const sw = liveSubs.find((s) => s.title.toLowerCase().includes(key)) || liveSubs[0]; if (sw) na.push({ id: nid("a"), workId: sw.id, title: op.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(op.estimateHrs) || 2, actualHrs: null, actType: op.type || "self", unplanned: true }); } });
      const ids = subtreeIds(works, node.id);
      ops.forEach((op) => { if (op.op === "retype") acts.forEach((a) => { if (ids.includes(a.workId) && a.title.toLowerCase().includes((op.match || "").toLowerCase())) actPatches[a.id] = { ...(actPatches[a.id] || {}), actType: op.type }; }); });
    }
    const remark = { id: nid("r"), nodeId: node.id, level: node.level, title: node.title, fromId: me.id, toIds, text, changes: { newOwnerId: newOwnerId || null, shiftDays: shiftDays || 0, ops: opsCount }, ts: Date.now(), readBy: [] };
    if (nw.length) store.addWorks(nw);
    if (Object.keys(workPatches).length) store.patchWorks(workPatches);
    if (na.length) store.addActs(na);
    if (Object.keys(actPatches).length) store.patchActs(actPatches);
    persist(() => setRemarks((prev) => [remark, ...prev]), () => api.addRemark(remark));
    setRemarkNode(null);
    flash(toIds.length ? `Remark sent — ${toIds.length} owner${toIds.length !== 1 ? "s" : ""} nudged${opsCount ? `, ${opsCount} plan change${opsCount !== 1 ? "s" : ""} applied` : ""}.` : "Remark saved.");
  };
  const openNudges = () => { if (me) persist(() => setRemarks((prev) => prev.map((r) => (r.toIds.includes(me.id) && !r.readBy.includes(me.id) ? { ...r, readBy: [...r.readBy, me.id] } : r))), () => api.markRemarksRead(me.id)); setNudgeOpen(true); };

  const handleLogin = async (user) => { setLoading(true); try { applySnap(await api.snapshot()); } catch { flash("Couldn't reach the Cadence service — is it running on port 4000?"); } localStorage.setItem(SESSION_KEY, JSON.stringify(user)); setMe(user); setLoading(false); };
  const logout = () => { localStorage.removeItem(SESSION_KEY); if (typeof window !== "undefined") window.location.hash = ""; setMe(null); setTab("portfolio"); setOpenId(null); setWorks([]); setActs([]); setCrs([]); setTeams([]); setRemarks([]); };

  // probe service capabilities (voice mode) once
  useEffect(() => { api.health().then((h) => { CAP.deepgram = !!(h && h.deepgram); }).catch(() => {}); }, []);
  // poll the shared store so Teams changes appear here (and vice-versa)
  useEffect(() => {
    if (!me) return;
    const t = setInterval(() => { if (pendingWrites.current === 0) refresh(); }, 3500);
    return () => clearInterval(t);
  }, [me, refresh]);

  // On a refresh with a persisted session, `me` is already restored from storage —
  // reload the live data behind a brief splash so the page comes back where it was.
  useEffect(() => {
    let alive = true;
    if (loadSession()) refresh().finally(() => alive && setHydrating(false));
    else setHydrating(false);
    return () => { alive = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  // View -> URL: mirror the current page in the hash (creates history entries so
  // browser Back/Forward move between pages).
  useEffect(() => {
    if (!me) return;
    const desired = routeToHash(tab, openId);
    if (window.location.hash !== desired) window.location.hash = desired;
  }, [me, tab, openId]);
  // URL -> View: respond to Back/Forward and manual hash edits.
  useEffect(() => {
    if (!me) return;
    const onHash = () => { const r = hashToRoute(); setTab(r.tab); setOpenId(r.openId); };
    window.addEventListener("hashchange", onHash);
    return () => window.removeEventListener("hashchange", onHash);
  }, [me]);

  if (me && hydrating) return <div className="flex items-center justify-center bg-stone-50 text-sm text-slate-400" style={{ minHeight: 720 }}><Loader2 className="mr-2 animate-spin" size={18} /> Loading your workspace…</div>;
  if (!me) return <Login onLogin={handleLogin} />;
  const tops = works.filter((w) => w.parentId === null);
  const open = openId ? works.find((w) => w.id === openId) : null;
  const tabs = [["portfolio", me.level === "member" ? "My work" : "Portfolio", LayoutGrid], ["data", "Data", ClipboardList], ["myday", "My day", Calendar]];
  if (isOrg) tabs.push(["team", "Team", Users], ["approvals", "Approvals", ClipboardCheck]);

  return (
    <div className="w-full bg-stone-50 text-slate-800" style={{ minHeight: 720, fontFeatureSettings: '"tnum"' }}>
      <div className="flex items-center justify-between border-b border-slate-200 bg-white px-3 py-3 sm:px-5">
        <div className="flex min-w-0 items-center gap-2.5"><div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-md bg-slate-900 font-mono text-sm text-white">C</div><div className="min-w-0"><div className="text-sm font-medium leading-none text-slate-900">Cadence</div><div className="mt-0.5 hidden text-xs text-slate-400 sm:block">work &amp; initiative OS · prototype</div></div></div>
        <div className="flex shrink-0 items-center gap-2 sm:gap-3"><OneDriveConnect compact /><button onClick={() => setFollowUpsOpen(true)} className={`${btnLight} relative`} title="Follow-ups — email nudges"><Mail size={14} />{followUpCount > 0 && <span className="absolute -right-1.5 -top-1.5 rounded-full bg-rose-500 px-1 text-xs font-medium text-white">{followUpCount}</span>}</button><button onClick={openNudges} className={`${btnLight} relative`} title="Nudges"><MessageSquare size={14} />{unreadNudges > 0 && <span className="absolute -right-1.5 -top-1.5 rounded-full bg-rose-500 px-1 text-xs font-medium text-white">{unreadNudges}</span>}</button><div className="hidden items-center gap-2 sm:flex"><Avatar id={me.id} size={28} /><div className="text-right"><div className="text-xs font-medium text-slate-700">{me.name}</div><div className="text-xs text-slate-400">{me.title} · {me.level}</div></div></div><button onClick={logout} className={btnLight}><LogOut size={14} /></button></div>
      </div>
      <div className="flex flex-col gap-1.5 border-b border-slate-200 bg-white px-3 py-1.5 sm:flex-row sm:items-center sm:gap-1 sm:px-5 sm:py-0">
        <div className="-mx-1 flex items-center gap-1 overflow-x-auto px-1">
          {tabs.map(([k, l, I]) => <button key={k} onClick={() => { setTab(k); setOpenId(null); setDataFocus(null); }} className={`flex shrink-0 items-center gap-2 border-b-2 px-3 py-2.5 text-sm ${tab === k ? "border-slate-900 text-slate-900" : "border-transparent text-slate-400 hover:text-slate-600"}`}><I size={15} /><span className="font-medium">{l}</span>{k === "approvals" && pending.length > 0 && <span className="rounded-full bg-rose-500 px-1.5 text-xs font-medium text-white">{pending.length}</span>}</button>)}
        </div>
        {isOrg && <div className="flex items-center gap-2 sm:ml-auto sm:py-1.5">
          <button onClick={() => { setTab("portfolio"); setOpenId(null); setReviewMode(true); }} className={`${btnLight} flex-1 sm:flex-none`}><Pencil size={14} /> Review &amp; update</button>
          <button onClick={() => (me.level === "md" ? setObjModal(true) : setCapture(true))} className={`${btnDark} flex-1 sm:flex-none`}><Plus size={14} /> {me.level === "md" ? "New objective" : "New initiative"}</button>
        </div>}
      </div>
      {note && <div className="mx-3 mt-3 rounded-md border border-slate-200 bg-white px-3 py-2 text-xs text-slate-600 sm:mx-5">{note}</div>}
      <div className="p-3 sm:p-5">
        {tab === "portfolio" && !open && <Portfolio {...{ works, acts, crs, me, isOrg, teams, onOpen: openNode, onFocus: goData, goApprovals, view: portView, setView: setPortView, reviewMode, setReviewMode, onRemark: setRemarkNode, patchAct, flash }} />}
        {tab === "portfolio" && open && <NodeView {...{ nodeId: openId, user: me, works, acts, crs, teams, isOrg, busy, setBusy, flash, patchAct, patchWork, store, onOpen: openNode, onRemark: setRemarkNode, goApprovals }} />}
        {tab === "data" && <DataPage {...{ me, works, acts, onOpen: openNode, focus: dataFocus, patchAct, flash }} />}
        {tab === "myday" && <MyDay {...{ me, works, acts, busy, setBusy, flash, patchAct, store }} />}
        {tab === "team" && isOrg && <TeamView {...{ user: me, teams, store, works, acts, onOpen: openNode }} />}
        {tab === "approvals" && isOrg && <Approvals {...{ crs, store, works, me, flash }} />}
      </div>
      {capture && <Capture {...{ me, teams, store, works, busy, setBusy, flash, onClose: () => setCapture(false), onOpen: (id) => { setCapture(false); setTab("portfolio"); setOpenId(id); } }} />}
      {objModal && <QuickCreate {...{ me, parent: null, level: "objective", works, store, busy, setBusy, flash, onClose: () => setObjModal(false), onCreated: (id) => { setTab("portfolio"); setPortView("scorecard"); setOpenId(id); } }} />}
      {nudgeOpen && <NudgeInbox {...{ remarks, me, onOpen: (id) => { setNudgeOpen(false); setTab("portfolio"); setOpenId(id); }, onClose: () => setNudgeOpen(false) }} />}
      {remarkNode && <RemarkModal {...{ node: remarkNode, works, acts, busy, setBusy, flash, onSubmit: addRemark, onClose: () => setRemarkNode(null) }} />}
      {smartAdd && <SmartAdd {...{ me, works, teams, store, busy, setBusy, flash, onClose: () => setSmartAdd(false), onCreated: (id) => { setSmartAdd(false); if (id) openNode(id); } }} />}
      {followUpsOpen && <FollowUps {...{ me, works, acts, busy, setBusy, flash, onClose: () => setFollowUpsOpen(false), onOpen: (id) => { setFollowUpsOpen(false); openNode(id); } }} />}
      {/* Universal add — dictate/write/attach, AI decides what & where. */}
      <button onClick={() => setSmartAdd(true)} title="Add an objective, initiative, work or task" className="fixed bottom-5 right-5 z-20 inline-flex items-center gap-2 rounded-full bg-slate-900 px-4 py-3 text-sm font-medium text-white shadow-lg transition hover:bg-slate-700 hover:shadow-xl"><Plus size={18} /><span className="hidden sm:inline">Add</span></button>
      {MSAL_CONFIGURED && <OneDrivePickerHost />}
    </div>
  );
}

/* ---------- Login ---------- */
const LOGIN_HINTS = [
  { username: "vikram", password: "md@2026", role: "MD · md" },
  { username: "meera", password: "vp@2026", role: "VP, Supply Chain · vp" },
  { username: "priya", password: "vp@2026", role: "VP, IT · vp" },
  { username: "rohit", password: "team@2026", role: "Executive, IT · member" },
  { username: "neha", password: "team@2026", role: "Executive, SC · member" },
];
function Login({ onLogin }) {
  const [u, setU] = useState(""); const [p, setP] = useState(""); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const go = async () => { setErr(""); setBusy(true); try { const user = await api.login(u.trim(), p); onLogin(user); } catch (e) { setErr(e.status === 401 ? "Wrong username or password." : (e.message || "Couldn't reach the Cadence service.")); } setBusy(false); };
  return (
    <div className="flex w-full items-center justify-center bg-stone-50 p-6" style={{ minHeight: 720 }}>
      <div className="w-full max-w-sm">
        <div className="mb-6 flex items-center gap-2.5"><div className="flex h-9 w-9 items-center justify-center rounded-md bg-slate-900 font-mono text-white">C</div><div><div className="text-lg font-medium text-slate-900">Cadence</div><div className="text-xs text-slate-400">sign in</div></div></div>
        <div className="rounded-xl border border-slate-200 bg-white p-5">
          <label className="mb-1 block text-xs text-slate-500">Username</label><div className="mb-3 flex items-center gap-2 rounded-md border border-slate-200 px-2"><User size={15} className="text-slate-400" /><input value={u} onChange={(e) => setU(e.target.value)} className="w-full py-2 text-sm outline-none" placeholder="username" /></div>
          <label className="mb-1 block text-xs text-slate-500">Password</label><div className="mb-4 flex items-center gap-2 rounded-md border border-slate-200 px-2"><Lock size={15} className="text-slate-400" /><input type="password" value={p} onChange={(e) => setP(e.target.value)} onKeyDown={(e) => e.key === "Enter" && go()} className="w-full py-2 text-sm outline-none" placeholder="password" /></div>
          {err && <div className="mb-3 text-xs text-rose-600">{err}</div>}<button onClick={go} disabled={busy} className={`${btnDark} w-full`}>{busy ? <><Loader2 size={14} className="animate-spin" /> Signing in…</> : "Sign in"}</button>
        </div>
        <div className="mt-4 rounded-xl border border-dashed border-slate-300 bg-white p-4"><div className="mb-2 text-xs font-medium text-slate-500">Test accounts</div><div className="space-y-1 text-xs text-slate-600">{LOGIN_HINTS.map((x) => <div key={x.username} className="flex justify-between"><span className="font-mono">{x.username} / {x.password}</span><span className="text-slate-400">{x.role}</span></div>)}</div></div>
      </div>
    </div>
  );
}

/* ---------- Portfolio (executive scorecard + rich cards) ---------- */
function Tile({ label, value, tone }) { return <div className="rounded-xl bg-white border border-slate-200 px-4 py-3"><div className="text-xs text-slate-400">{label}</div><div className={`mt-1 font-mono text-2xl font-medium ${tone || "text-slate-900"}`}>{value}</div></div>; }
function Panel({ title, right, children, fill }) { return <div className={`rounded-xl border border-slate-200 bg-white p-4 ${fill ? "flex h-full flex-col" : ""}`}><div className="mb-3 flex items-center justify-between"><div className="text-sm font-medium text-slate-700">{title}</div>{right}</div>{children}</div>; }
// Close a popover when clicking outside it.
function useOutside(onClose) {
  const ref = useRef(null);
  useEffect(() => { const h = (e) => { if (ref.current && !ref.current.contains(e.target)) onClose(); }; document.addEventListener("mousedown", h); return () => document.removeEventListener("mousedown", h); }, [onClose]);
  return ref;
}
// Fully in-house dropdown (no native <select>) so the open menu matches the portal.
function Dropdown({ value, onChange, options, align = "right" }) {
  const [open, setOpen] = useState(false);
  const ref = useOutside(() => setOpen(false));
  const cur = options.find((o) => o.value === value);
  return (
    <div ref={ref} className="relative">
      <button onClick={() => setOpen((o) => !o)} className="inline-flex items-center gap-2 rounded-lg border border-slate-200 bg-white py-1.5 pl-3 pr-2 text-xs font-medium text-slate-700 shadow-sm outline-none transition hover:border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100">
        {cur ? cur.label : "Select"}<ChevronRight size={13} className={`text-slate-400 transition-transform ${open ? "-rotate-90" : "rotate-90"}`} />
      </button>
      {open && <div className={`absolute z-30 mt-1 min-w-[9rem] overflow-hidden rounded-lg border border-slate-200 bg-white py-1 shadow-lg ${align === "right" ? "right-0" : "left-0"}`}>
        {options.map((o) => <button key={o.value} onClick={() => { onChange(o.value); setOpen(false); }} className={`block w-full px-3 py-1.5 text-left text-xs ${o.value === value ? "bg-blue-50 font-medium text-blue-800" : "text-slate-600 hover:bg-slate-50"}`}>{o.label}</button>)}
      </div>}
    </div>
  );
}
// Custom calendar picker (no native date input) — its own styled popup.
function DatePicker({ value, onChange, placeholder = "Any date", align = "right" }) {
  const [open, setOpen] = useState(false);
  const ref = useOutside(() => setOpen(false));
  const [view, setView] = useState(() => (value ? sod(parseISO(value)) : TODAY));
  const y = view.getFullYear(), m = view.getMonth();
  const startDow = new Date(y, m, 1).getDay();
  const days = new Date(y, m + 1, 0).getDate();
  const cells = [...Array(startDow).fill(null), ...Array.from({ length: days }, (_, i) => i + 1)];
  const sel = value ? iso(parseISO(value)) : null;
  return (
    <div ref={ref} className="relative">
      <button onClick={() => setOpen((o) => !o)} className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white py-1.5 pl-2.5 pr-3 text-xs font-medium text-slate-700 shadow-sm outline-none transition hover:border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100">
        <CalendarClock size={13} className="text-slate-400" />{value ? fmtFull(parseISO(value)) : <span className="text-slate-400">{placeholder}</span>}
      </button>
      {open && <div className={`absolute z-30 mt-1 w-64 rounded-xl border border-slate-200 bg-white p-3 shadow-lg ${align === "right" ? "right-0" : "left-0"}`}>
        <div className="mb-2 flex items-center justify-between">
          <button onClick={() => setView(new Date(y, m - 1, 1))} className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><ChevronLeft size={16} /></button>
          <div className="text-sm font-medium text-slate-800">{view.toLocaleDateString("en-GB", { month: "long", year: "numeric" })}</div>
          <button onClick={() => setView(new Date(y, m + 1, 1))} className="rounded p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><ChevronRight size={16} /></button>
        </div>
        <div className="grid grid-cols-7 gap-0.5 text-center text-[10px] font-medium text-slate-400">{["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"].map((d) => <div key={d}>{d}</div>)}</div>
        <div className="mt-0.5 grid grid-cols-7 gap-0.5">
          {cells.map((c, i) => c == null ? <div key={i} /> : (() => { const iso1 = iso(new Date(y, m, c)); const isSel = iso1 === sel, isToday = iso1 === iso(TODAY);
            return <button key={i} onClick={() => { onChange(iso1); setOpen(false); }} className={`rounded-md py-1 text-xs ${isSel ? "bg-blue-700 font-medium text-white" : isToday ? "bg-blue-50 font-medium text-blue-800" : "text-slate-700 hover:bg-slate-100"}`}>{c}</button>; })())}
        </div>
        <div className="mt-2 flex items-center justify-between border-t border-slate-100 pt-2 text-xs">
          <button onClick={() => { onChange(""); setOpen(false); }} className="font-medium text-slate-500 hover:text-slate-800">Clear</button>
          <button onClick={() => { onChange(iso(TODAY)); setView(TODAY); setOpen(false); }} className="font-medium text-blue-700 hover:text-blue-900">Today</button>
        </div>
      </div>}
    </div>
  );
}


// Team-level progress + capacity in one card: execution done% bar (the headline),
// plus open-hours load, a Free/Busy chip, member avatars, and roll-up stats.
function TeamPerformance({ teams, works, acts, me, teamFilter }) {
  const CAP = 20; // planned open-hours per person before a team reads "busy"
  let inScope = me.level === "md" ? teams : (teams || []).filter((t) => t.memberIds.some((id) => fnOf(id) === me.fn));
  if (teamFilter) inScope = inScope.filter((t) => teamFilter.has(t.id)); // cross-filter
  const rows = inScope.map((t) => {
    const cap = t.memberIds.length * CAP;
    const load = t.memberIds.reduce((s, id) => s + loadOf(acts, id), 0);
    const ta = acts.filter((a) => t.memberIds.includes(a.assigneeId) && a.status !== "cancelled");
    const done = ta.filter((a) => a.status === "executed").length;
    const exec = ta.length ? Math.round((done / ta.length) * 100) : 0;
    const open = ta.filter((a) => a.status !== "executed").length;
    const overdue = ta.filter(isOverdue).length;
    const inits = works.filter((w) => w.level === "initiative" && w.teamId === t.id).length;
    return { t, cap, load, exec, open, overdue, inits, busy: cap ? load > cap * 0.8 : false };
  }).sort((a, b) => b.exec - a.exec);
  return (
    <Panel title="Team performance" right={teamFilter && <span className="text-[10px] font-medium text-blue-800">filtered</span>}>
      <div className="space-y-2">
        {rows.length === 0 && <div className="text-xs text-slate-400">{teamFilter ? "No teams in the current selection." : "No teams in scope."}</div>}
        {rows.map(({ t, cap, load, exec, open, overdue, inits, busy }) => (
          <div key={t.id} className="rounded-lg border border-slate-100 p-2">
            <div className="mb-1 flex items-center gap-2">
              <span className="truncate text-sm font-medium text-slate-800">{t.name}</span>
              <span className="flex -space-x-1.5">{t.memberIds.slice(0, 5).map((id) => <span key={id} className="rounded-full ring-2 ring-white"><Avatar id={id} size={16} /></span>)}</span>
              <span className="ml-auto flex shrink-0 items-center gap-1">
                {overdue > 0 && <span className="rounded-full bg-rose-50 px-1.5 py-0.5 text-[10px] font-medium text-rose-700">{overdue} overdue</span>}
                <span className={`rounded-full px-2 py-0.5 text-[10px] font-medium ${busy ? "bg-amber-50 text-amber-700" : "bg-emerald-50 text-emerald-700"}`}>{busy ? "Busy" : "Free"}</span>
              </span>
            </div>
            <div className="flex items-center gap-2">
              <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-blue-700" style={{ width: `${exec}%` }} /></div>
              <span className="w-16 shrink-0 text-right font-mono text-xs text-slate-500">{exec}% done</span>
            </div>
            <div className="mt-1 text-[11px] text-slate-400">{inits} initiative{inits !== 1 ? "s" : ""} · {open} open · {load}/{cap}h · {t.memberIds.length} people</div>
          </div>
        ))}
      </div>
    </Panel>
  );
}

// Ranked highlights that switch category via a dropdown: who's executing, who's
// planning best, and which teams are ahead. Scoped to the viewer.
function Leaderboard({ works, acts, teams, me, userFilter }) {
  const [cat, setCat] = useState("finishers");
  let perfUsers = me.level === "md" ? USERS.filter((u) => u.level !== "md") : USERS.filter((u) => u.fn === me.fn && u.id !== me.id);
  if (userFilter) perfUsers = perfUsers.filter((u) => userFilter.has(u.id)); // cross-filter
  const stats = perfUsers.map((u) => {
    const mine = acts.filter((a) => a.assigneeId === u.id && a.status !== "cancelled");
    const done = mine.filter((a) => a.status === "executed").length;
    const overdue = mine.filter((a) => isOverdue(a)).length;
    const blocked = mine.filter((a) => a.blocked && a.status !== "executed").length;
    const owned = works.filter((w) => w.ownerId === u.id && (w.level === "initiative" || w.level === "work"));
    const planning = owned.length ? Math.round(owned.reduce((s, w) => s + computeMeters(works, acts, w.id).planning, 0) / owned.length) : null;
    return { id: u.id, name: u.name, done, total: mine.length, execPct: mine.length ? Math.round((done / mine.length) * 100) : 0, planning, owns: owned.length, overdue, blocked };
  });
  let teamScope = me.level === "md" ? (teams || []) : (teams || []).filter((t) => t.memberIds.some((id) => fnOf(id) === me.fn));
  if (userFilter) teamScope = teamScope.filter((t) => t.memberIds.some((id) => userFilter.has(id)));
  const teamStats = teamScope.map((t) => { const ta = acts.filter((a) => t.memberIds.includes(a.assigneeId) && a.status !== "cancelled"); const done = ta.filter((a) => a.status === "executed").length; return { id: t.id, name: t.name, done, total: ta.length, execPct: ta.length ? Math.round((done / ta.length) * 100) : 0 }; });
  const CATS = {
    finishers: { label: "Top finishers", rows: () => stats.filter((s) => s.total > 0).sort((a, b) => (b.done * b.execPct) - (a.done * a.execPct) || b.done - a.done).map((s) => ({ id: s.id, kind: "user", name: s.name, pct: s.execPct, text: `${s.done} done · ${s.execPct}%` })) },
    needsAttention: { label: "Needs attention", tone: "attn", rows: () => stats.filter((s) => s.total > 0 && (s.overdue > 0 || s.blocked > 0 || s.execPct < 100)).sort((a, b) => ((b.overdue + b.blocked) - (a.overdue + a.blocked)) || (a.execPct - b.execPct) || ((b.total - b.done) - (a.total - a.done))).map((s) => ({ id: s.id, kind: "user", name: s.name, pct: 100 - s.execPct, text: (s.overdue || s.blocked) ? `${s.overdue} overdue${s.blocked ? ` · ${s.blocked} blocked` : ""}` : `${s.execPct}% done · ${s.total - s.done} open` })) },
    planners: { label: "Top planners", rows: () => stats.filter((s) => s.planning != null).sort((a, b) => b.planning - a.planning).map((s) => ({ id: s.id, kind: "user", name: s.name, pct: s.planning, text: `${s.planning}% planned · ${s.owns} owned` })) },
    underplanned: { label: "Under-planned", tone: "attn", rows: () => stats.filter((s) => s.planning != null).sort((a, b) => a.planning - b.planning).map((s) => ({ id: s.id, kind: "user", name: s.name, pct: 100 - s.planning, text: `${s.planning}% planned · ${s.owns} owned` })) },
    teams: { label: "Best teams", rows: () => teamStats.slice().sort((a, b) => b.execPct - a.execPct).map((s) => ({ id: s.id, kind: "team", name: s.name, pct: s.execPct, text: `${s.execPct}% done · ${s.done}/${s.total}` })) },
    teamsRisk: { label: "Teams at risk", tone: "attn", rows: () => teamStats.filter((s) => s.total > 0).slice().sort((a, b) => a.execPct - b.execPct).map((s) => ({ id: s.id, kind: "team", name: s.name, pct: 100 - s.execPct, text: `${s.execPct}% done · ${s.done}/${s.total}` })) },
  };
  const attn = CATS[cat].tone === "attn";
  const rows = CATS[cat].rows().slice(0, 5);
  // Ranked emphasis for the top three: a gold-podium for the "top" lists, a
  // rose "concern" scale for the needs-attention lists; plain for the rest.
  const PODIUM = attn ? [
    { row: "border-rose-300 bg-rose-50", badge: "bg-rose-500 text-white" },
    { row: "border-rose-200 bg-rose-50", badge: "bg-rose-400 text-white" },
    { row: "border-amber-200 bg-amber-50", badge: "bg-amber-500 text-white" },
  ] : [
    { row: "border-amber-200 bg-amber-50", badge: "bg-amber-400 text-white" },
    { row: "border-slate-300 bg-slate-100", badge: "bg-slate-400 text-white" },
    { row: "border-orange-200 bg-orange-50", badge: "bg-orange-400 text-white" },
  ];
  return (
    <Panel fill title="Leaderboard" right={<Dropdown value={cat} onChange={setCat} options={Object.entries(CATS).map(([k, v]) => ({ value: k, label: v.label }))} />}>
      <div className="flex flex-1 flex-col gap-1.5">
        {rows.length === 0 && <div className="text-xs text-slate-400">Not enough data to rank yet.</div>}
        {rows.map((r, i) => { const p = PODIUM[i];
          return (
            <div key={r.id} className={`flex flex-1 items-center gap-2.5 rounded-lg border px-2.5 ${p ? p.row : "border-slate-100"}`}>
              <span className={`flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-bold ${p ? p.badge : "bg-slate-100 text-slate-400"}`}>{i + 1}</span>
              {r.kind === "user" ? <Avatar id={r.id} size={24} /> : <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-white"><Users size={13} className="text-slate-500" /></span>}
              <div className="min-w-0 flex-1"><div className="truncate text-sm font-medium text-slate-800">{r.name}</div><div className="truncate text-xs text-slate-500">{r.text}</div></div>
              <div className="hidden w-20 shrink-0 sm:block"><div className="h-1.5 overflow-hidden rounded-full bg-white/70"><div className={`h-full rounded-full ${attn ? "bg-rose-500" : "bg-blue-700"}`} style={{ width: `${Math.max(0, Math.min(100, r.pct))}%` }} /></div></div>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}


/* ==================================================================== */
/* Portfolio dashboard — one health/analytics view + a filterable,        */
/* level-tagged tracker, so overdue / blocked / gap items are always      */
/* shown as *what* they are (objective / initiative / work / activity).   */
/* ==================================================================== */
// A single status vocabulary spanning every level, so charts + tracker agree.
const TRACK_STATUS = {
  overdue: { label: "Overdue", dot: "bg-rose-500", text: "text-rose-700", soft: "bg-rose-50" },
  blocked: { label: "Blocked", dot: "bg-rose-400", text: "text-rose-700", soft: "bg-rose-50" },
  atrisk: { label: "At risk", dot: "bg-amber-500", text: "text-amber-700", soft: "bg-amber-50" },
  ontrack: { label: "On track", dot: "bg-emerald-500", text: "text-emerald-700", soft: "bg-emerald-50" },
  done: { label: "Done", dot: "bg-blue-500", text: "text-blue-700", soft: "bg-blue-50" },
};
const STATUS_KEYS = ["overdue", "blocked", "atrisk", "ontrack", "done"];
const LEVEL_PLURAL = { objective: "Objectives", initiative: "Initiatives", work: "Work items", activity: "Activities" };
// Hex mirrors of the TRACK_STATUS dot colours, for SVG donut strokes.
const STATUS_HEX = { overdue: "#f43f5e", blocked: "#fb7185", atrisk: "#f59e0b", ontrack: "#10b981", done: "#3b82f6" };
// Roll-up status for a work/initiative/objective (blockers/overdue win, then RAG).
function nodeStatus(works, acts, id) {
  const att = attentionCount(works, acts, id);
  if (att.overdue > 0) return "overdue";
  if (att.blocked > 0) return "blocked";
  return nodeRag(works, acts, id) === "amber" ? "atrisk" : "ontrack";
}
// Status for a single activity (leaf).
// Fraction of an activity's start→end window that has elapsed (0..1, null if no span).
function scheduleFraction(a) {
  if (!a.startDate || !a.date) return null;
  const s = parseISO(a.startDate), e = parseISO(a.date); const span = e - s;
  if (span <= 0) return null;
  return (TODAY - s) / span;
}
function activityStatus(a) {
  if (a.status === "executed") return "done";
  if (a.blocked) return "blocked";
  if (isOverdue(a)) return "overdue";
  const d = daysLeft(a.date);
  if (d != null && d <= 3) return "atrisk"; // due within 3 days
  // Behind schedule: well into its start→end window but not yet started (no effort).
  const frac = scheduleFraction(a);
  if (frac != null && frac >= 0.6 && !a.actualHrs && !a.inProgress) return "atrisk";
  return "ontrack";
}
const emptyCounts = () => ({ overdue: 0, blocked: 0, atrisk: 0, ontrack: 0, done: 0 });
const statusCountsFromItems = (rows) => { const c = emptyCounts(); rows.forEach((r) => { if (c[r.status] !== undefined) c[r.status]++; }); return c; };
const statusCountsOfNodes = (nodes, works, acts) => { const c = emptyCounts(); nodes.forEach((n) => c[nodeStatus(works, acts, n.id)]++); return c; };

// Flattens the user's scope into one uniform, level-tagged row model.
function trackerRows(works, acts, me) {
  const member = me.level === "member";
  const objs = member ? [] : scopedObjectives(works, acts, me);
  const inits = member ? [] : scopeInitiatives(works, acts, me);
  const initIds = new Set(inits.map((i) => i.id));
  const wks = member ? homeNodes(works, acts, me).filter((w) => w.level === "work")
    : works.filter((w) => w.level === "work" && initIds.has(w.parentId));
  const workIds = new Set(wks.map((w) => w.id));
  const nameOf = (id) => works.find((w) => w.id === id)?.title || null;
  const rows = [];
  const nodeRow = (w) => {
    const m = computeMeters(works, acts, w.id);
    const att = attentionCount(works, acts, w.id);
    const status = nodeStatus(works, acts, w.id);
    const gap = (w.level === "initiative" && m.resultPct != null && m.resultPct < m.execution - 15) || (w.level === "work" && m.planning < 70);
    return {
      id: w.id, level: w.level, title: w.title, parentTitle: nameOf(w.parentId), ownerId: w.ownerId,
      status, due: w.deadline || null, execution: Math.round(m.execution), planning: Math.round(m.planning),
      resultPct: m.resultPct != null ? Math.round(m.resultPct) : null, result: w.result || null,
      childCount: works.filter((x) => x.parentId === w.id).length, overdue: att.overdue, blocked: att.blocked,
      span: nodeSpan(works, acts, w), effort: effortRollup(works, acts, w.id),
      openId: w.id, flags: { overdue: att.overdue > 0, blocked: att.blocked > 0, atrisk: status === "atrisk", ontrack: status === "ontrack", done: false, gap },
    };
  };
  objs.forEach((o) => rows.push(nodeRow(o)));
  inits.forEach((i) => rows.push(nodeRow(i)));
  wks.forEach((w) => rows.push(nodeRow(w)));
  acts.filter((a) => workIds.has(a.workId) && a.status !== "cancelled").forEach((a) => {
    const status = activityStatus(a);
    rows.push({
      id: a.id, level: "activity", title: a.title, parentTitle: nameOf(a.workId), ownerId: a.assigneeId,
      status, due: a.date || null, execution: a.status === "executed" ? 100 : 0, planning: null, resultPct: null,
      result: null, hrs: a.plannedHrs, openId: a.workId,
      span: { start: a.startDate || null, end: a.date || null }, effort: { planned: a.plannedHrs || 0, actual: a.actualHrs || 0 },
      flags: { overdue: status === "overdue", blocked: status === "blocked", atrisk: status === "atrisk", ontrack: status === "ontrack", done: status === "done", gap: false },
    });
  });
  return rows;
}

/* ---------- small dashboard atoms ---------- */
function StatusTag({ status }) { const s = TRACK_STATUS[status] || TRACK_STATUS.ontrack; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${s.soft} ${s.text}`}><span className={`h-1.5 w-1.5 rounded-full ${s.dot}`} />{s.label}</span>; }
function MiniBar({ pct, tone = "bg-teal-500" }) { return <div className="h-1.5 w-full overflow-hidden rounded-full bg-slate-100"><div className={`h-full rounded-full ${tone}`} style={{ width: `${Math.max(0, Math.min(100, pct || 0))}%` }} /></div>; }
function SegBar({ counts, total, onSeg }) {
  const sum = total || STATUS_KEYS.reduce((s, k) => s + (counts[k] || 0), 0) || 1;
  return <div className="flex h-2.5 overflow-hidden rounded-full bg-slate-100">{STATUS_KEYS.map((k) => counts[k] ? <button key={k} onClick={onSeg ? () => onSeg(k) : undefined} title={`${counts[k]} ${TRACK_STATUS[k].label}`} className={`${TRACK_STATUS[k].dot} ${onSeg ? "cursor-pointer" : "cursor-default"}`} style={{ width: `${(counts[k] / sum) * 100}%` }} /> : null)}</div>;
}
function KpiCard({ label, value, sub, tone = "text-slate-900", onClick, children }) {
  const cls = `rounded-xl border border-slate-200 bg-white p-3 ${onClick ? "text-left transition hover:border-slate-300 hover:shadow-sm" : ""}`;
  const inner = <><div className="text-xs text-slate-400">{label}</div><div className={`mt-1 font-mono text-2xl font-medium ${tone}`}>{value}</div>{sub && <div className="mt-0.5 truncate text-xs text-slate-400">{sub}</div>}{children}</>;
  return onClick ? <button onClick={onClick} className={cls}>{inner}</button> : <div className={cls}>{inner}</div>;
}

function HealthStrip({ items, onFocus }) {
  const objs = items.filter((i) => i.level === "objective");
  const inits = items.filter((i) => i.level === "initiative");
  const acts = items.filter((i) => i.level === "activity");
  const oc = statusCountsFromItems(objs);
  const attn = oc.overdue + oc.blocked + oc.atrisk;
  const withResult = inits.filter((i) => i.resultPct != null);
  const avgTarget = withResult.length ? Math.round(withResult.reduce((s, i) => s + i.resultPct, 0) / withResult.length) : null;
  const done = acts.filter((a) => a.status === "done").length;
  const execPct = acts.length ? Math.round((done / acts.length) * 100) : 0;
  const overdue = acts.filter((a) => a.status === "overdue").length;
  const blocked = acts.filter((a) => a.status === "blocked").length;
  // Effort roll-up across activities (uses the new estimated/actual hours).
  const planned = acts.reduce((s, a) => s + (a.effort ? a.effort.planned : 0), 0);
  const actual = acts.reduce((s, a) => s + (a.effort ? a.effort.actual : 0), 0);
  const overBudget = planned > 0 && actual > planned;
  return (
    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
      <KpiCard label={objs.length === 1 ? "Objective" : "Objectives"} value={objs.length} sub={`${oc.ontrack} on track · ${attn} to watch`}>
        {/* <div className="mt-2"><SegBar counts={oc} total={objs.length} /></div> */}
      </KpiCard>
      <KpiCard label="Initiatives to target" value={avgTarget != null ? `${avgTarget}%` : "—"} sub="avg result vs target" tone="text-blue-800" />
      <KpiCard label="Execution" value={`${execPct}%`} sub={`${done}/${acts.length} activities done`} tone="text-amber-700" />
      <KpiCard label="Effort" value={`${actual}/${planned}h`} sub={overBudget ? "over estimate ▲" : "actual vs estimate"} tone={overBudget ? "text-rose-600" : "text-slate-900"} />
      <KpiCard label="Overdue" value={overdue} sub="activities past due" tone={overdue ? "text-rose-600" : "text-slate-900"} onClick={() => onFocus({ level: "activity", status: "overdue" })} />
      <KpiCard label="Blocked" value={blocked} sub="need help — view" tone={blocked ? "text-rose-600" : "text-slate-900"} onClick={() => onFocus({ level: "activity", status: "blocked" })} />
    </div>
  );
}

// A donut of the status split for one level. Clicking a count cross-filters the
// whole dashboard (it does NOT navigate away). When another donut drives the
// filter this one shows the in-scope subset; when it IS the selected donut the
// chosen slice is highlighted and the rest dim.
function StatusDonut({ level, counts, total, onSelect, selectedStatus, idx = 0 }) {
  const R = 40, SW = 15, C = 2 * Math.PI * R;
  const segs = STATUS_KEYS.filter((k) => counts[k] > 0);
  let acc = 0;
  const arcs = segs.map((k) => { const dash = (counts[k] / total) * C; const a = { k, dash, off: acc }; acc += dash; return a; });
  const divider = `${idx % 2 === 1 ? "border-l border-slate-200 " : ""}${idx > 0 ? "sm:border-l sm:border-slate-200" : ""}`;
  return (
    <div className={`flex flex-col items-center px-2 ${divider}`}>
      <div className="mb-1.5 inline-flex items-center gap-1.5 text-xs font-medium text-slate-600"><span className={`h-2 w-2 rounded-full ${LEVEL_THEME[level].bar}`} />{LEVEL_PLURAL[level]}</div>
      <div className="relative">
        <svg width="125" height="125" viewBox="0 0 100 100">
          <g transform="rotate(-90 50 50)">
            {total === 0 ? <circle r={R} cx="50" cy="50" fill="none" stroke="#e2e8f0" strokeWidth={SW} />
              : arcs.map((a) => <circle key={a.k} r={R} cx="50" cy="50" fill="none" stroke={STATUS_HEX[a.k]} strokeWidth={SW} strokeDasharray={`${a.dash} ${C - a.dash}`} strokeDashoffset={-a.off} opacity={selectedStatus && selectedStatus !== a.k ? 0.2 : 1} />)}
          </g>
          <text x="50" y="50" textAnchor="middle" dominantBaseline="central" className="fill-slate-800 font-mono" style={{ fontSize: 24, fontWeight: 500 }}>{total}</text>
        </svg>
      </div>
      <div className="mt-1.5 flex max-w-[8.5rem] flex-wrap justify-center gap-x-1.5 gap-y-0.5 text-[11px]">
        {segs.length === 0 && <span className="text-slate-300">none</span>}
        {segs.map((k) => { const on = selectedStatus === k;
          return <button key={k} onClick={() => onSelect(k)} title={`${TRACK_STATUS[k].label} — cross-filter the dashboard`} className={`inline-flex items-center gap-1 rounded-full px-1 transition ${on ? "bg-slate-900 font-medium text-white" : "text-slate-500 hover:bg-slate-100 hover:text-slate-900"}`}><span className={`h-2 w-2 rounded-full ${TRACK_STATUS[k].dot}`} />{counts[k]}</button>;
        })}
      </div>
    </div>
  );
}
function LevelHealthChart({ items, crossFilter, scopeNodeIds, onSelect }) {
  const levels = ["objective", "initiative", "work", "activity"].filter((l) => items.some((i) => i.level === l));
  return (
    <Panel title="Portfolio health — status by level" right={<span className="hidden text-xs text-slate-400 sm:inline">click a count to cross-filter</span>}>
      <div className="grid grid-cols-2 gap-x-2 gap-y-6 py-3 sm:grid-cols-4">
        {levels.map((l, idx) => {
          const levelItems = items.filter((i) => i.level === l);
          // The selected donut shows its full split (highlighted); the others show
          // only the items connected to the current selection.
          const shown = !crossFilter || crossFilter.level === l ? levelItems
            : levelItems.filter((i) => scopeNodeIds && scopeNodeIds.has(l === "activity" ? i.openId : i.id));
          const selectedStatus = crossFilter && crossFilter.level === l ? crossFilter.status : null;
          return <StatusDonut key={l} idx={idx} level={l} counts={statusCountsFromItems(shown)} total={shown.length} selectedStatus={selectedStatus} onSelect={(status) => onSelect(l, status)} />;
        })}
      </div>
      <div className="mt-4 flex flex-wrap justify-center gap-x-3 gap-y-1 border-t border-slate-100 pt-3 text-xs text-slate-400">{STATUS_KEYS.map((k) => <span key={k} className="inline-flex items-center gap-1"><span className={`h-2 w-2 rounded-full ${TRACK_STATUS[k].dot}`} />{TRACK_STATUS[k].label}</span>)}</div>
    </Panel>
  );
}

// Expandable Objective → Initiative → Work cascade: the "flow" that reads top-down.
function ObjectiveFlow({ works, acts, me, onOpen }) {
  const objectives = scopedObjectives(works, acts, me);
  const scopedInits = scopeInitiatives(works, acts, me);
  const [open, setOpen] = useState(() => new Set());
  const toggle = (id) => setOpen((s) => { const n = new Set(s); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const initsOf = (oid) => scopedInits.filter((i) => i.parentId === oid);
  const worksOf = (iid) => works.filter((w) => w.level === "work" && w.parentId === iid);
  if (!objectives.length) return null;
  return (
    <Panel title="Objective progress → initiatives → work">
      <div className="space-y-2">
        {objectives.map((o) => {
          const om = computeMeters(works, acts, o.id); const inits = initsOf(o.id); const oOpen = open.has(o.id);
          return (
            <div key={o.id} className="overflow-hidden rounded-lg border border-slate-200">
              <div className="flex items-center gap-2 p-2.5 sm:gap-3">
                <button onClick={() => toggle(o.id)} className="shrink-0 rounded p-0.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><ChevronRight size={16} className={`transition-transform ${oOpen ? "rotate-90" : ""}`} /></button>
                <button onClick={() => onOpen(o.id)} className="min-w-0 flex-1 text-left">
                  <div className="flex items-center gap-2"><span className={`h-2 w-2 shrink-0 rounded-full ${LEVEL_THEME.objective.bar}`} /><span className="truncate text-sm font-medium text-slate-800">{o.title}</span></div>
                  <div className="mt-0.5 truncate text-xs text-slate-400">{uFirst(o.ownerId)} · {inits.length} initiative{inits.length !== 1 ? "s" : ""}</div>
                </button>
                <div className="hidden w-24 shrink-0 sm:block"><div className="mb-0.5 flex justify-between text-[10px] text-slate-400"><span>done</span><span className="font-mono">{om.execution}%</span></div><MiniBar pct={om.execution} tone="bg-amber-500" /></div>
                <div className="hidden w-24 shrink-0 md:block"><SegBar counts={statusCountsOfNodes(inits, works, acts)} total={inits.length} /></div>
                <span className="shrink-0"><StatusTag status={nodeStatus(works, acts, o.id)} /></span>
                {o.deadline && <span className="hidden shrink-0 lg:block"><DueChip date={o.deadline} small /></span>}
              </div>
              {oOpen && (
                <div className="space-y-1 border-t border-slate-100 bg-slate-50/60 p-2">
                  {!inits.length && <div className="px-2 py-1 text-xs text-slate-400">No initiatives in your scope.</div>}
                  {inits.map((it) => {
                    const im = computeMeters(works, acts, it.id); const wks = worksOf(it.id); const iOpen = open.has(it.id);
                    return (
                      <div key={it.id} className="overflow-hidden rounded-md border border-slate-200 bg-white">
                        <div className="flex items-center gap-2 p-2 sm:gap-3">
                          <button onClick={() => wks.length && toggle(it.id)} className="shrink-0 rounded p-0.5 text-slate-400 hover:bg-slate-100"><ChevronRight size={14} className={`transition-transform ${iOpen ? "rotate-90" : ""} ${wks.length ? "" : "opacity-20"}`} /></button>
                          <button onClick={() => onOpen(it.id)} className="min-w-0 flex-1 text-left">
                            <div className="flex items-center gap-2"><span className={`h-1.5 w-1.5 shrink-0 rounded-full ${LEVEL_THEME.initiative.bar}`} /><span className="truncate text-sm text-slate-700">{it.title}</span></div>
                            <div className="mt-0.5 truncate text-xs text-slate-400">{it.result ? `${it.result.metric}: ${it.result.current}/${it.result.target} ${it.result.unit}` : `${wks.length} work item${wks.length !== 1 ? "s" : ""}`}</div>
                          </button>
                          {im.resultPct != null && <div className="hidden w-24 shrink-0 sm:block"><div className="mb-0.5 flex justify-between text-[10px] text-slate-400"><span>to target</span><span className="font-mono text-blue-800">{Math.round(im.resultPct)}%</span></div><MiniBar pct={im.resultPct} tone="bg-blue-700" /></div>}
                          <div className="hidden w-24 shrink-0 md:block"><div className="mb-0.5 flex justify-between text-[10px] text-slate-400"><span>done</span><span className="font-mono">{im.execution}%</span></div><MiniBar pct={im.execution} tone="bg-amber-500" /></div>
                          <span className="shrink-0"><StatusTag status={nodeStatus(works, acts, it.id)} /></span>
                          {it.deadline && <span className="hidden shrink-0 lg:block"><DueChip date={it.deadline} small /></span>}
                        </div>
                        {iOpen && wks.length > 0 && (
                          <div className="space-y-1 border-t border-slate-100 p-2">
                            {wks.map((w) => { const wm = computeMeters(works, acts, w.id); const att = attentionCount(works, acts, w.id);
                              return (
                                <button key={w.id} onClick={() => onOpen(w.id)} className="flex w-full items-center gap-2 rounded-md px-2 py-1.5 text-left hover:bg-slate-50 sm:gap-3">
                                  <span className={`h-1.5 w-1.5 shrink-0 rounded-full ${LEVEL_THEME.work.bar}`} />
                                  <span className="min-w-0 flex-1 truncate text-sm text-slate-600">{w.title}</span>
                                  {att.overdue > 0 && <span className="shrink-0 rounded-full bg-rose-50 px-1.5 py-0.5 text-[10px] font-medium text-rose-700">{att.overdue} overdue</span>}
                                  {att.blocked > 0 && <span className="shrink-0 rounded-full bg-rose-50 px-1.5 py-0.5 text-[10px] font-medium text-rose-700">{att.blocked} blocked</span>}
                                  <span className="hidden w-20 shrink-0 sm:block"><MiniBar pct={wm.execution} tone="bg-teal-500" /></span>
                                  <span className="shrink-0"><StatusTag status={nodeStatus(works, acts, w.id)} /></span>
                                  {w.deadline && <span className="hidden shrink-0 lg:block"><DueChip date={w.deadline} small /></span>}
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

// Per-level column sets: each level shows the metrics that matter for it and
// names the parent it rolls up into, rather than one generic "Metric" column.
const OWNER_CELL = (i, fallback) => i.ownerId ? <><Avatar id={i.ownerId} size={18} /><span className="truncate">{uFirst(i.ownerId)}</span></> : <span className="text-slate-300">{fallback}</span>;
const GAP_BADGE = (i) => i.flags.gap ? <span className="ml-1 shrink-0 rounded bg-amber-50 px-1 text-[10px] font-medium text-amber-700">gap</span> : null;
const TRACKER_COLS = {
  objective: [
    { label: "Owner", w: "hidden w-28 shrink-0 items-center gap-1.5 sm:flex", cell: (i) => OWNER_CELL(i, "—") },
    { label: "Initiatives", w: "hidden w-24 shrink-0 text-right sm:block", cell: (i) => i.childCount },
    { label: "Done", w: "hidden w-24 shrink-0 text-right sm:block", cell: (i) => `${i.execution}%` },
  ],
  initiative: [
    { label: "Objective", w: "hidden w-44 shrink-0 truncate md:block", cell: (i) => i.parentTitle || "—" },
    { label: "Owner", w: "hidden w-24 shrink-0 items-center gap-1.5 sm:flex", cell: (i) => OWNER_CELL(i, "—") },
    { label: "To target", w: "hidden w-32 shrink-0 items-center justify-end gap-1 sm:flex", cell: (i) => i.result ? <span className="truncate">{i.result.current}/{i.result.target} {i.result.unit}{i.resultPct != null && <span className="ml-1 font-medium text-blue-800">{i.resultPct}%</span>}</span> : <span>—</span> },
  ],
  work: [
    { label: "Initiative", w: "hidden w-44 shrink-0 truncate md:block", cell: (i) => i.parentTitle || "—" },
    { label: "Owner", w: "hidden w-24 shrink-0 items-center gap-1.5 sm:flex", cell: (i) => OWNER_CELL(i, "—") },
    { label: "Done", w: "hidden w-24 shrink-0 items-center justify-end gap-1 sm:flex", cell: (i) => <><span>{i.execution}%</span>{GAP_BADGE(i)}</> },
  ],
  activity: [
    { label: "Work", w: "hidden w-40 shrink-0 truncate md:block", cell: (i) => i.parentTitle || "—" },
    { label: "Assignee", w: "hidden w-24 shrink-0 items-center gap-1.5 sm:flex", cell: (i) => OWNER_CELL(i, "Unassigned") },
  ],
};
function Tracker({ items, filter, setFilter, onOpen, title = "Tracker" }) {
  const set = (patch) => setFilter((f) => ({ ...f, ...patch }));
  const levels = ["objective", "initiative", "work", "activity"].filter((l) => items.some((i) => i.level === l));
  const level = levels.includes(filter.level) ? filter.level : levels[0];
  const byLevel = items.filter((i) => i.level === level);
  const sc = statusCountsFromItems(byLevel);
  const gapCount = byLevel.filter((i) => i.flags.gap).length;
  const filtered = byLevel.filter((i) => {
    if (filter.status !== "all" && i.status !== filter.status) return false;
    if (filter.gap && !i.flags.gap) return false;
    if (filter.dueBefore && (!i.due || i.due > filter.dueBefore)) return false;
    if (filter.q) { const hay = `${i.title} ${i.parentTitle || ""} ${uName(i.ownerId)}`.toLowerCase(); if (!hay.includes(filter.q.toLowerCase())) return false; }
    return true;
  });
  const sorted = [...filtered].sort((a, b) => {
    const ar = a.status === "overdue" ? 0 : 1, br = b.status === "overdue" ? 0 : 1; if (ar !== br) return ar - br;
    if (a.due && b.due) return a.due < b.due ? -1 : a.due > b.due ? 1 : 0;
    if (a.due) return -1; if (b.due) return 1;
    return 0;
  });
  const cols = TRACKER_COLS[level];
  // Phones (<sm) hide the middle columns, so fold owner + the level's key metric
  // into a compact meta line under the title.
  const metricText = (i) => i.level === "activity" ? (i.hrs ? `${i.hrs}h` : "—")
    : i.level === "initiative" ? (i.result ? `${i.result.current}/${i.result.target} ${i.result.unit}${i.resultPct != null ? ` · ${i.resultPct}%` : ""}` : `${i.execution}% done`)
    : i.level === "objective" ? `${i.childCount} initiative${i.childCount !== 1 ? "s" : ""} · ${i.execution}% done`
    : `${i.execution}% done`;
  const showGap = level === "initiative" || level === "work";
  const active = filter.status !== "all" || filter.gap || filter.dueBefore || filter.q;
  return (
    <Panel title={title} right={<span className="text-xs text-slate-400">{sorted.length} shown</span>}>
      <div className="mb-3 flex flex-wrap items-center gap-1">
        {levels.map((k) => <button key={k} onClick={() => set({ level: k, status: "all", gap: false })} className={`rounded-md px-2.5 py-1 text-xs font-medium ${level === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>{LEVEL_PLURAL[k]}<span className="ml-1 text-[10px] opacity-70">{items.filter((i) => i.level === k).length}</span></button>)}
      </div>
      <div className="mb-3 flex flex-wrap items-center gap-1.5">
        {STATUS_KEYS.map((k) => <button key={k} onClick={() => set({ status: filter.status === k ? "all" : k, gap: false })} className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${filter.status === k ? `${TRACK_STATUS[k].soft} ${TRACK_STATUS[k].text} ring-1 ring-inset ring-current` : "bg-slate-50 text-slate-500 hover:bg-slate-100"}`}><span className={`h-1.5 w-1.5 rounded-full ${TRACK_STATUS[k].dot}`} />{TRACK_STATUS[k].label}<span className="opacity-60">{sc[k] || 0}</span></button>)}
        {showGap && <button onClick={() => set({ gap: !filter.gap, status: "all" })} className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${filter.gap ? "bg-amber-50 text-amber-700 ring-1 ring-inset ring-current" : "bg-slate-50 text-slate-500 hover:bg-slate-100"}`}><Target size={11} /> Gap to target<span className="opacity-60">{gapCount}</span></button>}
        <div className="ml-auto flex items-center gap-1.5">
          <div className="relative"><Search size={13} className="pointer-events-none absolute left-2 top-1/2 -translate-y-1/2 text-slate-400" /><input value={filter.q} onChange={(e) => set({ q: e.target.value })} placeholder="Search…" className="w-32 rounded-md border border-slate-200 py-1 pl-7 pr-2 text-xs outline-none focus:border-slate-400 sm:w-40" /></div>
          <input type="date" value={filter.dueBefore} onChange={(e) => set({ dueBefore: e.target.value })} title="Show items due on or before this date" className="rounded-md border border-slate-200 px-2 py-1 text-xs text-slate-500" />
          {active && <button onClick={() => setFilter((f) => ({ ...f, status: "all", gap: false, dueBefore: "", q: "" }))} className="text-xs text-slate-400 hover:text-slate-700">clear</button>}
        </div>
      </div>
      <div className="hidden items-center gap-3 border-b border-slate-200 px-2 pb-1.5 text-[10px] font-medium uppercase tracking-wide text-slate-400 sm:flex">
        <span className="min-w-0 flex-1">{LEVEL_LABEL[level]}</span>
        {cols.map((c, idx) => <span key={idx} className={c.w}>{c.label}</span>)}
        <span className="hidden w-32 shrink-0 lg:block">Timeline</span>
        <span className="hidden w-20 shrink-0 md:block">Effort</span>
        <span className="w-24 shrink-0 text-right">Status</span><span className="w-20 shrink-0 text-right">Due</span>
      </div>
      <div className="max-h-[28rem] overflow-y-auto">
        {!sorted.length && <div className="py-8 text-center text-sm text-slate-400">Nothing matches these filters.</div>}
        {sorted.map((i) => (
          <button key={i.id} onClick={() => onOpen(i.openId || i.id)} className="flex w-full items-center gap-3 border-b border-slate-100 px-2 py-2 text-left last:border-0 hover:bg-slate-50">
            <span className="min-w-0 flex-1">
              <span className="block truncate text-sm text-slate-800">{i.title}</span>
              {i.parentTitle && <span className="block truncate text-xs text-slate-400 md:hidden">{i.parentTitle}</span>}
              <span className="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-slate-400 md:hidden">
                {i.ownerId && <span className="inline-flex shrink-0 items-center gap-1"><Avatar id={i.ownerId} size={14} />{uFirst(i.ownerId)}</span>}
                {i.span && <TimelineChip start={i.span.start} end={i.span.end} small />}
                {i.effort && <EffortBadge planned={i.effort.planned} actual={i.effort.actual} small />}
              </span>
            </span>
            {cols.map((c, idx) => <span key={idx} className={`${c.w} text-xs text-slate-500`}>{c.cell(i)}</span>)}
            <span className="hidden w-32 shrink-0 lg:block">{i.span ? <TimelineChip start={i.span.start} end={i.span.end} small /> : null}</span>
            <span className="hidden w-20 shrink-0 md:block">{i.effort ? <EffortBadge planned={i.effort.planned} actual={i.effort.actual} small /> : null}</span>
            <span className="w-24 shrink-0 text-right"><StatusTag status={i.status} /></span>
            <span className="w-20 shrink-0 text-right">{i.due ? <DueChip date={i.due} small /> : <span className="text-xs text-slate-300">—</span>}</span>
          </button>
        ))}
      </div>
    </Panel>
  );
}

function Portfolio({ works, acts, crs, me, isOrg, teams, onOpen, onFocus, goApprovals, view, setView, reviewMode, setReviewMode, onRemark, patchAct, flash }) {
  const [dueDate, setDueDate] = useState(""); // global "as of" date (lives here so it sits next to the view toggle)
  const roots = homeNodes(works, acts, me);
  const rows = roots.map((w) => { const m = computeMeters(works, acts, w.id); const st = workStats(works, acts, w.id); const att = attentionCount(works, acts, w.id); return { w, m, st, rag: nodeRag(works, acts, w.id), issue: deepestIssue(works, acts, w.id), childCount: works.filter((x) => x.parentId === w.id).length, blocked: att.blocked, crc: crs.filter((c) => c.workId === w.id && c.status === "pending").length }; });
  return (
    <div>
      {reviewMode && isOrg && <div className="mb-3 flex items-center justify-between gap-2 rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs text-blue-900"><span className="flex items-center gap-1.5"><Pencil size={13} /> Review &amp; update mode — {view === "tree" ? "click any card in the tree" : "switch to Tree view and click any card"} to leave a remark, re-assign, or shift dates. Owners get nudged.</span><button onClick={() => setReviewMode(false)} className="rounded-md border border-blue-200 bg-white px-2 py-1 font-medium text-blue-800 hover:bg-blue-100">Exit</button></div>}
      {me.level !== "member" && <div className="mb-4 flex flex-wrap items-start justify-between gap-x-4 gap-y-2">
        <div className="min-w-0">
          <div className="text-lg font-medium text-slate-900">{me.level === "md" ? "Portfolio overview" : `Your function — ${me.fn}`}</div>
          {view !== "tree" && <div className="mt-0.5 text-xs text-slate-400">{me.level === "md" ? "Health and progress up top; filter the full tracker below by level, status, owner or due date." : "Your initiatives and their work — health up top, filterable tracker below."}</div>}
        </div>
        <div className="flex shrink-0 items-center gap-2">
          {view !== "tree" && <span className="hidden items-center gap-1.5 text-xs text-slate-400 sm:inline-flex">As of <DatePicker value={dueDate} onChange={setDueDate} placeholder="Any date" /></span>}
          {view !== "tree" && <span className="sm:hidden"><DatePicker value={dueDate} onChange={setDueDate} placeholder="Any date" /></span>}
          <div className="inline-flex items-center gap-0.5 rounded-lg border border-slate-200 bg-white p-0.5">
            <button onClick={() => setView("scorecard")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "tree" ? "text-slate-500 hover:bg-slate-50" : "bg-slate-900 text-white"}`}><LayoutGrid size={13} /> Dashboard</button>
            <button onClick={() => setView("tree")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "tree" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}><ChevronRight size={13} className="rotate-90" /> Tree</button>
          </div>
        </div>
      </div>}
      {view === "tree" && me.level !== "member" ? <BracketTree roots={roots} works={works} acts={acts} onOpen={onOpen} reviewMode={reviewMode} setReviewMode={setReviewMode} onRemark={onRemark} userId={me.id} />
        : me.level === "member" ? <MyWork {...{ me, works, acts, rows, onOpen, onFocus, patchAct, flash }} />
        : <LeaderDashboard {...{ me, works, acts, teams, onFocus, dueDate }} />}
    </div>
  );
}

function LeaderDashboard({ me, works, acts, teams, onFocus, dueDate }) {
  // The whole overview is one interconnected dashboard: a global "as of" date
  // (owned by Portfolio, next to the view toggle) scopes every section, and
  // clicking a donut slice cross-filters the rest.
  const [crossFilter, setCrossFilter] = useState(null); // { level, status } | null
  const sacts = actsUpTo(acts, dueDate); // date-scoped activities feed everything
  const items = trackerRows(works, sacts, me);
  const toggleCross = (level, status) => setCrossFilter((c) => (c && c.level === level && c.status === status ? null : { level, status }));

  // Cross-filter selection: the set of activities behind the clicked slice, and
  // the node ids on the path to them (so the other donuts show only what's related).
  let selectedActs = null, scopeNodeIds = null, teamFilter = null, userFilter = null;
  if (crossFilter) {
    if (crossFilter.level === "activity") selectedActs = sacts.filter((a) => a.status !== "cancelled" && activityStatus(a) === crossFilter.status);
    else { const ids = new Set(works.filter((w) => w.level === crossFilter.level && nodeStatus(works, sacts, w.id) === crossFilter.status).flatMap((n) => subtreeIds(works, n.id))); selectedActs = sacts.filter((a) => a.status !== "cancelled" && ids.has(a.workId)); }
    scopeNodeIds = new Set();
    selectedActs.forEach((a) => { let cur = works.find((w) => w.id === a.workId), g = 0; while (cur && g++ < 12) { scopeNodeIds.add(cur.id); cur = works.find((w) => w.id === cur.parentId); } });
    userFilter = new Set(selectedActs.map((a) => a.assigneeId).filter(Boolean));
    teamFilter = new Set((teams || []).filter((t) => t.memberIds.some((id) => userFilter.has(id))).map((t) => t.id));
  }

  return (
    <div className="space-y-5">
      {crossFilter && <div><button onClick={() => setCrossFilter(null)} className="inline-flex items-center gap-1.5 rounded-full bg-blue-50 px-2.5 py-1 text-xs font-medium text-blue-800 hover:bg-blue-100">Filtered: {TRACK_STATUS[crossFilter.status].label} {LEVEL_PLURAL[crossFilter.level].toLowerCase()}<X size={12} /></button></div>}
      <HealthStrip items={items} onFocus={onFocus} />
      <LevelHealthChart items={items} crossFilter={crossFilter} scopeNodeIds={scopeNodeIds} onSelect={toggleCross} />
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
        <TeamPerformance teams={teams} works={works} acts={sacts} me={me} teamFilter={teamFilter} />
        <Leaderboard works={works} acts={sacts} teams={teams} me={me} userFilter={userFilter} />
      </div>
    </div>
  );
}

// The Data tab: the Progress cascade + filterable Tracker, moved off the
// Portfolio overview. Drill-ins from the dashboard arrive via `focus`.
function DataPage({ me, works, acts, onOpen, focus, patchAct, flash }) {
  const items = trackerRows(works, acts, me);
  const hasProgress = me.level !== "member"; // members have no objective cascade
  const [view, setView] = useState("tracker");
  const [filter, setFilter] = useState({ level: me.level === "member" ? "activity" : "objective", status: "all", gap: false, dueBefore: "", q: "" });
  useEffect(() => {
    if (focus && focus.patch) { setView("tracker"); setFilter((f) => ({ level: f.level, status: "all", gap: false, dueBefore: "", q: "", ...focus.patch })); }
  }, [focus]);
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <div className="text-lg font-medium text-slate-900">Data</div>
          <div className="mt-0.5 text-xs text-slate-400">Filter every objective, initiative, work and activity by level, status, owner or due date{hasProgress ? " — or read it top-down as progress." : "."}</div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <OutlookSyncButton me={me} acts={acts} works={works} patchAct={patchAct} flash={flash} />
          {hasProgress && <div className="inline-flex items-center gap-0.5 rounded-lg border border-slate-200 bg-white p-0.5">
            <button onClick={() => setView("tracker")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "tracker" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}><ClipboardList size={13} /> Tracker</button>
            <button onClick={() => setView("progress")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "progress" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}><Gauge size={13} /> Progress</button>
          </div>}
        </div>
      </div>
      {hasProgress && view === "progress"
        ? <ObjectiveFlow works={works} acts={acts} me={me} onOpen={onOpen} />
        : <Tracker items={items} filter={filter} setFilter={setFilter} onOpen={onOpen} title={me.level === "member" ? "My items — tracker" : "Tracker"} />}
    </div>
  );
}

function MyWork({ me, works, acts, rows, onOpen, onFocus, patchAct, flash }) {
  const mine = acts.filter((a) => a.assigneeId === me.id && a.status !== "cancelled");
  const done = mine.filter((a) => a.status === "executed").length; const overdue = mine.filter((a) => isOverdue(a)).length; const blocked = mine.filter((a) => a.blocked && a.status !== "executed").length;
  const myDelivs = works.filter((w) => w.ownerId === me.id).flatMap((w) => w.deliverables || []).filter((d) => d.done && typeof d.score === "number");
  const avg = myDelivs.length ? Math.round(myDelivs.reduce((s, d) => s + d.score, 0) / myDelivs.length) : "—";
  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <div className="text-lg font-medium text-slate-900">My work</div>
          <div className="mt-0.5 text-xs text-slate-400">Everything you own or are assigned — {me.title}. Open the Data tab for the full, filterable tracker.</div>
        </div>
        <OutlookSyncButton me={me} acts={acts} works={works} patchAct={patchAct} flash={flash} />
      </div>
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        <KpiCard label="Open activities" value={mine.length - done} />
        <KpiCard label="Done" value={done} tone="text-emerald-600" />
        <KpiCard label="Overdue" value={overdue} tone={overdue ? "text-rose-600" : "text-slate-900"} sub={overdue ? "view →" : undefined} onClick={() => onFocus({ level: "activity", status: "overdue" })} />
        <KpiCard label={blocked ? "Blocked" : "Avg deliverable"} value={blocked ? blocked : (typeof avg === "number" ? avg + "/100" : avg)} tone={blocked ? "text-rose-600" : "text-blue-800"} sub={blocked ? "view →" : undefined} onClick={blocked ? () => onFocus({ level: "activity", status: "blocked" }) : undefined} />
      </div>
      <div>
        <div className="mb-2 text-sm font-medium text-slate-700">Work I'm on</div>
        <div className="grid grid-cols-1 gap-3 lg:grid-cols-2">
          {rows.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-400">Nothing assigned to you yet.</div>}
          {rows.map((row) => { const { w, m } = row; const ct = LEVEL_THEME[w.level]; const myActs = acts.filter((a) => subtreeIds(works, w.id).includes(a.workId) && a.assigneeId === me.id && a.status !== "cancelled"); const myDone = myActs.filter((a) => a.status === "executed").length; const wspan = nodeSpan(works, acts, w); const weff = effortRollup(works, acts, w.id);
            return (
              <div key={w.id} className={`overflow-hidden rounded-xl border bg-white ${row.rag === "red" ? "border-rose-200" : "border-slate-200"}`}>
                <div className={`h-1 ${ct.bar}`} />
                <button onClick={() => onOpen(w.id)} className="block w-full p-4 text-left hover:bg-slate-50">
                  <div className="flex items-start justify-between gap-2"><LevelChip level={w.level} /><StatusPill rag={row.rag} /></div>
                  <div className="mt-1.5 truncate text-sm font-medium text-slate-800">{w.title}</div>
                  <div className="mt-0.5 text-xs text-slate-500">My part: {myDone}/{myActs.length} activities done</div>
                  <div className="mt-1.5 flex flex-wrap items-center gap-x-2 gap-y-1"><TimelineChip start={wspan.start} end={wspan.end} /><EffortBadge planned={weff.planned} actual={weff.actual} small /></div>
                  <div className="mt-3"><ProgressPair planning={m.planning} execution={m.execution} /></div>
                </button>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ---------- Work detail ---------- */
function BracketTree({ roots, works, acts, onOpen, reviewMode, setReviewMode, onRemark, userId }) {
  const CARDW = 216, CARDH = 76, ROWGAP = 140, COLGAP = 248;
  const rootIds = roots.map((r) => r.id);
  const childIds = (id) => works.filter((w) => w.parentId === id).map((w) => w.id); // tree stops at sub-work
  const allWorkIds = [...new Set(rootIds.flatMap((id) => subtreeIds(works, id).filter((x) => works.find((w) => w.id === x))))];
  const init = () => { const s = new Set(rootIds); roots.forEach((r) => { const iss = deepestIssue(works, acts, r.id); if (iss) { let cur = works.find((w) => w.id === iss.workId); let g = 0; while (cur && g++ < 12) { s.add(cur.id); if (cur.id === r.id) break; cur = works.find((w) => w.id === cur.parentId); } } }); return s; };
  const [expanded, setExpanded] = useState(init);
  const toggle = (id) => setExpanded((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  // ----- layout pass: each visible node gets a row (depth) and an x-slot -----
  const nodes = {}; let leaf = 0; let maxDepth = 0;
  const place = (id, depth) => {
    maxDepth = Math.max(maxDepth, depth);
    const kids = expanded.has(id) ? childIds(id) : [];
    let x;
    if (kids.length === 0) { x = leaf * COLGAP; leaf++; }
    else { const xs = kids.map((k) => place(k, depth + 1)); x = (xs[0] + xs[xs.length - 1]) / 2; }
    nodes[id] = { depth, x };
    return x;
  };
  rootIds.forEach((r) => place(r, 0));
  const posFor = (id) => ({ x: nodes[id].x, y: nodes[id].depth * ROWGAP });
  const edges = [];
  Object.keys(nodes).forEach((id) => { if (!expanded.has(id)) return; childIds(id).forEach((cid) => { if (nodes[cid]) edges.push([id, cid]); }); });
  const W = Math.max(leaf, 1) * COLGAP + CARDW + 24;
  const H = maxDepth * ROWGAP + CARDH + 24;

  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <LevelLegend />
        <div className="flex flex-wrap gap-1">
          <button onClick={() => setExpanded(new Set(allWorkIds))} className={btnLight}>Expand all</button>
          <button onClick={() => setExpanded(new Set(rootIds))} className={btnLight}>Collapse</button>
        </div>
      </div>
      {reviewMode && <div className="mb-3 flex items-center justify-between gap-2 rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-xs text-blue-900"><span className="flex items-center gap-1.5"><Pencil size={13} /> Review mode — click any initiative or work to leave a remark &amp; nudge its owners.</span><button onClick={() => setReviewMode(false)} className="rounded-md border border-blue-200 bg-white px-2 py-1 font-medium text-blue-800 hover:bg-blue-100">Exit</button></div>}
      <div className="overflow-auto rounded-xl border border-slate-200 bg-slate-50 p-4">
        <div className="relative" style={{ width: W, height: H }}>
          <svg className="absolute inset-0" width={W} height={H}>
            {edges.map(([p, c], i) => { const pn = posFor(p), cn = posFor(c); const px = pn.x + CARDW / 2, py = pn.y + CARDH, cx = cn.x + CARDW / 2, cy = cn.y; const my = py + (cy - py) / 2; const red = nodeRag(works, acts, c) === "red"; return <path key={i} d={`M ${px} ${py} C ${px} ${my}, ${cx} ${my}, ${cx} ${cy}`} fill="none" stroke={red ? "#fb7185" : "#cbd5e1"} strokeWidth={red ? 2 : 1.5} />; })}
          </svg>
          {Object.keys(nodes).map((id) => {
            const { x: left, y: top } = posFor(id);
            const w = works.find((x) => x.id === id); const th = LEVEL_THEME[w.level]; const rag = nodeRag(works, acts, id); const catt = attentionCount(works, acts, id); const m = computeMeters(works, acts, id); const hasKids = childIds(id).length > 0; const open = expanded.has(id); const canRemark = reviewMode && w.level !== "objective"; const eff = effortRollup(works, acts, id);
            return (
              <div
                key={id}
                style={{ position: "absolute", left, top, width: CARDW, height: CARDH }}
                onClick={() => { if (canRemark) onRemark(w); else if (hasKids) toggle(id); }}
                className={`relative overflow-hidden rounded-lg border bg-white ${rag === "red" ? "border-rose-200 ring-1 ring-rose-100" : "border-slate-200"} ${canRemark ? "cursor-pointer ring-1 ring-blue-200 hover:ring-blue-400" : hasKids ? "cursor-pointer hover:shadow-sm" : ""}`}
              >
                <span className={`absolute inset-y-0 left-0 w-1 ${th.bar}`} />
                <div className="flex h-full flex-col justify-center gap-1 pl-3 pr-2">
                  <div className="flex items-center gap-1.5">
                    <span className={`h-2 w-2 shrink-0 rounded-full ${th.bar}`} />
                    <span className="min-w-0 flex-1 truncate text-xs font-medium text-slate-800">{w.title}</span>
                    {w.recurring && <span className="shrink-0 rounded bg-blue-100 px-1 text-xs font-medium text-blue-800" title="recurring">↻ {w.recurring.cadence}</span>}
                    {hasKids && <span onClick={(e) => { e.stopPropagation(); toggle(id); }} className="shrink-0 cursor-pointer rounded p-0.5 hover:bg-slate-100"><ChevronRight size={13} className={`text-slate-400 transition-transform ${open ? "rotate-90" : ""}`} /></span>}
                  </div>
                  <div className="flex items-center gap-1 text-xs text-slate-500"><Avatar id={w.ownerId} size={14} /><span className="min-w-0 truncate">{uFirst(w.ownerId)}</span>{w.result && <span className={`shrink-0 font-mono ${th.text}`}>· {m.resultPct != null ? Math.round(m.resultPct) + "%" : "—"}</span>}{w.deadline && <span className="ml-auto shrink-0"><DueChip date={w.deadline} small /></span>}</div>
                  <div className="flex items-center gap-1.5 text-xs">
                    <span className="text-teal-600">Plan {m.planning}%</span>
                    <span className="text-amber-600">Done {m.execution}%</span>
                    {(eff.planned > 0 || eff.actual > 0) && <span className="font-mono text-slate-400" title="actual / estimated effort (hours)">{eff.actual}/{eff.planned}h</span>}
                    {catt.blocked > 0 && <span className="text-rose-600">· {catt.blocked} blk</span>}
                    <span className={`ml-auto h-2 w-2 shrink-0 rounded-full ${RAG[rag][0]}`} title={RAG[rag][3]} />
                    <button onClick={(e) => { e.stopPropagation(); onOpen(id); }} className="shrink-0 rounded p-0.5 text-slate-300 hover:text-slate-700" title="Open detail"><ArrowRight size={12} /></button>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
      <div className="mt-2 px-1 text-xs text-slate-400">Shows Objective → Initiative → Work → Activity, top to bottom. Click a card to branch it out · ↻ = recurring work · the arrow opens full detail · scroll for deep or wide branches.</div>
    </div>
  );
}

// Kanban board of activities — drag a card between columns to change its status.
// Only the assignee (owner) or an org leader (exec) can move a card.
function ActivitiesKanban({ acts, canMove, onMove, onOpen }) {
  const [dragId, setDragId] = useState(null);
  const cols = [
    { key: "todo", label: "To do", cls: "bg-slate-400", match: (a) => a.status !== "executed" && !a.inProgress && !a.onHold },
    { key: "inprogress", label: "In progress", cls: "bg-amber-500", match: (a) => a.status !== "executed" && a.inProgress && !a.onHold },
    { key: "hold", label: "On hold", cls: "bg-slate-400", match: (a) => a.status !== "executed" && a.onHold },
    { key: "done", label: "Done", cls: "bg-emerald-500", match: (a) => a.status === "executed" },
  ];
  const patchFor = (a, key) => key === "todo" ? { status: "planned", inProgress: false, onHold: false }
    : key === "inprogress" ? { status: "planned", inProgress: true, onHold: false }
    : key === "hold" ? { status: "planned", inProgress: false, onHold: true }
    : { status: "executed", inProgress: false, onHold: false, actualHrs: a.actualHrs != null ? a.actualHrs : a.plannedHrs };
  const drop = (key) => { const a = acts.find((x) => x.id === dragId); setDragId(null); if (a && canMove(a) && !cols.find((c) => c.key === key).match(a)) onMove(a.id, patchFor(a, key)); };
  return (
    <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-4">
      {cols.map((col) => { const items = acts.filter(col.match);
        return (
          <div key={col.key} onDragOver={(e) => e.preventDefault()} onDrop={() => drop(col.key)} className="rounded-lg border border-slate-200 bg-slate-50 p-2">
            <div className="mb-2 flex items-center gap-1.5 px-1 text-xs font-medium text-slate-500"><span className={`h-2 w-2 rounded-full ${col.cls}`} /> {col.label} <span className="ml-auto rounded-full bg-white px-1.5 text-slate-400">{items.length}</span></div>
            <div className="min-h-[3rem] space-y-2">
              {items.map((a) => { const mine = canMove(a); const over = isOverdue(a); const Icon = ACT_ICON[a.actType] || User;
                return (
                  <div key={a.id} draggable={mine} onDragStart={() => setDragId(a.id)} onDragEnd={() => setDragId(null)} className={`rounded-md border bg-white p-2 ${a.blocked ? "border-rose-300" : over ? "border-rose-200" : "border-slate-200"} ${mine ? "cursor-grab active:cursor-grabbing hover:shadow-sm" : ""}`}>
                    <div className="flex items-start gap-1.5"><Icon size={13} className="mt-0.5 shrink-0 text-slate-400" /><span className="min-w-0 flex-1 text-sm text-slate-800">{a.title}</span></div>
                    <div className="mt-1 flex flex-wrap items-center gap-1.5 pl-5 text-xs text-slate-400">
                      {a.assigneeId ? <span className="inline-flex items-center gap-1"><Avatar id={a.assigneeId} size={13} /> {uFirst(a.assigneeId)}</span> : <span className="text-amber-700">Unassigned</span>}
                      <span className={over ? "text-rose-600" : ""}>· {a.date ? fmtFull(parseISO(a.date)) : "no date"}</span>
                      <span>· {a.plannedHrs}h</span>
                    </div>
                  </div>
                );
              })}
              {items.length === 0 && <div className="rounded-md border border-dashed border-slate-200 py-3 text-center text-xs text-slate-300">—</div>}
            </div>
          </div>
        );
      })}
    </div>
  );
}
// Compact result-vs-target strip, shown inside the header card (not a big box).
function ResultInline({ work, m, isOrg, patchWork }) {
  const [edit, setEdit] = useState(false); const [val, setVal] = useState(work.result.current); const rt = work.result;
  return (
    <div className="mt-4 rounded-lg border border-slate-100 bg-slate-50 p-3">
      <div className="flex flex-wrap items-center justify-between gap-x-3 gap-y-1">
        <div className="min-w-0 text-sm text-slate-700"><span className="text-xs font-medium uppercase tracking-wide text-slate-400">Result</span> <span className="ml-1">{rt.metric}: <span className="font-mono font-medium text-slate-900">{rt.current}</span> <span className="text-xs text-slate-400">of {rt.target} {rt.unit}</span></span></div>
        <div className="flex items-center gap-2">{m.resultPct != null && <span className="font-mono text-sm font-medium text-blue-800">{Math.round(m.resultPct)}%</span>}{isOrg && <button onClick={() => setEdit(!edit)} className="text-xs font-medium text-blue-700">{edit ? "close" : "update"}</button>}</div>
      </div>
      {m.resultPct != null && <div className="mt-2 h-2 overflow-hidden rounded-full bg-blue-100"><div className="h-full rounded-full bg-blue-700" style={{ width: `${m.resultPct}%` }} /></div>}
      {edit && <div className="mt-2 flex items-center gap-2"><input type="number" value={val} onChange={(e) => setVal(e.target.value)} className="w-24 rounded-md border border-slate-200 px-2 py-1 text-sm" /><button onClick={() => { patchWork(work.id, { result: { ...rt, current: Number(val) } }); setEdit(false); }} className={btnDark}>Save</button></div>}
    </div>
  );
}
// A single "AI actions" dropdown holding the level-appropriate tools. Each item
// runs its handler (which opens a popup for AI results, or a modal / inline op).
function AiActionsMenu({ items, busy }) {
  const [open, setOpen] = useState(false);
  if (!items.length) return null;
  return (
    <div className="relative">
      <button onClick={() => setOpen((o) => !o)} disabled={!!busy} className="inline-flex items-center gap-1.5 rounded-md border border-indigo-200 bg-indigo-50 px-2.5 py-2 text-xs font-medium text-indigo-700 hover:bg-indigo-100 disabled:opacity-50">
        {busy ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />} AI actions <ChevronRight size={12} className={`transition-transform ${open ? "rotate-90" : ""}`} />
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute right-0 z-20 mt-1 w-56 rounded-lg border border-slate-200 bg-white p-1 shadow-lg">
            {items.map((it) => { const Icon = it.icon; return (
              <button key={it.key} onClick={() => { setOpen(false); it.onClick(); }} disabled={!!it.disabled} className="flex w-full items-center gap-2 rounded-md px-2.5 py-2 text-left text-sm text-slate-700 hover:bg-slate-50 disabled:opacity-50">
                <Icon size={14} className="shrink-0 text-indigo-600" /> <span className="flex-1">{it.label}</span>
              </button>
            ); })}
          </div>
        </>
      )}
    </div>
  );
}
function NodeView({ nodeId, user, works, acts, crs, teams, isOrg, busy, setBusy, flash, patchAct, patchWork, store, onOpen, onRemark, goApprovals }) {
  const node = works.find((w) => w.id === nodeId);
  const [insight, setInsight] = useState(null);
  const [modify, setModify] = useState(false); const [addUn, setAddUn] = useState(false); const [edit, setEdit] = useState(null); const [confirmDel, setConfirmDel] = useState(null); const [addChild, setAddChild] = useState(false); const [suggest, setSuggest] = useState(null); const [blockerOpen, setBlockerOpen] = useState(false); const [approvalsOpen, setApprovalsOpen] = useState(false); const [childView, setChildView] = useState(() => (typeof window !== "undefined" && window.innerWidth < 768 ? "cards" : "list"));
  if (!node) return <button onClick={() => onOpen(null)} className="text-sm text-slate-500">← Portfolio</button>;
  const childLevel = CHILD_LEVEL[node.level]; const childLabel = childLevel === "activity" ? "task" : (LEVEL_LABEL[childLevel] || "").toLowerCase();
  const team = node.level === "initiative" && node.teamId ? (teams || []).find((t) => t.id === node.teamId) : null;
  const teamLead = team ? (team.memberIds.includes(node.ownerId) ? node.ownerId : team.memberIds[0]) : null;
  const canCreate = childLevel && (isOrg || node.ownerId === user.id || (childLevel === "activity" && user.level === "member"));
  const canRemark = isOrg || node.ownerId === user.id;
  const m = computeMeters(works, acts, node.id); const st = workStats(works, acts, node.id);
  const rag = nodeRag(works, acts, node.id); const rg = RAG[rag]; const th = LEVEL_THEME[node.level] || LEVEL_THEME.activity; const att = attentionCount(works, acts, node.id); const issue = deepestIssue(works, acts, node.id);
  const children = works.filter((w) => w.parentId === node.id);
  const leaf = children.length === 0;
  const nodeActs = acts.filter((a) => a.workId === node.id);
  const subs = children.length ? children : [node];
  const path = breadcrumbPath(works, user, node.id);
  const myCRs = crs.filter((c) => c.workId === node.id && c.status === "pending");
  const delAct = (id) => { store.deleteAct(id); setConfirmDel(null); flash("Activity deleted."); };
  const planText = () => subs.map((s) => `- ${s.title}: ${acts.filter((a) => a.workId === s.id).map((a) => a.title + " (" + a.actType + ")").join(", ") || "none"}`).join("\n");
  const expand = async () => {
    const ids = subtreeIds(works, node.id); const leaves = works.filter((w) => ids.includes(w.id) && !works.some((x) => x.parentId === w.id));
    const empty = leaves.filter((l) => !acts.some((a) => a.workId === l.id)); if (!empty.length) return flash("Every branch already has activities.");
    setBusy("expand");
    try { const raw = await aiComplete('For each work, 1-3 activities. Return ONLY JSON: {"fills":[{"title":<given>,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Parent "${node.title}". Empty works: ${JSON.stringify(empty.map((e) => e.title))}`); const fills = parseJSON(raw).fills || []; const add = []; fills.forEach((f) => { const lf = empty.find((l) => l.title === f.title) || empty[0]; (f.activities || []).forEach((ac) => add.push({ id: nid("a"), workId: lf.id, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }); if (add.length) store.addActs(add); flash("AI filled the missing branches."); } catch { flash("AI unavailable."); }
    setBusy(null);
  };
  const doInsight = async () => { setBusy("insight"); setInsight(null); try { setInsight(await AI.insight(node.title, m)); } catch { setInsight({ read: `${Math.round(m.completion)}% complete.`, action: m.stuck ? `Unblock "${m.stuck}".` : "Schedule remaining." }); } setBusy(null); };
  const doSuggest = async () => { setBusy("suggest"); setSuggest(null); try { const existing = works.filter((w) => w.parentId === node.id).map((w) => w.title); const raw = await aiComplete('Propose 3-4 initiatives that would fulfil this objective. Distinct from existing ones. Return ONLY JSON: {"initiatives":[{"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general"}]}', `Objective: "${node.title}". Existing initiatives: ${JSON.stringify(existing)}.`); setSuggest(parseJSON(raw).initiatives || []); } catch { flash("AI unavailable."); } setBusy(null); };
  const addSuggested = (s) => { const tpl = METRIC_BY_TYPE[s.type] || METRIC_BY_TYPE.general; const id = nid("w"); store.addWorks([{ id, parentId: node.id, level: "initiative", title: s.title, type: s.type || "general", ownerId: user.id, result: { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 } }]); setSuggest((prev) => prev.filter((x) => x !== s)); flash("Initiative added — open it to plan the work."); };
  const autoAssign = () => {
    const ids = subtreeIds(works, node.id); const un = acts.filter((a) => ids.includes(a.workId) && !a.assigneeId && a.status !== "cancelled");
    if (!un.length) return flash("Everyone's already assigned.");
    const members = USERS.filter((u) => u.level === "member"); const load = {}; USERS.forEach((u) => (load[u.id] = acts.filter((a) => a.assigneeId === u.id && a.status !== "executed").reduce((s, a) => s + a.plannedHrs, 0)));
    const up = {}; let i = 0; un.forEach((a) => { members.sort((x, y) => load[x.id] - load[y.id]); const pick = members[0]; load[pick.id] += a.plannedHrs; up[a.id] = { assigneeId: pick.id, date: iso(addDays(TODAY, i % 5)) }; i++; });
    store.patchActs(up); flash(`Auto-assigned ${un.length} activities, balanced by load.`);
  };

  const closeResults = () => { setInsight(null); setSuggest(null); };
  const aiItems = node.level === "objective" ? [
    { key: "suggest", label: "Suggest initiatives", icon: Wand2, onClick: doSuggest },
    { key: "insight", label: "Where does this stand", icon: Gauge, onClick: doInsight },
  ] : [
    { key: "auto", label: "Auto-assign", icon: Wand2, onClick: autoAssign },
    { key: "fill", label: "Fill missing activities", icon: Plus, onClick: expand },
    { key: "insight", label: "Where do I stand", icon: Gauge, onClick: doInsight },
    { key: "modify", label: "Modify plan", icon: Pencil, onClick: () => setModify(true) },
    { key: "unplanned", label: "Add unplanned activity", icon: Plus, onClick: () => setAddUn(true) },
  ];
  const blockerCount = att.blocked + att.overdue;
  const span = nodeSpan(works, acts, node); const eff = effortRollup(works, acts, node.id);
  const headerBlock = (
    <div className="p-5">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex items-center gap-2"><LevelChip level={node.level} />{node.recurring && <Chip tone="violet">↻ {node.recurring.cadence}</Chip>}{node.scope && <Chip tone={node.scope === "group" ? "blue" : "slate"}>{node.scope === "group" ? "group" : "individual"}</Chip>}</div>
          <h2 className="mt-1.5 text-lg font-medium text-slate-900">{node.title}</h2>
          <div className="mt-1 flex flex-wrap items-center gap-x-2.5 gap-y-1 text-xs text-slate-400">
            {team ? <><Chip tone="blue">{team.name}</Chip><span>lead {uName(teamLead)}</span></> : <><span className="inline-flex items-center gap-1"><Avatar id={node.ownerId} size={18} /> Owner {uName(node.ownerId)}</span></>}
            <TimelineChip start={span.start} end={span.end} />
            <EffortBadge planned={eff.planned} actual={eff.actual} />
            {st.nextDue && <span className="inline-flex items-center gap-1"><CalendarClock size={12} /> next {fmtFull(parseISO(st.nextDue))}</span>}
          </div>
        </div>
        <div className="flex shrink-0 flex-wrap items-center gap-1.5">
          {blockerCount > 0 && issue && <button onClick={() => setBlockerOpen(true)} title="Attention needed — click for the root cause" className="inline-flex items-center gap-1 rounded-md border border-rose-200 bg-rose-50 px-2 py-2 text-xs font-medium text-rose-700 hover:bg-rose-100"><AlertTriangle size={13} /> {att.blocked > 0 ? `${att.blocked} blocked` : ""}{att.blocked > 0 && att.overdue > 0 ? " · " : ""}{att.overdue > 0 ? `${att.overdue} overdue` : ""}</button>}
          {isOrg && myCRs.length > 0 && <button onClick={() => setApprovalsOpen(true)} title="Plan changes awaiting your approval" className="inline-flex items-center gap-1 rounded-md border border-amber-200 bg-amber-50 px-2 py-2 text-xs font-medium text-amber-800 hover:bg-amber-100"><ClipboardCheck size={13} /> {myCRs.length} approval{myCRs.length > 1 ? "s" : ""}</button>}
          {isOrg && <AiActionsMenu items={aiItems} busy={busy} />}
          {canRemark && <button onClick={() => onRemark(node)} title="Remark & update" className="inline-flex items-center gap-1 rounded-md border border-blue-200 bg-white px-2 py-2 text-xs font-medium text-blue-800 hover:bg-blue-50"><Pencil size={12} /></button>}
          <StatusPill rag={rag} />
        </div>
      </div>
      <div className="mt-4 max-w-md"><ProgressPair planning={m.planning} execution={m.execution} size="lg" /></div>
      <div className="mt-3 text-xs text-slate-500">{!leaf ? `${children.length} ${CHILD_LABEL[node.level]} inside · ` : ""}{st.done}/{st.total} activities done{st.overdue > 0 ? ` · ${st.overdue} overdue` : ""}</div>
      {node.result && <ResultInline work={node} m={m} isOrg={isOrg} patchWork={patchWork} />}
    </div>
  );
  const insideDivider = (
    <div className="mb-3 mt-2 flex flex-wrap items-center gap-2">
      <span className={`h-2.5 w-2.5 rounded-full ${leaf ? LEVEL_THEME.activity.bar : LEVEL_THEME[CHILD_LEVEL[node.level]].bar}`} />
      <span className="text-sm font-medium text-slate-700">Inside — {leaf ? `${nodeActs.length} ${nodeActs.length === 1 ? "activity" : "activities"}` : `${children.length} ${CHILD_LABEL[node.level]}`}</span>
      <span className="hidden text-xs text-slate-400 sm:inline">{leaf ? "the actual tasks to execute" : "click any to drill in one level"}</span>
      <div className="h-px flex-1 bg-slate-200" />
      {!leaf ? <div className="inline-flex items-center gap-0.5 rounded-lg border border-slate-200 bg-white p-0.5">
        <button onClick={() => setChildView("cards")} className={`rounded-md px-2 py-1 ${childView === "cards" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`} title="Cards"><LayoutGrid size={13} /></button>
        <button onClick={() => setChildView("list")} className={`rounded-md px-2 py-1 ${childView === "list" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`} title="List / table"><ClipboardList size={13} /></button>
      </div> : <div className="inline-flex items-center gap-0.5 rounded-lg border border-slate-200 bg-white p-0.5">
        <button onClick={() => setChildView("list")} className={`rounded-md px-2 py-1 ${childView !== "kanban" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`} title="List"><ClipboardList size={13} /></button>
        <button onClick={() => setChildView("kanban")} className={`rounded-md px-2 py-1 ${childView === "kanban" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`} title="Kanban board"><LayoutGrid size={13} /></button>
      </div>}
      {leaf && <OutlookSyncButton me={user} acts={acts} works={works} patchAct={patchAct} flash={flash} scopeActs={nodeActs} label="Sync to Outlook" />}
      {canCreate && <button onClick={() => setAddChild(true)} className={btnLight}><Plus size={14} /> Add {childLabel}</button>}
    </div>
  );
  const childrenGrid = (
    <div className="grid grid-cols-1 gap-3 lg:grid-cols-2">
      {children.map((c) => {
        const ct = LEVEL_THEME[c.level]; const cm = computeMeters(works, acts, c.id); const cst = workStats(works, acts, c.id); const crag = nodeRag(works, acts, c.id); const cissue = deepestIssue(works, acts, c.id); const grand = works.filter((x) => x.parentId === c.id).length; const cspan = nodeSpan(works, acts, c); const ceff = effortRollup(works, acts, c.id);
        return (
          <div key={c.id} className={`overflow-hidden rounded-xl border bg-white ${crag === "red" ? "border-rose-200" : "border-slate-200"}`}>
            <div className={`h-1 ${ct.bar}`} />
            <div className="p-4">
              <div className="flex items-start justify-between gap-2">
                <button onClick={() => onOpen(c.id)} className="min-w-0 text-left">
                  <LevelChip level={c.level} />
                  <div className="mt-1 truncate text-sm font-medium text-slate-800 hover:underline">{c.title}</div>
                  <div className="mt-0.5 flex items-center gap-1 text-xs text-slate-400"><Avatar id={c.ownerId} size={14} /> {uFirst(c.ownerId)}{c.result && <span> · {cm.resultPct != null ? Math.round(cm.resultPct) + "%" : "—"} result</span>}</div>
                </button>
                <StatusPill rag={crag} />
              </div>
              <div className="mt-3"><ProgressPair planning={cm.planning} execution={cm.execution} /></div>
              <div className="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1"><TimelineChip start={cspan.start} end={cspan.end} /><EffortBadge planned={ceff.planned} actual={ceff.actual} small /></div>
              <div className="mt-2 flex items-center justify-between text-xs text-slate-500">
                <span>{grand || cst.total} {grand ? CHILD_LABEL[c.level] : "activities"} · {cst.done}/{cst.total} done{cst.overdue > 0 ? ` · ${cst.overdue} overdue` : ""}</span>
                <button onClick={() => onOpen(c.id)} className="inline-flex items-center gap-1 font-medium text-slate-500 hover:text-slate-800">Open <ArrowRight size={12} /></button>
              </div>
              {cissue && <button onClick={() => onOpen(cissue.workId)} className="mt-2 flex w-full items-center gap-1.5 rounded-md border border-rose-200 bg-rose-50 px-2.5 py-1.5 text-left text-xs text-rose-800 hover:bg-rose-100"><AlertTriangle size={12} className="shrink-0 text-rose-500" /><span className="min-w-0 flex-1 truncate">{cissue.blocked ? "Blocked" : "Overdue"}: {cissue.title}</span><ArrowRight size={11} className="shrink-0" /></button>}
            </div>
          </div>
        );
      })}
    </div>
  );
  const childrenList = (
    <div className="overflow-hidden rounded-lg border border-slate-200">
      {children.map((c, i) => { const ct = LEVEL_THEME[c.level]; const cm = computeMeters(works, acts, c.id); const cst = workStats(works, acts, c.id); const crag = nodeRag(works, acts, c.id); const cspan = nodeSpan(works, acts, c); const ceff = effortRollup(works, acts, c.id);
        return (
          <div key={c.id} className={`flex items-center gap-2 px-3 py-2.5 ${i ? "border-t border-slate-100" : ""} hover:bg-slate-50`}>
            <span className={`h-6 w-1 shrink-0 rounded ${ct.bar}`} />
            <span className="shrink-0"><LevelChip level={c.level} /></span>
            <button onClick={() => onOpen(c.id)} className="min-w-0 flex-1 truncate text-left text-sm font-medium text-slate-800 hover:underline">{c.title}</button>
            <span className="hidden shrink-0 lg:inline"><TimelineChip start={cspan.start} end={cspan.end} small /></span>
            <span className="hidden shrink-0 md:inline"><EffortBadge planned={ceff.planned} actual={ceff.actual} small /></span>
            {c.result && <span className={`hidden shrink-0 font-mono text-xs sm:inline ${ct.text}`}>{cm.resultPct != null ? Math.round(cm.resultPct) + "%" : "—"}</span>}
            <span className="hidden shrink-0 text-xs text-teal-600 md:inline">Plan {cm.planning}%</span>
            <span className="hidden shrink-0 text-xs text-amber-600 md:inline">Done {cm.execution}%</span>
            {c.deadline && <span className="hidden shrink-0 sm:inline"><DueChip date={c.deadline} small /></span>}
            {cst.overdue > 0 && <span className="shrink-0 rounded-full bg-rose-50 px-1.5 py-0.5 text-xs font-medium text-rose-600">{cst.overdue} od</span>}
            <span className="shrink-0"><StatusPill rag={crag} /></span>
            <button onClick={() => onOpen(c.id)} className="shrink-0 rounded p-1 text-slate-300 hover:text-slate-700" title="Open detail"><ArrowRight size={13} /></button>
          </div>
        );
      })}
    </div>
  );
  const canMoveActivity = (a) => isOrg || a.assigneeId === user.id;
  const moveActivity = (id, patch) => { patchAct(id, patch); flash("Activity updated."); };
  const activitiesKanban = <ActivitiesKanban acts={nodeActs} canMove={canMoveActivity} onMove={moveActivity} onOpen={onOpen} />;
  const activitiesList = (
    <div className="space-y-2">
      {nodeActs.length === 0 && <div className="rounded-lg border border-dashed border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-400">No activities yet{isOrg ? " — use AI actions → Auto-assign / Fill missing, or Add task." : "."}</div>}
      {nodeActs.map((a) => { const Icon = ACT_ICON[a.actType] || User; const over = isOverdue(a);
        return (
          <div key={a.id} className={`overflow-hidden rounded-lg border bg-white ${a.blocked ? "border-rose-300" : over ? "border-rose-200" : "border-slate-200"}`}>
            <div className={`px-3 py-2 ${a.status === "executed" ? "bg-blue-50" : a.blocked ? "bg-rose-50" : over ? "bg-rose-50" : a.assigneeId && a.date ? "" : "bg-amber-50"}`}>
              <div className="flex items-center gap-2">
                <Icon size={14} className="shrink-0 text-slate-400" />
                <span className={`min-w-0 flex-1 truncate text-sm ${a.status === "cancelled" ? "text-slate-400 line-through" : "text-slate-800"}`}>{a.title}{a.unplanned && <span className="ml-1 rounded bg-orange-100 px-1 text-xs text-orange-700">unplanned</span>}</span>
                {a.blocked && <Chip tone="rose"><AlertTriangle size={10} /> blocked</Chip>}{a.status === "executed" && <Chip tone="blue"><Check size={10} /> done</Chip>}
                {isOrg && (confirmDel === a.id
                  ? <span className="flex shrink-0 items-center gap-1"><span className="text-xs text-rose-600">Delete?</span><button onClick={() => delAct(a.id)} className="rounded bg-rose-500 p-1 text-white hover:bg-rose-600"><Check size={12} /></button><button onClick={() => setConfirmDel(null)} className="rounded border border-slate-200 p-1 text-slate-400 hover:bg-slate-50"><X size={12} /></button></span>
                  : <span className="flex shrink-0 items-center gap-1"><button onClick={() => setEdit(a)} className="rounded border border-slate-200 p-1 text-slate-400 hover:bg-slate-50 hover:text-slate-700"><Pencil size={12} /></button><button onClick={() => setConfirmDel(a.id)} className="rounded border border-slate-200 p-1 text-slate-400 hover:bg-rose-50 hover:text-rose-600"><Trash2 size={12} /></button></span>)}
              </div>
              <div className="mt-1.5 flex flex-wrap items-center gap-x-2 gap-y-1 pl-6 text-xs">
                {a.assigneeId ? <span className="inline-flex items-center gap-1 text-slate-600"><Avatar id={a.assigneeId} size={16} /> {uFirst(a.assigneeId)}</span> : <span className="text-amber-700">Unassigned</span>}
                {(a.startDate || a.date) ? <span className={`inline-flex items-center gap-1 ${over ? "text-rose-600" : "text-slate-500"}`}><CalendarClock size={11} className="shrink-0" />{a.startDate ? fmtShort(a.startDate) : "?"} <ArrowRight size={9} className="shrink-0 text-slate-300" /> {a.date ? fmtShort(a.date) : "?"}</span> : <span className="text-amber-700">No dates</span>}
                <span className="text-slate-500">{a.actType}</span>
                <span className="inline-flex items-center gap-1 font-mono text-slate-400" title="actual / estimated effort (hours)"><Clock size={10} className="shrink-0" />{a.actualHrs || 0}/{a.plannedHrs}h{a.plannedHrs > 0 && a.actualHrs > a.plannedHrs ? <span className="text-rose-600">▲</span> : null}</span>
              </div>
              {a.description && <div className="mt-1 pl-6 text-xs text-slate-500">{a.description}</div>}
              {a.blocked && <div className="mt-1.5 pl-6 text-xs text-rose-600">Blocked — this task can't move forward until it's cleared. It's what's turning the levels above red.</div>}
            </div>
          </div>
        );
      })}
    </div>
  );

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center gap-1.5 text-sm">
        <button onClick={() => onOpen(node.parentId)} className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-medium text-slate-600 hover:bg-slate-50"><ChevronLeft size={14} /> Back</button>
        <span className="mx-0.5 text-slate-300">|</span>
        <button onClick={() => onOpen(null)} className="text-slate-500 hover:text-slate-800">Portfolio</button>
        {path.map((n, i) => { const t = LEVEL_THEME[n.level]; const last = i === path.length - 1; return <span key={n.id} className="flex items-center gap-1.5"><ChevronRight size={14} className="text-slate-300" /><span className={`h-2 w-2 rounded-full ${t.bar}`} />{last ? <span className="font-medium text-slate-800">{n.title}</span> : <button onClick={() => onOpen(n.id)} className="max-w-xs truncate text-slate-500 hover:text-slate-800">{n.title}</button>}</span>; })}
      </div>

      {/* one card for every level: this node + its children (initiatives / works / activities) inside it */}
      <div className="relative overflow-hidden rounded-xl border border-slate-200 bg-white">
        <span className={`absolute inset-y-0 left-0 w-1 ${th.bar}`} />
        {headerBlock}
        <div className="border-t border-slate-100 px-5 pb-5 pt-4">
          {insideDivider}
          {leaf ? (childView === "kanban" ? activitiesKanban : activitiesList) : (childView === "list" ? childrenList : childrenGrid)}
          {leaf && <WorkDeliverables work={node} acts={acts} store={store} canManage={canRemark} busy={busy} setBusy={setBusy} flash={flash} />}
        </div>
      </div>


      {modify && <ModifyPlan {...{ work: node, planText: planText(), subs, acts, store, busy, setBusy, flash, onClose: () => setModify(false) }} />}
      {addUn && <AddUnplanned {...{ subs, store, flash, onClose: () => setAddUn(false) }} />}
      {edit && <ActivityEdit {...{ activity: edit, onSave: (p) => { patchAct(edit.id, p); setEdit(null); flash("Activity updated."); }, onClose: () => setEdit(null) }} />}
      {addChild && <QuickCreate {...{ me: user, parent: node, level: childLevel, works, store, busy, setBusy, flash, onClose: () => setAddChild(false), onCreated: (id) => { if (childLevel !== "activity") onOpen(id); } }} />}
      {(insight || suggest) && <Modal onClose={closeResults}>
        <div className="mb-3 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><Sparkles size={15} className="text-indigo-600" /> {suggest ? "Suggested initiatives" : `Where ${node.level === "objective" ? "this stands" : "you stand"}`}</h3><button onClick={closeResults} className="text-slate-400"><X size={18} /></button></div>
        <div className="mb-3 rounded-md bg-slate-50 px-3 py-2 text-sm text-slate-700">{node.title}</div>
        {insight && <div className="space-y-2 text-sm"><p className="text-slate-700">{insight.read}</p><p className="rounded-md bg-indigo-50 px-3 py-2 text-indigo-800"><span className="font-medium">Next: </span>{insight.action}</p></div>}
        {suggest && <div className="space-y-1.5">{suggest.length === 0 && <div className="text-sm text-slate-500">No new suggestions.</div>}{suggest.map((s, i) => <div key={i} className="flex items-center gap-2 rounded-md border border-slate-100 bg-white px-3 py-2 text-sm"><span className="min-w-0 flex-1 truncate text-slate-700">{s.title}</span><Chip tone="blue">{s.type}</Chip><button onClick={() => addSuggested(s)} className="shrink-0 rounded bg-slate-900 px-2 py-1 text-xs font-medium text-white hover:bg-slate-700">Add</button></div>)}</div>}
      </Modal>}
      {blockerOpen && issue && <Modal onClose={() => setBlockerOpen(false)}>
        <div className="mb-3 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-rose-800"><AlertTriangle size={15} className="text-rose-500" /> Needs attention</h3><button onClick={() => setBlockerOpen(false)} className="text-slate-400"><X size={18} /></button></div>
        <div className="text-sm text-slate-700">{att.blocked > 0 ? `${att.blocked} blocked` : ""}{att.blocked > 0 && att.overdue > 0 ? " and " : ""}{att.overdue > 0 ? `${att.overdue} overdue` : ""} deeper down — that's why this {LEVEL_LABEL[node.level].toLowerCase()} shows red.</div>
        <div className="mt-2 rounded-md bg-rose-50 px-3 py-2 text-sm text-rose-800"><span className="font-medium">Root cause:</span> {issue.title} — {issue.blocked ? "a task that's blocked (someone can't move forward until it's cleared)." : "a task past its due date."}</div>
        {issue.workId !== node.id && <button onClick={() => { setBlockerOpen(false); onOpen(issue.workId); }} className={`${btnDark} mt-4 w-full`}>Go to the cause <ArrowRight size={14} /></button>}
      </Modal>}
      {approvalsOpen && <Modal onClose={() => setApprovalsOpen(false)}>
        <div className="mb-3 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><ClipboardCheck size={15} className="text-amber-600" /> Plan changes awaiting approval</h3><button onClick={() => setApprovalsOpen(false)} className="text-slate-400"><X size={18} /></button></div>
        <div className="space-y-1.5">{myCRs.map((c) => <div key={c.id} className="rounded-md border border-slate-100 bg-slate-50 px-3 py-2 text-sm"><div className="font-medium text-slate-800">{uName(c.proposerId)} — {c.kind.replace("_", " ")}</div><div className="text-xs text-slate-500">{c.desc}</div></div>)}</div>
        <button onClick={() => { setApprovalsOpen(false); goApprovals(); }} className={`${btnDark} mt-4 w-full`}><ClipboardCheck size={14} /> Review in Approvals <ArrowRight size={14} /></button>
      </Modal>}
    </div>
  );
}
function ResultCard({ work, m, isOrg, patchWork }) {
  const [edit, setEdit] = useState(false); const [val, setVal] = useState(work.result.current); const rt = work.result;
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <div className="mb-1 flex items-center justify-between"><span className="text-xs font-medium uppercase tracking-wide text-slate-400">Result — planned vs actual</span>{isOrg && <button onClick={() => setEdit(!edit)} className="text-xs text-blue-700">{edit ? "close" : "update actual"}</button>}</div>
      <div className="text-sm text-slate-700">{rt.metric}</div>
      <div className="mt-1 flex items-baseline gap-2"><span className="font-mono text-xl font-medium text-slate-900">{rt.current}</span><span className="text-xs text-slate-400">now · planned target {rt.target} {rt.unit} (from {rt.baseline})</span></div>
      {m.resultPct != null && <div className="mt-2 flex items-center gap-2"><div className="h-2 flex-1 rounded-full bg-blue-100 overflow-hidden"><div className="h-full rounded-full bg-blue-700" style={{ width: `${m.resultPct}%` }} /></div><span className="font-mono text-xs font-medium text-blue-800">{Math.round(m.resultPct)}%</span></div>}
      {edit && <div className="mt-3 flex items-center gap-2"><input type="number" value={val} onChange={(e) => setVal(e.target.value)} className="w-28 rounded-md border border-slate-200 px-2 py-1 text-sm" /><button onClick={() => { patchWork(work.id, { result: { ...rt, current: Number(val) } }); setEdit(false); }} className={btnDark}>Save</button></div>}
    </div>
  );
}
function ActivityEdit({ activity, onSave, onClose }) {
  const [title, setTitle] = useState(activity.title); const [desc, setDesc] = useState(activity.description || ""); const [assignee, setAssignee] = useState(activity.assigneeId || ""); const [date, setDate] = useState(activity.date || ""); const [type, setType] = useState(activity.actType); const [hrs, setHrs] = useState(activity.plannedHrs);
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Edit activity</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="space-y-3">
        <div><label className="mb-1 block text-xs text-slate-500">Title</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={inputCls} /></div>
        <div><label className="mb-1 block text-xs text-slate-500">Description — what needs to be produced (used to score the deliverable)</label><textarea value={desc} onChange={(e) => setDesc(e.target.value)} rows={2} className={inputCls} placeholder="e.g. A one-page cost sheet with three vendor quotes compared." /></div>
        <div><label className="mb-1 block text-xs text-slate-500">Assign to</label><select value={assignee} onChange={(e) => setAssignee(e.target.value)} className={inputCls}><option value="">Unassigned</option>{USERS.map((u) => <option key={u.id} value={u.id}>{u.name} — {u.title}</option>)}</select></div>
        <div className="grid grid-cols-1 gap-2 sm:grid-cols-3">
          <div className="sm:col-span-2"><label className="mb-1 block text-xs text-slate-500">Due date</label><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} /></div>
          <div><label className="mb-1 block text-xs text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div>
        </div>
        <div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
      </div>
      <button onClick={() => onSave({ title: title.trim() || activity.title, description: desc.trim() || null, assigneeId: assignee || null, date: date || null, actType: type, plannedHrs: hrs })} className={`${btnDark} mt-4 w-full`}>Save changes</button>
    </Modal>
  );
}

/* ---------- My day (real calendar) ---------- */
// One-click "push my upcoming tasks to Outlook", with a subtle in-sync indicator.
// Reused on My day, My work, and the Data tab. Syncs only the signed-in user's
// own activities (create / update moved dates / remove completed).
// Sync upcoming tasks into the signed-in user's Outlook calendar. By default it
// takes that user's own assigned activities; pass `scopeActs` (e.g. one work's
// activities) to sync exactly that list instead — used on the work detail page.
function OutlookSyncButton({ me, acts, works, patchAct, flash, scopeActs, label = "Add to Outlook" }) {
  const [syncing, setSyncing] = useState(false);
  if (!MSAL_CONFIGURED) return null;
  const mine = scopeActs || acts.filter((a) => a.assigneeId === me.id);
  const upcoming = mine.filter((a) => a.date && a.status !== "executed" && a.status !== "cancelled");
  const unsynced = upcoming.filter((a) => !a.outlookEventId).length;
  const wTitle = (id) => works.find((w) => w.id === id)?.title || "";
  if (scopeActs && upcoming.length === 0) return null; // nothing datable on this work → hide entirely
  const sync = async () => {
    setSyncing(true);
    try {
      const tasks = mine.map((a) => ({ id: a.id, title: a.title, description: a.description, date: a.date, planned: a.plannedHrs, actual: a.actualHrs, status: a.status, outlookEventId: a.outlookEventId, workTitle: wTitle(a.workId) }));
      const r = await syncTasksToOutlook(tasks, (id, evId) => patchAct(id, { outlookEventId: evId }));
      flash(`Outlook calendar synced — ${r.created} added, ${r.updated} updated${r.removed ? `, ${r.removed} removed` : ""}.`);
    } catch (e) { flash(e.message || "Couldn't reach your Outlook calendar."); }
    setSyncing(false);
  };
  return (
    <div className="flex items-center gap-2">
      {upcoming.length > 0 && (unsynced > 0
        ? <span className="hidden text-xs text-amber-600 sm:inline">{unsynced} to sync</span>
        : <span className="hidden items-center gap-1 text-xs text-emerald-600 sm:inline-flex"><Check size={12} /> In sync</span>)}
      <button onClick={sync} disabled={syncing} className={btnLight} title="Add these upcoming tasks to your Outlook / Microsoft 365 calendar">{syncing ? <><Loader2 size={14} className="animate-spin" /> Syncing…</> : <><CalendarClock size={14} /> {label}</>}</button>
    </div>
  );
}
function MyDay({ me, works, acts, busy, setBusy, flash, patchAct, store }) {
  const [anchor, setAnchor] = useState(startOfWeek(TODAY));
  const [sel, setSel] = useState(iso(TODAY));
  const [quick, setQuick] = useState(false); const [propose, setPropose] = useState(null);
  const week = [0, 1, 2, 3, 4, 5, 6].map((i) => addDays(anchor, i));
  const wTitle = (id) => works.find((w) => w.id === id)?.title || "";
  const mineAll = acts.filter((a) => a.assigneeId === me.id && a.status !== "cancelled");
  const mine = mineAll.filter((a) => a.date === sel);
  const overdue = mineAll.filter((a) => isOverdue(a));
  const undated = mineAll.filter((a) => !a.date);
  const pool = acts.filter((a) => !a.assigneeId && a.status === "planned");
  const weekActs = mineAll.filter((a) => a.date && week.some((d) => iso(d) === a.date));
  const load = mine.reduce((s, a) => s + a.plannedHrs, 0); const cap = 7;
  const shift = (n) => { setAnchor(addDays(anchor, n * 7)); setSel(iso(addDays(parseISO(sel), n * 7))); };
  const countOn = (d) => mineAll.filter((a) => a.date === iso(d)).length;

  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <div><div className="text-lg font-medium text-slate-900">My day</div><div className="text-xs text-slate-400">Your assigned activities, by day. Sync the upcoming ones to your Outlook calendar.</div></div>
        <OutlookSyncButton me={me} acts={acts} works={works} patchAct={patchAct} flash={flash} />
      </div>
      {/* week calendar */}
      <div className="mb-4 rounded-xl border border-slate-200 bg-white p-3">
        <div className="mb-2 flex items-center justify-between px-1">
          <div className="text-sm font-medium text-slate-700">{anchor.toLocaleDateString("en-GB", { month: "long", year: "numeric" })}</div>
          <div className="flex items-center gap-1"><button onClick={() => shift(-1)} className={btnLight}><ChevronLeft size={14} /></button><button onClick={() => { setAnchor(startOfWeek(TODAY)); setSel(iso(TODAY)); }} className={btnLight}>Today</button><button onClick={() => shift(1)} className={btnLight}><ChevronRight size={14} /></button></div>
        </div>
        <div className="grid grid-cols-7 gap-1">
          {week.map((d) => { const on = iso(d) === sel; const today = iso(d) === iso(TODAY); const c = countOn(d);
            return (
              <button key={iso(d)} onClick={() => setSel(iso(d))} className={`rounded-lg border px-1 py-2 text-center ${on ? "border-slate-900 bg-slate-900 text-white" : today ? "border-slate-300 bg-slate-50" : "border-slate-100 hover:bg-slate-50"}`}>
                <div className={`text-xs ${on ? "text-slate-300" : "text-slate-400"}`}>{d.toLocaleDateString("en-GB", { weekday: "short" })}</div>
                <div className={`font-mono text-base ${on ? "text-white" : today ? "text-slate-900 font-medium" : "text-slate-700"}`}>{d.getDate()}</div>
                <div className="mt-0.5 flex h-1.5 justify-center">{c > 0 && <span className={`h-1.5 w-1.5 rounded-full ${on ? "bg-white" : "bg-amber-400"}`} />}</div>
              </button>
            );
          })}
        </div>
      </div>

      <div className="grid grid-cols-1 gap-5 lg:grid-cols-3">
        <div className="lg:col-span-2 space-y-4">
          {overdue.length > 0 && (
            <div className="rounded-xl border border-rose-200 bg-rose-50 p-3">
              <div className="mb-1 flex items-center gap-1.5 text-sm font-medium text-rose-700"><AlertCircle size={15} /> {overdue.length} overdue</div>
              <div className="space-y-1">{overdue.map((a) => <div key={a.id} className="flex items-center gap-2 text-sm"><span className="min-w-0 flex-1 truncate text-slate-700">{a.title}</span><span className="text-xs text-rose-600">{fmtFull(parseISO(a.date))}</span><button onClick={() => patchAct(a.id, { date: sel })} className="text-xs font-medium text-slate-500 hover:text-slate-800">move to {new Date(sel).toLocaleDateString("en-GB", { weekday: "short" })}</button></div>)}</div>
            </div>
          )}
          <div className="rounded-xl border border-slate-200 bg-white p-4">
            <div className="mb-3 flex items-center justify-between">
              <div className="text-sm font-medium text-slate-800">{fmtFull(parseISO(sel))}</div>
              <div className="flex items-center gap-3 text-xs text-slate-400"><span>Load <span className="font-mono text-slate-600">{load}h</span>/{cap}h</span><button onClick={() => setQuick(true)} className="inline-flex items-center gap-1 font-medium text-slate-500 hover:text-slate-800"><Plus size={12} /> add</button></div>
            </div>
            {mine.length === 0 && <div className="py-6 text-center text-sm text-slate-400">Nothing on this day. Pick from the pool → or add one.</div>}
            <div className="space-y-2">
              {mine.map((a) => { const Icon = ACT_ICON[a.actType] || User;
                return (
                  <div key={a.id} className={`rounded-lg border px-3 py-2.5 ${a.status === "executed" ? "border-blue-200 bg-blue-50" : a.inProgress ? "border-amber-300 bg-amber-50" : "border-slate-200 bg-white"}`}>
                    <div className="flex items-center gap-3">
                      <Icon size={15} className="shrink-0 text-slate-400" />
                      <div className="min-w-0 flex-1"><div className={`truncate text-sm ${a.status === "executed" ? "text-slate-500" : "text-slate-800"}`}>{a.title}</div><div className="truncate text-xs text-slate-400">{wTitle(a.workId)} · {a.actType} · {a.plannedHrs}h</div></div>
                      {a.status === "executed" ? <Chip tone="blue"><Check size={11} /> Done</Chip> : a.inProgress ? <button onClick={() => patchAct(a.id, { status: "executed", inProgress: false, actualHrs: a.plannedHrs })} className="inline-flex items-center gap-1 rounded-md bg-amber-500 px-2 py-1 text-xs font-medium text-white"><Square size={11} /> Stop</button> : <button onClick={() => patchAct(a.id, { inProgress: true })} className="inline-flex items-center gap-1 rounded-md bg-emerald-600 px-2 py-1 text-xs font-medium text-white"><Play size={11} /> Start</button>}
                    </div>
                    <div className="mt-1.5 flex items-center gap-3 pl-6">
                      <button onClick={() => setPropose(a)} className="inline-flex items-center gap-1 text-xs text-blue-700 hover:text-blue-900"><Pencil size={11} /> propose change</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
        <div className="space-y-3">
          <div className="rounded-xl border border-slate-200 bg-white p-4">
            <div className="mb-2 flex items-center gap-1.5 text-xs font-medium uppercase tracking-wider text-slate-400"><ClipboardList size={13} /> This week</div>
            {weekActs.length === 0 && <div className="text-xs text-slate-400">No activities scheduled this week.</div>}
            <div className="space-y-1.5">{week.map((d) => { const day = weekActs.filter((a) => a.date === iso(d)); if (!day.length) return null; return <div key={iso(d)}><div className="text-xs font-medium text-slate-500">{d.toLocaleDateString("en-GB", { weekday: "short", day: "numeric" })}</div>{day.map((a) => <div key={a.id} className="flex items-center gap-1.5 pl-2 text-xs text-slate-600"><span className={`h-1.5 w-1.5 rounded-full ${a.status === "executed" ? "bg-blue-400" : "bg-slate-300"}`} /><span className="truncate">{a.title}</span></div>)}</div>; })}</div>
          </div>
          <div className="rounded-xl border border-dashed border-slate-300 bg-white p-4">
            <div className="mb-1 flex items-center gap-1.5 text-xs font-medium uppercase tracking-wider text-slate-400"><Users size={13} /> Waiting to be scheduled</div>
            <p className="mb-3 text-xs text-slate-400">Placing these lifts the planning meters.</p>
            <div className="space-y-2">
              {undated.map((a) => <div key={a.id} className="rounded-md border border-slate-200 bg-white px-3 py-2"><div className="truncate text-sm text-slate-700">{a.title}</div><div className="mb-2 truncate text-xs text-slate-400">assigned to me · no date</div><button onClick={() => patchAct(a.id, { date: sel })} className="w-full rounded-md bg-slate-900 px-2 py-1 text-xs font-medium text-white">Set to {new Date(sel).toLocaleDateString("en-GB", { weekday: "short", day: "numeric" })}</button></div>)}
              {pool.map((a) => <div key={a.id} className="rounded-md border border-slate-200 bg-slate-50 px-3 py-2"><div className="truncate text-sm text-slate-700">{a.title}</div><div className="mb-2 truncate text-xs text-slate-400">{wTitle(a.workId)} · open pool</div><button onClick={() => patchAct(a.id, { assigneeId: me.id, date: sel })} className="w-full rounded-md bg-slate-900 px-2 py-1 text-xs font-medium text-white">Take &amp; set to {new Date(sel).toLocaleDateString("en-GB", { weekday: "short", day: "numeric" })}</button></div>)}
              {undated.length === 0 && pool.length === 0 && <div className="text-xs text-slate-400">All caught up.</div>}
            </div>
          </div>
        </div>
      </div>
      {quick && <QuickAdd {...{ me, works, date: sel, store, flash, onClose: () => setQuick(false) }} />}
      {propose && <ProposeChange {...{ activity: propose, me, store, flash, onClose: () => setPropose(null) }} />}
    </div>
  );
}
function QuickAdd({ me, works, date, store, flash, onClose }) {
  const leaves = works.filter((w) => w.level === "work");
  const [wid, setWid] = useState(leaves[0] ? leaves[0].id : ""); const [title, setTitle] = useState(""); const [desc, setDesc] = useState(""); const [hrs, setHrs] = useState(1); const [type, setType] = useState("self"); const [due, setDue] = useState(date);
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Add an activity</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Assigned to you. Add a description so the deliverable can be scored against it later.</p>
      <label className="mb-1 block text-xs text-slate-500">Under work</label><select value={wid} onChange={(e) => setWid(e.target.value)} className={`${inputCls} mb-3`}>{leaves.map((w) => <option key={w.id} value={w.id}>{w.title}</option>)}</select>
      <label className="mb-1 block text-xs text-slate-500">Activity title</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder="What are you doing?" />
      <label className="mb-1 block text-xs text-slate-500">Description — type, speak, or attach (what needs to be produced)</label>
      <div className="mb-3"><MultiModalInput value={desc} onChange={setDesc} placeholder="Describe the task or the deliverable — you can dictate it…" /></div>
      <div className="mb-4 grid grid-cols-3 gap-2"><div><label className="mb-1 block text-xs text-slate-500">Due date</label><input type="date" value={due} onChange={(e) => setDue(e.target.value)} className={inputCls} /></div><div><label className="mb-1 block text-xs text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
      <button onClick={() => { if (!title.trim() || !wid) return; store.addActs([{ id: nid("a"), workId: wid, title: title.trim(), description: desc.trim() || null, assigneeId: me.id, date: due || date, status: "planned", plannedHrs: hrs, actualHrs: null, actType: type }]); flash("Added to your day."); onClose(); }} disabled={!title.trim()} className={`${btnDark} w-full`}>Add</button>
    </Modal>
  );
}

/* ---------- Team ---------- */
function TeamView({ user, teams, store, works, acts, onOpen }) {
  const scopedTeams = user.level === "md" ? teams : teams.filter((t) => t.memberIds.some((id) => id === user.id || fnOf(id) === user.fn || USERS.find((u) => u.id === id)?.reports_to === user.id));
  const [sel, setSel] = useState("all");
  const [open, setOpen] = useState(null);
  const [showCreate, setShowCreate] = useState(false);
  const shown = sel === "all" ? scopedTeams : scopedTeams.filter((t) => t.id === sel);
  const stats = (uid) => { const mine = acts.filter((a) => a.assigneeId === uid && a.status !== "cancelled"); const done = mine.filter((a) => a.status === "executed"); const over = mine.filter((a) => isOverdue(a)); const ds = works.filter((w) => w.ownerId === uid).flatMap((w) => w.deliverables || []).filter((d) => d.done && typeof d.score === "number"); return { assigned: mine.length, done: done.length, over: over.length, exec: mine.length ? Math.round((done.length / mine.length) * 100) : 0, avg: ds.length ? Math.round(ds.reduce((s, d) => s + d.score, 0) / ds.length) : null }; };
  const wt = (id) => { const w = works.find((x) => x.id === id); if (!w) return ""; const top = w.parentId ? works.find((x) => x.id === w.parentId) : w; return `${top ? top.title.slice(0, 24) : ""} › ${w.title}`; };
  const teamExec = (t) => { const es = t.memberIds.map((id) => stats(id)).filter((s) => s.assigned > 0); return es.length ? Math.round(es.reduce((a, s) => a + s.exec, 0) / es.length) : 0; };
  const key = (tid, uid) => tid + "|" + uid;
  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center gap-1">
        <button onClick={() => setSel("all")} className={`rounded-md px-3 py-1.5 text-xs font-medium ${sel === "all" ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>All teams</button>
        {scopedTeams.map((t) => <button key={t.id} onClick={() => setSel(t.id)} className={`rounded-md px-3 py-1.5 text-xs font-medium ${sel === t.id ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>{t.name} <span className="opacity-60">{t.memberIds.length}</span></button>)}
        {user.level === "md" && <button onClick={() => setShowCreate(true)} className={`${btnDark} ml-auto`}><Plus size={14} /> Create team</button>}
      </div>
      {user.level === "md" && <div className="mb-3 text-xs text-slate-400">Build a team from anyone across functions, then it's available when creating initiatives.</div>}
      <div className="space-y-4">
        {shown.map((t) => (
          <div key={t.id} className="rounded-xl border border-slate-200 bg-white">
            <div className="flex items-center justify-between border-b border-slate-100 px-4 py-3">
              <div className="flex items-center gap-2"><Users size={16} className="text-slate-400" /><span className="text-sm font-medium text-slate-800">{t.name}</span><span className="text-xs text-slate-400">· {t.memberIds.length} members · {[...new Set(t.memberIds.map(fnOf))].join(", ")}</span></div>
              <div className="text-right"><div className="font-mono text-sm font-medium text-amber-700">{teamExec(t)}%</div><div className="text-xs text-slate-400">avg execution</div></div>
            </div>
            <div className="divide-y divide-slate-50">
              {t.memberIds.map((uid) => { const u = USERS.find((x) => x.id === uid); if (!u) return null; const st = stats(uid); const isOpen = open === key(t.id, uid); const mine = acts.filter((a) => a.assigneeId === uid && a.status !== "cancelled");
                return (
                  <div key={uid}>
                    <button onClick={() => setOpen(isOpen ? null : key(t.id, uid))} className="flex w-full items-center gap-4 px-4 py-3 text-left hover:bg-slate-50">
                      <Avatar id={uid} size={34} />
                      <div className="min-w-0 flex-1"><div className="text-sm font-medium text-slate-800">{u.name}</div><div className="flex flex-wrap items-center gap-1.5 text-xs text-slate-400"><span>{u.title}</span><span>·</span><span>{st.assigned} activities</span><span>·</span><span>{st.done} done</span>{st.over > 0 && <Chip tone="rose">{st.over} overdue</Chip>}{st.avg != null && <Chip tone="blue">deliverables {st.avg}/100</Chip>}</div></div>
                      <div className="text-right"><div className="font-mono text-base font-medium text-amber-700">{st.exec}%</div><div className="text-xs text-slate-400">execution</div></div>
                    </button>
                    {isOpen && <div className="bg-slate-50 px-4 py-3"><div className="space-y-1.5">{mine.length === 0 && <div className="text-xs text-slate-400">No activities.</div>}{mine.map((a) => <div key={a.id} className="flex items-center gap-2 rounded-md border border-slate-100 bg-white px-3 py-1.5 text-sm"><button onClick={() => { const top = works.find((x) => x.id === a.workId); onOpen(top && top.parentId ? top.parentId : a.workId); }} className="min-w-0 flex-1 truncate text-left text-slate-700 hover:text-slate-900">{a.title}</button><span className="hidden shrink-0 truncate text-xs text-slate-400 sm:inline" style={{ maxWidth: 180 }}>{wt(a.workId)}</span><span className="shrink-0 text-xs text-slate-400">{a.date ? fmtFull(parseISO(a.date)) : "—"}</span><Chip tone={a.status === "executed" ? "blue" : isOverdue(a) ? "rose" : "slate"}>{a.status === "executed" ? "done" : isOverdue(a) ? "overdue" : a.status}</Chip></div>)}</div></div>}
                  </div>
                );
              })}
            </div>
          </div>
        ))}
        {shown.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-400">No teams in your scope yet.</div>}
      </div>
      {showCreate && <TeamModal teams={teams} store={store} onSelect={() => {}} onClose={() => setShowCreate(false)} />}
    </div>
  );
}

/* ---------- Approvals ---------- */
function Approvals({ crs, store, works, me, flash }) {
  const pending = crs.filter((c) => c.status === "pending");
  const [remarks, setRemarks] = useState({}); const [spin, setSpin] = useState({});
  const wt = (id) => works.find((w) => w.id === id)?.title || "";
  const KL = { add_activity: "add follow-up", extend: "needs more time", blocked: "flag blocked", reassign: "change owner", retype: "reclassify" };
  const detail = (cr) => cr.kind === "add_activity" ? `add “${cr.payload.title}” (${cr.payload.hrs}h)` : cr.kind === "extend" ? `+${cr.payload.hrs}h to this activity` : cr.kind === "reassign" ? `hand to ${uName(cr.payload.to)}` : cr.kind === "blocked" ? "mark blocked, needs help" : `retype to ${cr.payload.type}`;
  const decide = (cr, ok) => {
    const remark = remarks[cr.id] || "";
    const doSpin = ok && !!spin[cr.id] && remark.trim();
    // The service applies the change (add_activity / extend / blocked / reassign /
    // retype) and spins off a follow-up work under the initiative — same logic
    // the Teams agent uses, so the portal and Teams stay consistent.
    store.decideCr(cr.id, { approve: ok, remark, spinoff: !!spin[cr.id], approverId: me.id });
    flash(ok ? (doSpin ? "Approved — and a follow-up work was added under the initiative." : "Approved and applied to the plan.") : "Change rejected.");
  };
  return (
    <div className="space-y-3">
      <div className="text-xs font-medium text-slate-500">Plan changes proposed by the team — add a remark, then decide. Your remark can spin off a follow-up work (execution) under the initiative.</div>
      {pending.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-400">No pending approvals.</div>}
      {pending.map((cr) => (
        <div key={cr.id} className="rounded-xl border border-slate-200 bg-white p-4">
          <div className="flex items-center gap-1.5 text-sm font-medium text-slate-800"><Avatar id={cr.proposerId} size={20} /> {uName(cr.proposerId)}<Chip tone="violet">{KL[cr.kind]}</Chip><span className="font-normal text-slate-500">— {detail(cr)}</span></div>
          <div className="mt-1 text-xs text-slate-400">{wt(cr.workId)}</div>
          <div className="mt-2 rounded-md bg-slate-50 px-3 py-2 text-sm text-slate-600">{cr.desc}</div>
          <label className="mt-3 block text-xs text-slate-500">Your remark (optional)</label>
          <textarea value={remarks[cr.id] || ""} onChange={(e) => setRemarks((r) => ({ ...r, [cr.id]: e.target.value }))} rows={2} className={`${inputCls} mt-1`} placeholder="e.g. Approved — also add a hand-over checklist as a follow-up work." />
          <div className="mt-2 flex items-center justify-between">
            <label className="flex items-center gap-2 text-xs text-slate-500"><input type="checkbox" checked={!!spin[cr.id]} onChange={(e) => setSpin((s) => ({ ...s, [cr.id]: e.target.checked }))} /> Add a follow-up work from my remark</label>
            <div className="flex gap-2"><button onClick={() => decide(cr, true)} className="inline-flex items-center gap-1 rounded-md bg-emerald-600 px-3 py-1.5 text-sm font-medium text-white"><Check size={14} /> Approve</button><button onClick={() => decide(cr, false)} className={btnLight}><X size={14} /> Reject</button></div>
          </div>
        </div>
      ))}
    </div>
  );
}

/* ---------- Capture ---------- */
function Capture({ me, teams, store, works, busy, setBusy, flash, onClose, onOpen }) {
  const objectives = works.filter((w) => w.level === "objective");
  const [title, setTitle] = useState(""); const [type, setType] = useState("general"); const [parentObj, setParentObj] = useState(objectives[0] ? objectives[0].id : ""); const [objective, setObjective] = useState(""); const [deadline, setDeadline] = useState(""); const [teamId, setTeamId] = useState(teams[0] ? teams[0].id : ""); const [text, setText] = useState(""); const [teamModal, setTeamModal] = useState(false);
  const team = teams.find((t) => t.id === teamId);
  const build = (topTitle, ty, worksList) => {
    const topId = nid("w"); const tpl = METRIC_BY_TYPE[ty] || METRIC_BY_TYPE.general; const memberIds = team ? team.memberIds : [];
    const nw = [{ id: topId, parentId: parentObj || null, level: "initiative", title: topTitle, type: ty, ownerId: me.id, teamId: teamId || null, scope: memberIds.length > 1 ? "group" : "individual", objective: objective || null, deadline: deadline || null, result: { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 } }];
    const na = []; const load = {}; memberIds.forEach((id) => (load[id] = 0)); let ai = 0;
    const dl = deadline ? parseISO(deadline) : null; const span = dl ? Math.max(1, Math.round((dl - TODAY) / MSD)) : 5;
    (worksList || []).forEach((wk) => {
      const wid = nid("w"); nw.push({ id: wid, parentId: topId, level: "work", title: wk.title, type: ty, ownerId: me.id });
      (wk.activities || []).forEach((ac) => { let assignee = null, date = null; if (memberIds.length) { assignee = memberIds.reduce((a, b) => (load[a] <= load[b] ? a : b)); load[assignee] += Number(ac.estimateHrs) || 2; date = iso(addDays(TODAY, ai % span)); ai++; } na.push({ id: nid("a"), workId: wid, title: ac.title, assigneeId: assignee, date, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" }); });
    });
    store.addWorks(nw); if (na.length) store.addActs(na); return topId;
  };
  const run = async () => {
    if (!title.trim() && !text.trim()) return; setBusy("cap");
    try { const ctx = `${title.trim()}${objective ? ". Objective: " + objective : ""}${text ? ". Plan shared by the lead: " + text : ""}`.slice(0, 1500); const out = await AI.decompose(ctx, type); const id = build(title.trim() || text.slice(0, 50), type, out.works); flash(team ? `AI drafted the plan and assigned it to ${team.name}.` : "AI drafted the plan."); onOpen(id); }
    catch { const id = build(title.trim() || "New work", type, [{ title: "Plan", activities: [{ title: "Define scope & owners", estimateHrs: 2, type: "self" }] }]); flash("AI unavailable — created a starter you can edit."); onOpen(id); }
    setBusy(null);
  };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-4 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Capture &amp; plan a work</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 grid grid-cols-1 gap-2 sm:grid-cols-3"><div className="sm:col-span-2"><label className="mb-1 block text-xs text-slate-500">Title</label><input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Roll out MFA to all field staff" className={inputCls} /></div><div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{Object.keys(METRIC_BY_TYPE).map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
      <div className="mb-3"><label className="mb-1 block text-xs text-slate-500">Sits under objective</label><select value={parentObj} onChange={(e) => setParentObj(e.target.value)} className={inputCls}><option value="">None (top-level initiative)</option>{objectives.map((o) => <option key={o.id} value={o.id}>{o.title}</option>)}</select></div>
      <div className="mb-3 grid grid-cols-1 gap-2 sm:grid-cols-3">
        <div className="sm:col-span-2"><label className="mb-1 block text-xs text-slate-500">Assign to team</label><div className="flex gap-2"><select value={teamId} onChange={(e) => setTeamId(e.target.value)} className={inputCls}><option value="">No team (leave unassigned)</option>{teams.map((t) => <option key={t.id} value={t.id}>{t.name} ({t.memberIds.length})</option>)}</select><button onClick={() => setTeamModal(true)} className={btnLight}><Users size={14} /> New / merge</button></div></div>
        <div><label className="mb-1 block text-xs text-slate-500">Deadline</label><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className={inputCls} /></div>
      </div>
      {team && <div className="mb-3 flex flex-wrap items-center gap-1.5">{team.memberIds.map((id) => <span key={id} className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs text-slate-600"><Avatar id={id} size={16} /> {uFirst(id)} <span className="text-slate-400">{fnOf(id)}</span></span>)}</div>}
      <label className="mb-1 block text-xs text-slate-500">Objective (optional)</label><input value={objective} onChange={(e) => setObjective(e.target.value)} className={`${inputCls} mb-3`} placeholder="The result this should move" />
      <label className="mb-1 block text-xs text-slate-500">Plan from the head — type, speak, or attach (the AI uses this to build the works)</label>
      <MultiModalInput value={text} onChange={setText} placeholder="Paste or speak the plan the lead shared, or attach the note…" />
      <button onClick={run} disabled={busy || (!title.trim() && !text.trim())} className={`${btnViolet} mt-3 w-full`}>{busy ? <><Loader2 size={15} className="animate-spin" /> Drafting &amp; assigning…</> : <><Sparkles size={15} /> Draft plan &amp; assign to team</>}</button>
      <p className="mt-2 text-center text-xs text-slate-400">AI builds initiative → works → activities, and distributes them across the team by load.</p>
      {teamModal && <TeamModal {...{ teams, store, onSelect: setTeamId, onClose: () => setTeamModal(false) }} />}
    </Modal>
  );
}
function TeamModal({ teams, store, onSelect, onClose }) {
  const [name, setName] = useState(""); const [sel, setSel] = useState([]);
  const toggle = (id) => setSel((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const create = () => { if (!name.trim() || !sel.length) return; const id = nid("t"); store.addTeam({ id, name: name.trim(), memberIds: sel }); onSelect(id); onClose(); };
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Create / merge a team</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Pick people from any function — this is how you merge members across teams for one initiative.</p>
      <label className="mb-1 block text-xs text-slate-500">Team name</label><input value={name} onChange={(e) => setName(e.target.value)} className={`${inputCls} mb-3`} placeholder="e.g. MFA rollout squad" />
      <div className="mb-3 max-h-60 space-y-1 overflow-y-auto">{USERS.map((u) => <label key={u.id} className={`flex cursor-pointer items-center gap-2 rounded-md border px-3 py-2 text-sm ${sel.includes(u.id) ? "border-slate-900 bg-slate-50" : "border-slate-200"}`}><input type="checkbox" checked={sel.includes(u.id)} onChange={() => toggle(u.id)} /><Avatar id={u.id} size={22} /> <span className="flex-1">{u.name}</span><span className="text-xs text-slate-400">{u.fn} · {u.level}</span></label>)}</div>
      <button onClick={create} disabled={!name.trim() || !sel.length} className={`${btnDark} w-full`}>Create team ({sel.length})</button>
    </Modal>
  );
}

/* ---------- Remark / Nudge / Objective ---------- */
function RemarkModal({ node, works, acts, busy, setBusy, flash, onSubmit, onClose }) {
  const [text, setText] = useState("");
  const [owner, setOwner] = useState(node.ownerId);
  const [shift, setShift] = useState(0);
  const [ops, setOps] = useState(null);
  const chain = []; { let cur = node, g = 0; while (cur && g++ < 12) { chain.push(cur); if (cur.level === "initiative") break; cur = works.find((w) => w.id === cur.parentId); } }
  const targets = [...new Set(chain.map((n) => n.ownerId))];
  const th = LEVEL_THEME[node.level] || LEVEL_THEME.activity;
  const subs = works.filter((w) => w.parentId === node.id); const container = subs.length ? subs : [node];
  const planText = container.map((s) => `- ${s.title}: ${acts.filter((a) => a.workId === s.id).map((a) => a.title + " (" + a.actType + ")").join(", ") || "none"}`).join("\n");
  const draftOps = async () => { if (!text.trim()) return; setBusy("remarkops"); try { setOps((await AI.modifyPlan(planText, text)).ops || []); } catch { flash("AI unavailable — you can still send the remark."); } setBusy(null); };
  const opLabel = (op) => op.op === "add_activity" ? `Add task “${op.title}” to ${op.work || op.subwork || container[0].title}` : (op.op === "add_work" || op.op === "add_subwork") ? `Add work “${op.title}”` : `Retype “${op.match}” → ${op.type}`;
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Remark &amp; update plan</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 flex items-center gap-2 rounded-md bg-slate-50 px-3 py-2"><span className={`h-2.5 w-2.5 rounded-full ${th.bar}`} /><span className={`rounded px-1.5 py-0.5 text-xs font-medium ${th.chip}`}>{LEVEL_LABEL[node.level]}</span><span className="min-w-0 truncate text-sm font-medium text-slate-800">{node.title}</span></div>
      <label className="mb-1 block text-xs font-medium text-slate-600">Your remark</label>
      <textarea value={text} onChange={(e) => setText(e.target.value)} rows={3} className={`${inputCls} mb-3`} placeholder="e.g. Pull the vendor demo forward; loop in Finance before the PO goes out; HOD sign-off should be a meeting." />
      <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Re-assign owner (optional)</label><select value={owner} onChange={(e) => setOwner(e.target.value)} className={inputCls}>{USERS.map((u) => <option key={u.id} value={u.id}>{u.name} · {u.fn}</option>)}</select></div>
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Shift dates + deadline (days)</label><input type="number" value={shift} onChange={(e) => setShift(Number(e.target.value) || 0)} className={inputCls} placeholder="e.g. -3 or 5" /></div>
      </div>
      <div className="mb-3 rounded-lg border border-blue-200 bg-blue-50 p-3">
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-1.5 text-xs font-medium text-blue-900"><Sparkles size={14} /> Turn my remark into plan changes for the team below</div>
          <button onClick={draftOps} disabled={busy || !text.trim()} className={btnLight}>{busy === "remarkops" ? <Loader2 size={13} className="animate-spin" /> : <Wand2 size={13} />} Draft</button>
        </div>
        {ops && <div className="mt-2 space-y-1">{ops.length === 0 && <div className="text-xs text-slate-500">AI didn't propose concrete changes — the remark alone will be sent.</div>}{ops.map((op, i) => <div key={i} className="flex items-center gap-2 rounded-md bg-white px-3 py-1.5 text-sm text-slate-700"><span className="min-w-0 flex-1 truncate">{opLabel(op)}</span><button onClick={() => setOps(ops.filter((_, j) => j !== i))} className="shrink-0 text-slate-300 hover:text-rose-500"><X size={13} /></button></div>)}{ops.length > 0 && <div className="text-xs text-blue-800">These apply to the plan beneath this {LEVEL_LABEL[node.level].toLowerCase()} when you send.</div>}</div>}
      </div>
      <div className="mb-3 rounded-md border border-blue-100 bg-blue-50 px-3 py-2 text-xs text-blue-700">Will nudge: {targets.map((id) => uName(id)).join(", ")} <span className="text-blue-400">— owner{targets.length > 1 ? "s" : ""} up the chain to the initiative.</span></div>
      <button onClick={() => { if (!text.trim()) return; onSubmit({ text: text.trim(), newOwnerId: owner !== node.ownerId ? owner : null, shiftDays: shift, ops: ops || [] }); }} disabled={!text.trim()} className={`${btnDark} w-full`}>Send remark, apply changes &amp; nudge</button>
    </Modal>
  );
}

function NudgeInbox({ remarks, me, onOpen, onClose }) {
  const mine = remarks.filter((r) => r.toIds.includes(me.id));
  const ago = (ts) => { const d = Math.round((Date.now() - ts) / 60000); if (d < 1) return "just now"; if (d < 60) return d + "m ago"; const h = Math.round(d / 60); if (h < 24) return h + "h ago"; return Math.round(h / 24) + "d ago"; };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Your nudges</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      {mine.length === 0 && <div className="rounded-md border border-dashed border-slate-200 px-3 py-8 text-center text-sm text-slate-400">No nudges yet. When a leader remarks on your initiative or work, it lands here.</div>}
      <div className="space-y-2">{mine.map((r) => { const th = LEVEL_THEME[r.level] || LEVEL_THEME.activity; return (
        <div key={r.id} className="rounded-lg border border-slate-200 p-3">
          <div className="flex items-center gap-2"><Avatar id={r.fromId} size={20} /><span className="text-sm font-medium text-slate-800">{uName(r.fromId)}</span><span className="text-xs text-slate-400">on</span><span className={`rounded px-1.5 py-0.5 text-xs font-medium ${th.chip}`}>{LEVEL_LABEL[r.level]}</span><span className="min-w-0 flex-1 truncate text-xs text-slate-500">{r.title}</span><span className="shrink-0 text-xs text-slate-300">{ago(r.ts)}</span></div>
          <div className="mt-1.5 text-sm text-slate-700">{r.text}</div>
          {(r.changes.newOwnerId || r.changes.shiftDays || r.changes.ops) && <div className="mt-1.5 flex flex-wrap gap-1.5">{r.changes.newOwnerId && <Chip tone="amber">re-assigned → {uName(r.changes.newOwnerId)}</Chip>}{r.changes.shiftDays ? <Chip tone="amber">dates {r.changes.shiftDays > 0 ? "+" : ""}{r.changes.shiftDays}d</Chip> : null}{r.changes.ops ? <Chip tone="violet">{r.changes.ops} plan change{r.changes.ops !== 1 ? "s" : ""} applied</Chip> : null}</div>}
          <button onClick={() => onOpen(r.nodeId)} className="mt-2 inline-flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-800">Open it <ArrowRight size={12} /></button>
        </div>
      ); })}</div>
    </Modal>
  );
}

// Work in the viewer's scope that warrants an email nudge: blockers to clear,
// overdue items to chase, and unassigned-but-imminent work to get an owner.
function followUps(works, acts, me) {
  const homeIds = new Set(homeNodes(works, acts, me).flatMap((h) => subtreeIds(works, h.id)));
  const workOf = (id) => works.find((w) => w.id === id);
  // Never draft an email to yourself — escalate your own items to your manager.
  const escalate = (id) => (id && id === me.id ? (me.reports_to || id) : id);
  const list = [];
  acts.filter((a) => homeIds.has(a.workId) && a.status !== "cancelled" && a.status !== "executed").forEach((a) => {
    const w = workOf(a.workId);
    if (a.blocked) list.push({ id: "b-" + a.id, kind: "blocked", act: a, work: w, purpose: "help clear the blocker", situation: "it's blocked and can't move forward", recipientId: escalate(a.assigneeId || (w && w.ownerId)) });
    else if (isOverdue(a)) list.push({ id: "o-" + a.id, kind: "overdue", act: a, work: w, purpose: "get a status update", situation: `it's overdue (was due ${a.date ? fmtFull(parseISO(a.date)) : "recently"})`, recipientId: escalate(a.assigneeId || (w && w.ownerId)) });
    else if (!a.assigneeId && a.date && daysLeft(a.date) != null && daysLeft(a.date) <= 7) list.push({ id: "u-" + a.id, kind: "unassigned", act: a, work: w, purpose: "get it an owner", situation: `it has no owner and is due ${daysLeft(a.date) < 0 ? "already" : "in " + daysLeft(a.date) + "d"}`, recipientId: escalate(w && w.ownerId) });
  });
  const order = { blocked: 0, overdue: 1, unassigned: 2 };
  return list.sort((a, b) => order[a.kind] - order[b.kind]);
}
// The Follow-ups panel: AI-drafts an email per item, editable, then opens it in
// the user's mail app (Outlook) pre-filled so they just review and hit send.
function FollowUps({ me, works, acts, busy, setBusy, flash, onClose, onOpen }) {
  const list = followUps(works, acts, me);
  const [drafts, setDrafts] = useState({});
  const draftEmail = async (fi) => {
    const rec = USERS.find((u) => u.id === fi.recipientId);
    setBusy(fi.id);
    try {
      const d = await AI.draftEmail({ from: `${me.name}, ${me.title}`, to: rec ? `${rec.name}, ${rec.title}` : "the owner", item: fi.act.title, parent: fi.work ? fi.work.title : null, context: fi.situation, purpose: fi.purpose });
      setDrafts((p) => ({ ...p, [fi.id]: { to: rec ? rec.email : "", subject: d.subject || `Following up: ${fi.act.title}`, body: d.body || "" } }));
    } catch { const first = rec ? rec.name.split(" ")[0] : "there"; setDrafts((p) => ({ ...p, [fi.id]: { to: rec ? rec.email : "", subject: `Following up: ${fi.act.title}`, body: `Hi ${first},\n\nCould you help move "${fi.act.title}" forward? ${fi.situation}.\n\nThanks,\n${me.name.split(" ")[0]}` } })); flash("AI unavailable — a simple draft is ready to edit."); }
    setBusy(null);
  };
  const openOutlook = (d) => { const url = `mailto:${encodeURIComponent(d.to)}?subject=${encodeURIComponent(d.subject)}&body=${encodeURIComponent(d.body)}`; const a = document.createElement("a"); a.href = url; a.rel = "noopener"; document.body.appendChild(a); a.click(); a.remove(); };
  const setField = (id, f, v) => setDrafts((p) => ({ ...p, [id]: { ...p[id], [f]: v } }));
  const KIND = { blocked: { label: "Blocked", tone: "text-rose-700", Icon: AlertTriangle }, overdue: { label: "Overdue", tone: "text-rose-700", Icon: Clock }, unassigned: { label: "No owner", tone: "text-amber-700", Icon: User } };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-1 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><Mail size={16} className="text-blue-700" /> Follow-ups</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Work that needs a nudge. Draft an email with AI, tweak it, then open it in your mail app (Outlook) and hit send.</p>
      {list.length === 0 && <div className="rounded-lg border border-dashed border-slate-200 py-8 text-center text-sm text-slate-400">Nothing to chase — no blockers, overdue, or unassigned work in your scope.</div>}
      <div className="max-h-[60vh] space-y-2 overflow-y-auto">
        {list.map((fi) => { const k = KIND[fi.kind]; const rec = USERS.find((u) => u.id === fi.recipientId); const d = drafts[fi.id];
          return (
            <div key={fi.id} className="rounded-lg border border-slate-200">
              <div className="flex items-start gap-2 p-3">
                <k.Icon size={15} className={`mt-0.5 shrink-0 ${k.tone}`} />
                <div className="min-w-0 flex-1">
                  <button onClick={() => onOpen(fi.act.workId)} className="block max-w-full truncate text-left text-sm font-medium text-slate-800 hover:underline">{fi.act.title}</button>
                  <div className="mt-0.5 truncate text-xs text-slate-400">{fi.work ? fi.work.title + " · " : ""}<span className={k.tone}>{k.label}</span> — {fi.situation}</div>
                  <div className="mt-1 flex items-center gap-1.5 text-xs text-slate-500">{rec && <Avatar id={rec.id} size={16} />}Email <span className="font-medium text-slate-700">{rec ? rec.name : "the owner"}</span> to {fi.purpose}.</div>
                </div>
                <button onClick={() => draftEmail(fi)} disabled={busy === fi.id} className={`${btnLight} shrink-0`}>{busy === fi.id ? <><Loader2 size={13} className="animate-spin" /> Drafting…</> : <><Sparkles size={13} /> {d ? "Redraft" : "Draft email"}</>}</button>
              </div>
              {d && <div className="space-y-2 border-t border-slate-100 bg-slate-50 p-3">
                <div className="flex items-center gap-2"><span className="w-14 shrink-0 text-xs text-slate-400">To</span><input value={d.to} onChange={(e) => setField(fi.id, "to", e.target.value)} className={`${inputCls} py-1.5`} /></div>
                <div className="flex items-center gap-2"><span className="w-14 shrink-0 text-xs text-slate-400">Subject</span><input value={d.subject} onChange={(e) => setField(fi.id, "subject", e.target.value)} className={`${inputCls} py-1.5`} /></div>
                <textarea value={d.body} onChange={(e) => setField(fi.id, "body", e.target.value)} rows={6} className={`${inputCls} text-sm`} />
                <div className="flex justify-end"><button onClick={() => openOutlook(d)} disabled={!d.to} className={btnDark}><Send size={14} /> Open in Outlook</button></div>
              </div>}
            </div>
          );
        })}
      </div>
    </Modal>
  );
}
// One universal "add" flow: describe / dictate / attach a doc, AI decides whether
// it's an objective / initiative / work / activity and where it belongs, then a
// level-appropriate form (pre-filled) captures the rest. Parents can be existing
// or created new (one level up).
const ADD_PARENT_LEVEL = { initiative: "objective", work: "initiative", activity: "work" };
function SmartAdd({ me, works, teams, store, busy, setBusy, flash, onClose, onCreated }) {
  const [text, setText] = useState(""); const [pdf, setPdf] = useState(null);
  const [analyzed, setAnalyzed] = useState(false);
  const [level, setLevel] = useState("activity");
  const [title, setTitle] = useState(""); const [description, setDescription] = useState("");
  const [type, setType] = useState("general");
  const [parentId, setParentId] = useState(""); const [newParent, setNewParent] = useState(false);
  const [newParentTitle, setNewParentTitle] = useState(""); const [grandparentId, setGrandparentId] = useState("");
  const [ownerId, setOwnerId] = useState(me.id);
  const [deadline, setDeadline] = useState(""); const [dueDate, setDueDate] = useState(""); const [startDate, setStartDate] = useState("");
  const [hrs, setHrs] = useState(2); const [actType, setActType] = useState("self"); const [teamId, setTeamId] = useState("");
  const [metric, setMetric] = useState(""); const [target, setTarget] = useState(100); const [unit, setUnit] = useState("%");

  // What each role may create: MD owns objectives; VPs run initiatives down; members add work/tasks.
  const allowedLevels = me.level === "md" ? ["objective", "initiative", "work", "activity"] : me.level === "vp" ? ["initiative", "work", "activity"] : ["work", "activity"];
  const parentLevel = ADD_PARENT_LEVEL[level] || null;
  const grandLevel = parentLevel ? ADD_PARENT_LEVEL[parentLevel] || null : null;
  const parentOptions = parentLevel ? works.filter((w) => w.level === parentLevel) : [];
  const grandOptions = grandLevel ? works.filter((w) => w.level === grandLevel) : [];
  const th = LEVEL_THEME[level] || LEVEL_THEME.activity;
  const findUser = (name) => name ? (USERS.find((u) => u.name.toLowerCase() === name.toLowerCase()) || USERS.find((u) => u.name.toLowerCase().includes(name.toLowerCase()))) : null;
  const findTeam = (name) => name ? (teams || []).find((t) => t.name.toLowerCase().includes(name.toLowerCase())) : null;

  const analyze = async () => {
    if (!text.trim()) return;
    setBusy("classify");
    try {
      const ctx = {
        objectives: works.filter((w) => w.level === "objective").map((w) => w.title).join("; ") || "none",
        initiatives: works.filter((w) => w.level === "initiative").map((w) => w.title).join("; ") || "none",
        works: works.filter((w) => w.level === "work").map((w) => w.title).join("; ") || "none",
        people: USERS.map((u) => u.name).join("; "), teams: (teams || []).map((t) => t.name).join("; "), today: iso(TODAY),
      };
      const d = await AI.classify(text, ctx);
      let lvl = ["objective", "initiative", "work", "activity"].includes(d.level) ? d.level : "activity";
      if (!allowedLevels.includes(lvl)) lvl = allowedLevels[0]; // clamp to the highest level this role may create
      setLevel(lvl); setTitle(d.title || ""); setDescription(d.description || ""); if (d.type) setType(d.type);
      if (d.deadline) { setDeadline(d.deadline); setDueDate(d.deadline); }
      if (d.estimateHrs) setHrs(Number(d.estimateHrs) || 2); if (d.actType) setActType(d.actType);
      if (d.metric) { setMetric(d.metric); if (d.unit) setUnit(d.unit); if (d.target) setTarget(Number(d.target) || 100); }
      const pl = ADD_PARENT_LEVEL[lvl];
      if (pl && d.parentTitle) { const opts = works.filter((w) => w.level === pl); const n = d.parentTitle.toLowerCase(); const p = opts.find((w) => w.title.toLowerCase() === n) || opts.find((w) => w.title.toLowerCase().includes(n)); if (p) { setParentId(p.id); setNewParent(false); } else { setNewParent(true); setNewParentTitle(d.parentTitle); } }
      const u = findUser(d.assignee); if (u) setOwnerId(u.id); const t = findTeam(d.team); if (t) setTeamId(t.id);
      flash("AI sorted it — review the fields and create.");
    } catch { const t = (text.split("\n").find((l) => l.trim()) || "").slice(0, 90).trim(); if (t) setTitle(t); flash("AI unavailable — pick the type and fill the fields."); }
    setAnalyzed(true); setBusy(null);
  };

  const parentReady = !parentLevel || (newParent ? newParentTitle.trim() && (!grandLevel || grandparentId) : parentId);
  const canCreate = title.trim() && parentReady;
  const create = () => {
    if (!canCreate) return;
    let pid = null;
    if (parentLevel) {
      if (newParent) {
        pid = nid("w");
        const tpl = METRIC_BY_TYPE[type] || METRIC_BY_TYPE.general;
        const pNode = { id: pid, parentId: grandLevel ? (grandparentId || null) : null, level: parentLevel, title: newParentTitle.trim(), ownerId: me.id, type };
        if (parentLevel === "initiative") pNode.result = { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 };
        store.addWorks([pNode]);
      } else pid = parentId;
    }
    if (level === "activity") {
      const start = startDate || (dueDate ? iso(addDays(parseISO(dueDate), -Math.max(1, Math.round((Number(hrs) || 2) / 3)))) : null);
      store.addActs([{ id: nid("a"), workId: pid, title: title.trim(), description: description || undefined, assigneeId: ownerId || null, startDate: start, date: dueDate || null, status: "planned", plannedHrs: Number(hrs) || 2, actualHrs: null, actType, unplanned: true }]);
      flash("Activity added."); onCreated && onCreated(pid); onClose(); return;
    }
    const tpl = METRIC_BY_TYPE[type] || METRIC_BY_TYPE.general;
    const node = { id: nid("w"), parentId: level === "objective" ? null : pid, level, title: title.trim(), ownerId: ownerId || me.id, type };
    if (description) node.description = description;
    if (deadline) node.deadline = deadline;
    if (level === "initiative") { if (teamId) { node.teamId = teamId; node.scope = "group"; } node.result = metric.trim() ? { metric: metric.trim(), unit, baseline: 0, target: Number(target) || 100, current: 0 } : { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 }; }
    if (level === "objective" && metric.trim()) node.result = { metric: metric.trim(), unit, baseline: 0, target: Number(target) || 100, current: 0 };
    store.addWorks([node]);
    flash(`${LEVEL_LABEL[level]} created.`); onCreated && onCreated(node.id); onClose();
  };

  const LEVELS = [["objective", "Objective"], ["initiative", "Initiative"], ["work", "Work"], ["activity", "Task"]].filter(([k]) => allowedLevels.includes(k));
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-1 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><Sparkles size={16} className="text-blue-700" /> Add anything</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Describe it, dictate, or attach a doc — AI decides whether it's an objective, initiative, work, or task, and where it belongs. Then review the fields and create.</p>

      <div className="mb-3 rounded-lg border border-blue-200 bg-blue-50 p-3">
        <MultiModalInput value={text} onChange={setText} onPdf={setPdf} placeholder="e.g. 'Get 3 quotes for the new forklifts by next Friday, assign to Neha' — or a whole objective…" />
        <div className="mt-2 flex items-center gap-2">
          <button onClick={analyze} disabled={busy || !text.trim()} className={`${btnViolet} flex-1`}>{busy === "classify" ? <><Loader2 size={15} className="animate-spin" /> Thinking…</> : <><Sparkles size={15} /> Analyze with AI</>}</button>
          {!analyzed && <button onClick={() => setAnalyzed(true)} className="shrink-0 text-xs text-slate-400 hover:text-slate-700">or add manually</button>}
        </div>
      </div>

      {analyzed && <div className="rounded-lg border border-slate-200 p-3">
        <label className="mb-1.5 block text-xs font-medium text-slate-600">This is a…</label>
        <div className="mb-3 flex flex-wrap gap-1.5">{LEVELS.map(([k, l]) => { const t = LEVEL_THEME[k]; const on = level === k; return <button key={k} onClick={() => { setLevel(k); setNewParent(false); setParentId(""); }} className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-xs font-medium ${on ? `${t.chip} border-transparent ring-2 ring-offset-1 ${t.ring}` : "border-slate-200 text-slate-500 hover:bg-slate-50"}`}><span className={`h-1.5 w-1.5 rounded-full ${t.bar}`} />{l}</button>; })}</div>

        <label className="mb-1 block text-xs font-medium text-slate-600">{level === "activity" ? "Task" : LEVEL_LABEL[level]} name</label>
        <input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder="Name it…" />

        {parentLevel && <div className="mb-3">
          <label className="mb-1 block text-xs font-medium text-slate-600">Lives under ({LEVEL_LABEL[parentLevel]})</label>
          {!newParent ? <div className="flex items-center gap-2">
            <select value={parentId} onChange={(e) => setParentId(e.target.value)} className={inputCls}><option value="">Pick an existing {LEVEL_LABEL[parentLevel].toLowerCase()}…</option>{parentOptions.map((p) => <option key={p.id} value={p.id}>{p.title}</option>)}</select>
            <button onClick={() => setNewParent(true)} className="shrink-0 whitespace-nowrap text-xs font-medium text-blue-700 hover:text-blue-900">+ new</button>
          </div> : <div className="space-y-2 rounded-md border border-blue-200 bg-blue-50 p-2">
            <div className="flex items-center gap-2"><input value={newParentTitle} onChange={(e) => setNewParentTitle(e.target.value)} className={inputCls} placeholder={`New ${LEVEL_LABEL[parentLevel].toLowerCase()} name…`} /><button onClick={() => setNewParent(false)} className="shrink-0 text-xs text-slate-500 hover:text-slate-800">use existing</button></div>
            {grandLevel && <select value={grandparentId} onChange={(e) => setGrandparentId(e.target.value)} className={inputCls}><option value="">Under which {LEVEL_LABEL[grandLevel].toLowerCase()}?</option>{grandOptions.map((g) => <option key={g.id} value={g.id}>{g.title}</option>)}</select>}
          </div>}
        </div>}

        <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
          {level !== "activity" && <div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{Object.keys(METRIC_BY_TYPE).map((t) => <option key={t} value={t}>{t}</option>)}</select></div>}
          <div><label className="mb-1 block text-xs font-medium text-slate-600">{level === "activity" ? "Assignee" : "Owner"}</label><select value={ownerId} onChange={(e) => setOwnerId(e.target.value)} className={inputCls}><option value="">Unassigned</option>{USERS.map((u) => <option key={u.id} value={u.id}>{u.name}</option>)}</select></div>
          {level === "initiative" && <div><label className="mb-1 block text-xs font-medium text-slate-600">Team</label><select value={teamId} onChange={(e) => setTeamId(e.target.value)} className={inputCls}><option value="">No team</option>{(teams || []).map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}</select></div>}
          {level !== "activity" && <div><label className="mb-1 block text-xs font-medium text-slate-600">Deadline</label><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className={inputCls} /></div>}
        </div>

        {level === "activity" && <div className="mb-3 grid grid-cols-2 gap-3 sm:grid-cols-4">
          <div><label className="mb-1 block text-xs font-medium text-slate-600">Start</label><input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className={inputCls} /></div>
          <div><label className="mb-1 block text-xs font-medium text-slate-600">Due</label><input type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} className={inputCls} /></div>
          <div><label className="mb-1 block text-xs font-medium text-slate-600">Est. hrs</label><input type="number" value={hrs} onChange={(e) => setHrs(e.target.value)} className={inputCls} /></div>
          <div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={actType} onChange={(e) => setActType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
        </div>}

        {(level === "objective" || level === "initiative") && <div className="mb-3">
          <label className="mb-1 block text-xs font-medium text-slate-600">Target metric{level === "objective" ? " (optional)" : ""}</label>
          <div className="grid grid-cols-1 gap-2 text-xs sm:grid-cols-3"><input value={metric} onChange={(e) => setMetric(e.target.value)} className={inputCls} placeholder="e.g. Cost / tonne" /><input type="number" value={target} onChange={(e) => setTarget(e.target.value)} className={inputCls} placeholder="Target" /><input value={unit} onChange={(e) => setUnit(e.target.value)} className={inputCls} placeholder="Unit" /></div>
        </div>}

        {level === "activity" && <div className="mb-3"><label className="mb-1 block text-xs font-medium text-slate-600">Description (what to produce)</label><textarea value={description} onChange={(e) => setDescription(e.target.value)} rows={2} className={inputCls} placeholder="Optional — used when the AI scores the deliverable." /></div>}

        <button onClick={create} disabled={!canCreate} className={`${btnDark} w-full`}><span className={`h-2 w-2 rounded-full ${th.bar}`} /> Create {level === "activity" ? "task" : LEVEL_LABEL[level].toLowerCase()}</button>
      </div>}
    </Modal>
  );
}
function QuickCreate({ me, parent, level, works, store, busy, setBusy, flash, onClose, onCreated }) {
  const isObj = level === "objective", isAct = level === "activity", isContainer = level === "initiative" || level === "work";
  const label = isAct ? "Task" : (LEVEL_LABEL[level] || "Item");
  const th = LEVEL_THEME[level] || LEVEL_THEME.activity;
  const [title, setTitle] = useState("");
  const [type, setType] = useState((parent && parent.type) || "general");
  const [deadline, setDeadline] = useState("");
  const [hasMetric, setHasMetric] = useState(false); const [metric, setMetric] = useState(""); const [target, setTarget] = useState(100); const [unit, setUnit] = useState("%");
  const [hrs, setHrs] = useState(2); const [actType, setActType] = useState("self"); const [date, setDate] = useState("");
  const [text, setText] = useState(""); const [pdf, setPdf] = useState(null); const [subs, setSubs] = useState([]);
  const draftIt = async () => {
    setBusy("draft");
    try {
      let sys, usr;
      if (isObj) { sys = 'Turn the note into ONE crisp enterprise objective (an outcome, not a task). Return ONLY JSON: {"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","metric":string,"unit":string,"target":number}'; usr = `Note:\n"""${text}"""`; }
      else if (level === "initiative") { sys = 'Draft an initiative that fulfils the objective, broken into 3-6 works (phases of execution), each with 1-4 concrete activities. Return ONLY JSON: {"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","works":[{"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}'; usr = `Objective: "${parent.title}". Note:\n"""${text}"""`; }
      else if (level === "work") { sys = 'Draft a work (a phase of execution) with 1-5 concrete activities. Return ONLY JSON: {"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}'; usr = `Initiative: "${parent.title}". Note:\n"""${text}"""`; }
      else { sys = 'Turn the note into ONE task. Return ONLY JSON: {"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}'; usr = `Work: "${parent.title}". Note:\n"""${text}"""`; }
      const d = parseJSON(await aiComplete(sys, usr, pdf ? pdf.data : undefined));
      if (d.title) setTitle(d.title);
      if (d.type && (isObj || isContainer)) setType(d.type);
      if (isObj && d.metric) { setHasMetric(true); setMetric(d.metric); setUnit(d.unit || "%"); setTarget(d.target || 100); }
      if (isAct) { if (d.estimateHrs) setHrs(Number(d.estimateHrs) || 2); if (d.type) setActType(d.type); }
      if (level === "initiative") setSubs(d.works || []);
      if (level === "work") setSubs(d.activities || []);
      flash("AI drafted it — review the fields, then create.");
    } catch { const t = (text.split("\n").find((l) => l.trim()) || "").slice(0, 90).trim(); if (t) setTitle(t); flash("AI unavailable — fill the fields in and create."); }
    setBusy(null);
  };
  const create = () => {
    if (!title.trim()) return;
    if (isAct) { store.addActs([{ id: nid("a"), workId: parent.id, title: title.trim(), assigneeId: null, date: date || null, status: "planned", plannedHrs: Number(hrs) || 2, actualHrs: null, actType, unplanned: true }]); flash("Task added."); onCreated && onCreated(parent.id); onClose(); return; }
    const nw = []; const na = []; const topId = nid("w"); const tpl = METRIC_BY_TYPE[type] || METRIC_BY_TYPE.general;
    const node = { id: topId, parentId: parent ? parent.id : null, level, title: title.trim(), ownerId: me.id, type };
    if (deadline) node.deadline = deadline;
    if (isObj && hasMetric && metric.trim()) node.result = { metric: metric.trim(), unit, baseline: 0, target: Number(target) || 100, current: 0 };
    if (level === "initiative") node.result = { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 };
    nw.push(node);
    if (level === "work") { subs.forEach((ac) => na.push({ id: nid("a"), workId: topId, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }
    else if (level === "initiative") { subs.forEach((wk) => { const wid = nid("w"); nw.push({ id: wid, parentId: topId, level: "work", title: wk.title, ownerId: me.id, type }); (wk.activities || []).forEach((ac) => na.push({ id: nid("a"), workId: wid, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }); }
    store.addWorks(nw); if (na.length) store.addActs(na);
    flash(`${label} created.`); onCreated && onCreated(topId); onClose();
  };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-1 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><span className={`h-2.5 w-2.5 rounded-full ${th.bar}`} /> New {label.toLowerCase()}</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">{parent ? <>Under <span className="font-medium text-slate-500">{parent.title}</span>. </> : null}Fill the fields, or describe / dictate / attach a PDF below and let AI fill them.</p>

      <label className="mb-1 block text-xs font-medium text-slate-600">{label} name</label>
      <input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder={isObj ? "e.g. Become the lowest-cost clinker producer in the region" : isAct ? "e.g. Call the vendor to confirm delivery dates" : `Name this ${label.toLowerCase()}`} />

      {(isObj || isContainer) && <div className="mb-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{Object.keys(METRIC_BY_TYPE).map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Deadline</label><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className={inputCls} /></div>
      </div>}

      {isObj && <div className="mb-3">
        <label className="mb-1 flex items-center gap-2 text-xs font-medium text-slate-600"><input type="checkbox" checked={hasMetric} onChange={(e) => setHasMetric(e.target.checked)} /> North-star metric</label>
        {hasMetric && <div className="grid grid-cols-1 gap-2 text-xs sm:grid-cols-3"><div><label className="mb-1 block text-slate-500">Metric</label><input value={metric} onChange={(e) => setMetric(e.target.value)} className={inputCls} placeholder="Cost / tonne" /></div><div><label className="mb-1 block text-slate-500">Target</label><input type="number" value={target} onChange={(e) => setTarget(e.target.value)} className={inputCls} /></div><div><label className="mb-1 block text-slate-500">Unit</label><input value={unit} onChange={(e) => setUnit(e.target.value)} className={inputCls} /></div></div>}
      </div>}

      {isAct && <div className="mb-3 grid grid-cols-1 gap-2 sm:grid-cols-3"><div><label className="mb-1 block text-xs font-medium text-slate-600">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={actType} onChange={(e) => setActType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div><div><label className="mb-1 block text-xs font-medium text-slate-600">Date</label><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} /></div></div>}

      <div className="mb-3 rounded-lg border border-blue-200 bg-blue-50 p-3">
        <div className="mb-2 flex items-center gap-1.5 text-xs font-medium text-blue-900"><Sparkles size={14} /> Describe it, dictate, or attach a PDF — AI fills the fields{isContainer ? " and a starter breakdown" : ""}</div>
        <MultiModalInput value={text} onChange={setText} onPdf={setPdf} placeholder="Speak or paste the note, or attach the doc…" />
        <button onClick={draftIt} disabled={busy || (!text.trim() && !pdf)} className={`${btnViolet} mt-2 w-full`}>{busy === "draft" ? <><Loader2 size={15} className="animate-spin" /> Reading…</> : <><Sparkles size={15} /> Draft with AI</>}</button>
      </div>

      {subs.length > 0 && <div className="mb-3"><div className="mb-1 text-xs font-medium text-slate-500">{level === "initiative" ? "Works" : "Activities"} AI drafted ({subs.length})</div><div className="max-h-40 space-y-1 overflow-y-auto">{subs.map((c, i) => <div key={i} className="flex items-center gap-2 rounded-md bg-slate-50 px-3 py-1.5 text-sm text-slate-700"><span className="min-w-0 flex-1 truncate">{c.title}</span>{c.activities ? <span className="shrink-0 text-xs text-slate-400">{c.activities.length} activities</span> : null}<button onClick={() => setSubs(subs.filter((_, j) => j !== i))} className="shrink-0 text-slate-300 hover:text-rose-500"><X size={13} /></button></div>)}</div></div>}

      <button onClick={create} disabled={!title.trim()} className={`${btnDark} w-full`}><Check size={14} /> Create {label.toLowerCase()}</button>
    </Modal>
  );
}

function ObjectiveModal({ me, setWorks, flash, onClose, onOpen }) {
  const [title, setTitle] = useState(""); const [hasMetric, setHasMetric] = useState(false); const [metric, setMetric] = useState(""); const [target, setTarget] = useState(100); const [unit, setUnit] = useState("%");
  const create = () => { if (!title.trim()) return; const id = nid("w"); const node = { id, parentId: null, level: "objective", title: title.trim(), type: "general", ownerId: me.id }; if (hasMetric && metric.trim()) node.result = { metric: metric.trim(), unit, baseline: 0, target: Number(target) || 100, current: 0 }; setWorks((p) => [...p, node]); flash("Objective created — now add initiatives under it."); onOpen(id); onClose(); };
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">New objective</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">A top-level outcome you own. Your VPs create the initiatives that fulfil it.</p>
      <label className="mb-1 block text-xs text-slate-500">Objective</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder="e.g. Become the lowest-cost producer in the region" />
      <label className="mb-2 flex items-center gap-2 text-xs text-slate-500"><input type="checkbox" checked={hasMetric} onChange={(e) => setHasMetric(e.target.checked)} /> Track a north-star metric</label>
      {hasMetric && <div className="mb-3 grid grid-cols-3 gap-2 text-xs"><div><label className="mb-1 block text-slate-500">Metric</label><input value={metric} onChange={(e) => setMetric(e.target.value)} className={inputCls} placeholder="Cost / tonne" /></div><div><label className="mb-1 block text-slate-500">Target</label><input type="number" value={target} onChange={(e) => setTarget(e.target.value)} className={inputCls} /></div><div><label className="mb-1 block text-slate-500">Unit</label><input value={unit} onChange={(e) => setUnit(e.target.value)} className={inputCls} /></div></div>}
      <button onClick={create} disabled={!title.trim()} className={`${btnDark} w-full`}>Create objective</button>
    </Modal>
  );
}

/* ---------- Modify / Unplanned / Propose / Deliverable ---------- */
function ModifyPlan({ work, planText, subs, acts, store, busy, setBusy, flash, onClose }) {
  const [text, setText] = useState(""); const [ops, setOps] = useState(null);
  const run = async () => { setBusy("mod"); try { setOps((await AI.modifyPlan(planText, text)).ops || []); } catch { flash("AI unavailable."); } setBusy(null); };
  const apply = () => {
    const nw = []; (ops || []).forEach((op) => { if (op.op === "add_work" || op.op === "add_subwork") nw.push({ id: nid("w"), parentId: work.id, level: "work", title: op.title, type: work.type, ownerId: work.ownerId }); });
    const liveSubs = subs.concat(nw); const na = [];
    (ops || []).forEach((op) => { if (op.op === "add_activity") { const key = (op.work || op.subwork || "").toLowerCase(); const sw = liveSubs.find((s) => s.title.toLowerCase().includes(key)) || liveSubs[0]; if (sw) na.push({ id: nid("a"), workId: sw.id, title: op.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(op.estimateHrs) || 2, actualHrs: null, actType: op.type || "self", unplanned: true }); } });
    // retype: match by activity title within the current sub-works
    const subIds = new Set(subs.map((s) => s.id));
    const actPatches = {};
    (ops || []).forEach((op) => { if (op.op === "retype") (acts || []).forEach((a) => { if (subIds.has(a.workId) && a.title.toLowerCase().includes((op.match || "").toLowerCase())) actPatches[a.id] = { actType: op.type }; }); });
    if (nw.length) store.addWorks(nw);
    if (na.length) store.addActs(na);
    if (Object.keys(actPatches).length) store.patchActs(actPatches);
    flash("Plan updated by AI."); onClose();
  };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Modify the plan</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Tell AI what changed — type, speak, or paste a note.</p>
      <MultiModalInput value={text} onChange={setText} placeholder="e.g. Legal wants a data-privacy review before PO; HOD sign-off is a meeting not a self task." />
      {!ops && <button onClick={run} disabled={busy || !text.trim()} className={`${btnViolet} mt-3 w-full`}>{busy ? <><Loader2 size={15} className="animate-spin" /> Reading…</> : <><Sparkles size={15} /> Propose changes</>}</button>}
      {ops && <div className="mt-3"><div className="mb-2 text-xs font-medium text-slate-500">Proposed changes</div><div className="space-y-1">{ops.length === 0 && <div className="text-sm text-slate-400">No changes proposed.</div>}{ops.map((op, i) => <div key={i} className="rounded-md bg-slate-50 px-3 py-1.5 text-sm text-slate-700">{op.op === "add_activity" ? `Add “${op.title}” to ${op.work || op.subwork || "a work"}` : (op.op === "add_work" || op.op === "add_subwork") ? `Add work “${op.title}”` : `Retype “${op.match}” → ${op.type}`}</div>)}</div><div className="mt-3 flex gap-2"><button onClick={apply} className={`${btnDark} flex-1`}><Check size={14} /> Apply</button><button onClick={() => setOps(null)} className={btnLight}>Redo</button></div></div>}
    </Modal>
  );
}
function AddUnplanned({ subs, store, flash, onClose }) {
  const [title, setTitle] = useState(""); const [sw, setSw] = useState(subs[0] ? subs[0].id : ""); const [hrs, setHrs] = useState(2); const [type, setType] = useState("self");
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Add unplanned activity</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Something that came up mid-flight and changes the plan.</p>
      <label className="mb-1 block text-xs text-slate-500">What needs doing</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder="e.g. Emergency security patch review" />
      <div className="mb-4 grid grid-cols-1 gap-2 text-xs sm:grid-cols-3"><div><label className="mb-1 block text-slate-500">Under</label><select value={sw} onChange={(e) => setSw(e.target.value)} className={inputCls}>{subs.map((s) => <option key={s.id} value={s.id}>{s.title.slice(0, 18)}</option>)}</select></div><div><label className="mb-1 block text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
      <button onClick={() => { if (!title.trim() || !sw) return; store.addActs([{ id: nid("a"), workId: sw, title: title.trim(), assigneeId: null, date: null, status: "planned", plannedHrs: hrs, actualHrs: null, actType: type, unplanned: true }]); flash("Unplanned activity added."); onClose(); }} disabled={!title.trim()} className={`${btnDark} w-full`}>Add to plan</button>
    </Modal>
  );
}
function ProposeChange({ activity, me, store, flash, onClose }) {
  const TPL = { add_activity: "A follow-up task is needed here: ", extend: "This is taking longer than the estimate because ", blocked: "I'm blocked on this and need help — ", reassign: "This should move to someone else because " };
  const [kind, setKind] = useState("add_activity"); const [desc, setDesc] = useState(TPL.add_activity);
  const [title, setTitle] = useState(""); const [hrs, setHrs] = useState(2); const [to, setTo] = useState("");
  const pick = (k) => { setKind(k); setDesc(TPL[k]); };
  const submit = () => {
    const payload = kind === "add_activity" ? { title: title || "Follow-up task", hrs, type: "self" } : kind === "extend" ? { activityId: activity.id, hrs } : kind === "reassign" ? { activityId: activity.id, to } : { activityId: activity.id };
    store.addCr({ id: nid("cr"), workId: activity.workId, targetWorkId: activity.workId, proposerId: me.id, kind, desc: desc || "(no note)", status: "pending", payload });
    flash("Proposed. It now needs organizer approval."); onClose();
  };
  const kinds = [["add_activity", "Add follow-up"], ["extend", "Needs more time"], ["blocked", "Blocked / help"], ["reassign", "Change owner"]];
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Propose a plan change</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">On “{activity.title}”. Your organizer approves before it changes the plan.</p>
      <div className="mb-3 flex flex-wrap gap-1">{kinds.map(([k, l]) => <button key={k} onClick={() => pick(k)} className={`rounded-md px-2.5 py-1.5 text-xs font-medium ${kind === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>{l}</button>)}</div>
      {kind === "add_activity" && <div className="mb-3 grid grid-cols-1 gap-2 text-xs sm:grid-cols-3"><div className="sm:col-span-2"><label className="mb-1 block text-slate-500">New activity</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={inputCls} placeholder="e.g. Re-run for missed depots" /></div><div><label className="mb-1 block text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div></div>}
      {kind === "extend" && <div className="mb-3 text-xs"><label className="mb-1 block text-slate-500">Extra hours needed</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={`${inputCls} w-28`} /></div>}
      {kind === "reassign" && <div className="mb-3 text-xs"><label className="mb-1 block text-slate-500">Hand over to</label><select value={to} onChange={(e) => setTo(e.target.value)} className={inputCls}><option value="">Pick a person</option>{USERS.filter((u) => u.id !== me.id).map((u) => <option key={u.id} value={u.id}>{u.name} — {u.title}</option>)}</select></div>}
      {kind === "blocked" && <div className="mb-3 rounded-md bg-rose-50 px-3 py-2 text-xs text-rose-700">Flags the activity as blocked so your organizer sees it in “Needs attention”.</div>}
      <label className="mb-1 block text-xs text-slate-500">Why</label><textarea value={desc} onChange={(e) => setDesc(e.target.value)} rows={2} className={`${inputCls} mb-3`} placeholder="Reason for the change…" />
      <button onClick={submit} disabled={kind === "reassign" && !to} className={`${btnDark} w-full`}>Submit for approval</button>
    </Modal>
  );
}
// Deliverable kinds — icon + label + the file extensions that make sense for each.
const DELIV_KIND = {
  document: { label: "Document", Icon: FileText },
  spreadsheet: { label: "Spreadsheet", Icon: FileSpreadsheet },
  email: { label: "Email", Icon: Mail },
  slides: { label: "Slides", Icon: Presentation },
  other: { label: "File", Icon: Paperclip },
};
const DELIV_ORDER = ["document", "spreadsheet", "email", "slides", "other"];
const DELIV_ACCEPT = ".pdf,.docx,.doc,.xlsx,.xls,.csv,.eml,.msg,.txt,.md";

// Attach a file to ONE checklist item and (optionally) AI-score it against what
// the item asks for. Writes the updated item back into the work's deliverables.
function DeliverableItemModal({ work, item, store, busy, setBusy, flash, onClose }) {
  const [name, setName] = useState(item.file?.name || "");
  const [content, setContent] = useState("");
  const [reading, setReading] = useState(false);
  const [result, setResult] = useState(item.score != null ? { score: item.score, verdict: item.verdict, feedback: item.feedback } : null);
  const K = DELIV_KIND[item.kind] || DELIV_KIND.other;
  const readF = async (f) => {
    if (!f) return; setName(f.name); const n = f.name.toLowerCase();
    if (n.endsWith(".txt") || n.endsWith(".md") || f.type.startsWith("text/plain")) { const r = new FileReader(); r.onload = () => setContent(String(r.result || "").slice(0, 8000)); r.readAsText(f); }
    else { setReading(true); try { const b64 = await fileToB64(f); const text = await api.aiExtract(b64, f.name); setContent(text || ""); } catch (e) { flash(e.message || "Couldn't read that file."); } setReading(false); }
  };
  const pickOD = async () => {
    try {
      const picked = await pickOneDriveFile();
      if (!picked) return;
      setName(picked.name); setReading(true);
      const text = await api.aiExtract(picked.dataB64, picked.name);
      setContent(text || "");
    } catch (e) { flash(e.message || "Couldn't read that file from OneDrive."); }
    setReading(false);
  };
  const persist = (patch) => { const list = (work.deliverables || []).map((d) => (d.id === item.id ? { ...d, ...patch } : d)); store.patchWork(work.id, { deliverables: list }); };
  const saveDelivered = () => { persist({ done: true, doneAt: iso(TODAY), file: name ? { name } : item.file }); flash("Marked delivered."); onClose(); };
  const score = async () => {
    setBusy("score"); let out; const material = content.trim() || `Submitted file: ${name || item.label}`;
    try { out = await AI.score(work.title, item.label, item.label, material); }
    catch { out = { score: content.length > 200 ? 72 : 55, verdict: "Reasonable draft", feedback: "AI scoring unavailable; provisional score." }; }
    setResult(out);
    persist({ done: true, doneAt: iso(TODAY), file: name ? { name } : item.file, score: out.score, verdict: out.verdict, feedback: out.feedback });
    setBusy(null); flash(`Deliverable scored ${out.score}/100.`);
  };
  return (
    <Modal onClose={onClose}>
      <div className="mb-1 flex items-center justify-between"><h3 className="inline-flex items-center gap-2 text-sm font-medium text-slate-900"><K.Icon size={15} className="text-slate-400" /> {item.label}</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 text-xs text-slate-400">{K.label} deliverable for “{work.title}”. Attach the output and score it, or just mark it delivered.</div>
      <div className="mb-2 flex flex-col gap-1.5 sm:flex-row">
        <label className="flex flex-1 cursor-pointer items-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50"><Upload size={14} /> {reading ? <span className="inline-flex items-center gap-1"><Loader2 size={13} className="animate-spin" /> Reading {name}…</span> : (name || "Choose a file — PDF, Word, Excel/CSV, email (.eml/.msg), .txt/.md")}<input type="file" accept={DELIV_ACCEPT} className="hidden" onChange={(e) => readF(e.target.files && e.target.files[0])} /></label>
        {MSAL_CONFIGURED && <button onClick={pickOD} disabled={reading} className="flex flex-1 items-center justify-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50 disabled:opacity-50"><Cloud size={14} /> From OneDrive</button>}
      </div>
      <label className="mb-1 block text-xs text-slate-500">Content / summary <span className="text-slate-400">(auto-filled from the file; edit or paste your own)</span></label>
      <textarea value={content} onChange={(e) => setContent(e.target.value)} rows={4} className={`${inputCls} mb-3`} placeholder="Paste the deliverable text or a summary…" />
      <div className="flex flex-col gap-2 sm:flex-row">
        <button onClick={score} disabled={busy || (!content.trim() && !name)} className={`${btnDark} flex-1`}>{busy === "score" ? <><Loader2 size={15} className="animate-spin" /> Scoring…</> : <><Sparkles size={15} /> Score with AI</>}</button>
        <button onClick={saveDelivered} disabled={busy} className={`${btnLight} flex-1`}><Check size={15} /> Mark delivered</button>
      </div>
      {result && <div className="mt-4 rounded-lg border border-slate-200 bg-slate-50 p-3"><div className="flex items-center gap-2"><div className="font-mono text-2xl font-medium text-blue-800">{result.score}<span className="text-sm text-slate-400">/100</span></div><div className="text-sm font-medium text-slate-700">{result.verdict}</div></div><div className="mt-1 text-sm text-slate-600">{result.feedback}</div></div>}
    </Modal>
  );
}

// Work-level deliverables: a checklist of the outputs a work must produce. Owners
// tick items off (attaching + scoring optionally); once every item is delivered
// the work — and all its activities — can be marked complete.
function WorkDeliverables({ work, acts, store, canManage, busy, setBusy, flash }) {
  const items = work.deliverables || [];
  const [modalItem, setModalItem] = useState(null);
  const [adding, setAdding] = useState(false);
  const [newLabel, setNewLabel] = useState(""); const [newKind, setNewKind] = useState("document");
  const [suggesting, setSuggesting] = useState(false);
  const [confirmComplete, setConfirmComplete] = useState(false);
  const nodeActs = acts.filter((a) => a.workId === work.id && a.status !== "cancelled");
  const done = items.filter((d) => d.done).length;
  const scored = items.filter((d) => d.done && typeof d.score === "number");
  const avg = scored.length ? Math.round(scored.reduce((s, d) => s + d.score, 0) / scored.length) : null;
  const allDone = items.length > 0 && done === items.length;
  const openActs = nodeActs.filter((a) => a.status !== "executed").length;
  const blank = () => ({ id: nid("d"), done: false, doneAt: null, file: null, score: null, verdict: null, feedback: null });

  const patchList = (list, extra = {}) => store.patchWork(work.id, { deliverables: list, ...extra });
  const toggle = (id) => {
    const list = items.map((d) => (d.id === id ? { ...d, done: !d.done, doneAt: !d.done ? iso(TODAY) : null } : d));
    // reopening an item un-completes the work
    patchList(list, work.completedAt && list.some((d) => !d.done) ? { completedAt: null } : {});
  };
  const remove = (id) => patchList(items.filter((d) => d.id !== id));
  const add = () => { if (!newLabel.trim()) return; patchList([...items, { ...blank(), label: newLabel.trim(), kind: newKind }]); setNewLabel(""); setAdding(false); };
  const suggest = async () => {
    setSuggesting(true);
    try {
      const r = await AI.suggestDeliverables(work.title, nodeActs.map((a) => a.title));
      const have = new Set(items.map((d) => d.label.toLowerCase()));
      const fresh = (r.deliverables || []).filter((d) => d.label && !have.has(d.label.toLowerCase())).map((d) => ({ ...blank(), label: d.label, kind: DELIV_KIND[d.kind] ? d.kind : "other" }));
      if (fresh.length) { patchList([...items, ...fresh]); flash(`Added ${fresh.length} suggested deliverable${fresh.length > 1 ? "s" : ""}.`); }
      else flash("Nothing new to suggest — your checklist looks complete.");
    } catch { flash("AI unavailable."); }
    setSuggesting(false);
  };
  const markComplete = () => {
    const up = {}; nodeActs.forEach((a) => { if (a.status !== "executed") up[a.id] = { status: "executed", actualHrs: a.actualHrs != null ? a.actualHrs : a.plannedHrs, inProgress: false, onHold: false }; });
    if (Object.keys(up).length) store.patchActs(up);
    store.patchWork(work.id, { completedAt: iso(TODAY) });
    setConfirmComplete(false); flash("Work marked complete — all activities executed.");
  };

  return (
    <div className="mt-3 rounded-xl border border-slate-200 bg-white p-4">
      <div className="mb-3 flex flex-wrap items-center gap-2">
        <span className="inline-flex items-center gap-1.5 text-sm font-medium text-slate-700"><Package size={15} className="text-slate-400" /> Deliverables</span>
        {items.length > 0 && <span className="text-xs text-slate-400">{done}/{items.length} delivered</span>}
        {avg != null && <Chip tone="blue"><Star size={10} className="text-amber-500" /> avg {avg}/100</Chip>}
        {work.completedAt && <Chip tone="emerald"><CheckCircle2 size={10} /> completed {fmtShort(work.completedAt)}</Chip>}
        <div className="ml-auto flex items-center gap-1.5">
          {canManage && <button onClick={suggest} disabled={suggesting} className={btnLight} title="Let AI propose the outputs this work should produce">{suggesting ? <Loader2 size={13} className="animate-spin" /> : <Sparkles size={13} />} Suggest</button>}
          {canManage && <button onClick={() => setAdding((v) => !v)} className={btnLight}><Plus size={13} /> Add</button>}
        </div>
      </div>

      {canManage && adding && (
        <div className="mb-3 rounded-lg border border-slate-200 bg-slate-50 p-2.5">
          <input autoFocus value={newLabel} onChange={(e) => setNewLabel(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} placeholder="What output? e.g. Vendor comparison sheet" className={`${inputCls} mb-2`} />
          <div className="flex flex-wrap items-center gap-1.5">
            {DELIV_ORDER.map((k) => { const K = DELIV_KIND[k]; return <button key={k} onClick={() => setNewKind(k)} className={`inline-flex items-center gap-1 rounded-md border px-2 py-1 text-xs ${newKind === k ? "border-blue-300 bg-blue-50 text-blue-800" : "border-slate-200 text-slate-500 hover:bg-white"}`}><K.Icon size={12} /> {K.label}</button>; })}
            <button onClick={add} disabled={!newLabel.trim()} className={`${btnDark} ml-auto`}>Add</button>
          </div>
        </div>
      )}

      {items.length === 0 ? (
        <div className="rounded-lg border border-dashed border-slate-200 px-4 py-5 text-center text-sm text-slate-400">No deliverables yet — add the outputs this work should produce{canManage ? ", or tap Suggest to let AI propose them." : "."}</div>
      ) : (
        <div className="space-y-1.5">
          {items.map((d) => { const K = DELIV_KIND[d.kind] || DELIV_KIND.other; return (
            <div key={d.id} className={`flex items-center gap-2.5 rounded-lg border px-3 py-2 ${d.done ? "border-emerald-200 bg-emerald-50/40" : "border-slate-200 bg-white"}`}>
              <button onClick={() => canManage && toggle(d.id)} disabled={!canManage} className={`shrink-0 ${d.done ? "text-emerald-600" : "text-slate-300 hover:text-slate-400"} disabled:cursor-default`} title={d.done ? "Delivered — click to reopen" : "Mark delivered"}>{d.done ? <CheckCircle2 size={18} /> : <Circle size={18} />}</button>
              <K.Icon size={14} className="shrink-0 text-slate-400" />
              <div className="min-w-0 flex-1">
                <div className={`truncate text-sm ${d.done ? "text-slate-700" : "text-slate-800"}`}>{d.label}</div>
                {(d.file || d.verdict) && <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-slate-400">{d.file && <span className="inline-flex items-center gap-1 truncate"><Paperclip size={10} /> {d.file.name}</span>}{d.verdict && <span className="text-slate-500">{d.verdict}</span>}</div>}
              </div>
              {typeof d.score === "number" && <span className="shrink-0 font-mono text-xs font-medium text-blue-800">{d.score}/100</span>}
              {canManage && <button onClick={() => setModalItem(d)} className="shrink-0 rounded border border-slate-200 px-1.5 py-0.5 text-xs text-slate-500 hover:bg-slate-50" title="Attach a file / score">{d.file || d.score != null ? <Star size={11} className="text-amber-500" /> : <Paperclip size={11} />}</button>}
              {canManage && <button onClick={() => remove(d.id)} className="shrink-0 rounded border border-slate-200 p-1 text-slate-300 hover:bg-rose-50 hover:text-rose-500" title="Remove"><Trash2 size={11} /></button>}
            </div>
          ); })}
        </div>
      )}

      {canManage && allDone && !work.completedAt && (
        <div className="mt-3 rounded-lg border border-emerald-200 bg-emerald-50 p-3">
          {confirmComplete ? (
            <div className="flex flex-wrap items-center justify-between gap-2">
              <span className="text-sm text-emerald-900">Mark this work{openActs > 0 ? ` and its ${openActs} open ${openActs === 1 ? "activity" : "activities"}` : ""} complete?</span>
              <span className="flex items-center gap-1.5"><button onClick={markComplete} className={btnDark}><CheckCircle2 size={14} /> Yes, complete it</button><button onClick={() => setConfirmComplete(false)} className={btnLight}>Cancel</button></span>
            </div>
          ) : (
            <div className="flex flex-wrap items-center justify-between gap-2">
              <span className="inline-flex items-center gap-1.5 text-sm text-emerald-900"><CheckCircle2 size={15} /> All deliverables are in.</span>
              <button onClick={() => setConfirmComplete(true)} className={btnDark}>Mark work complete</button>
            </div>
          )}
        </div>
      )}

      {modalItem && <DeliverableItemModal work={work} item={modalItem} store={store} busy={busy} setBusy={setBusy} flash={flash} onClose={() => setModalItem(null)} />}
    </div>
  );
}
