import { useState, useRef, useEffect, useCallback } from "react";
import {
  Lock, User, LogOut, LayoutGrid, Calendar, Users, ClipboardCheck, Plus,
  Sparkles, Play, Square, Check, X, ChevronLeft, ChevronRight, Target, Clock,
  AlertTriangle, Mic, FileText, Upload, Pencil, Trash2, Loader2, MessageSquare, Star,
  Gauge, Type as TypeIcon, Wand2, CalendarClock, ArrowRight, AlertCircle, ClipboardList,
} from "lucide-react";
import { api } from "./api";

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
  async decompose(title, type) { return parseJSON(await aiComplete('Enterprise delivery planner. Break a work into 3-6 sub-works, each with 1-3 activities. Return ONLY JSON: {"subworks":[{"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Work: "${title}". Type: ${type}.`)); },
  async extractMom(text) { return parseJSON(await aiComplete('Read meeting minutes, extract work. Return ONLY JSON: {"works":[{"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Minutes:\n"""${text}"""`)); },
  async modifyPlan(planText, instruction) { return parseJSON(await aiComplete('Edit a project plan. Return ONLY JSON: {"ops":[{"op":"add_activity","subwork":string,"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}|{"op":"add_subwork","title":string}|{"op":"retype","match":string,"type":"self"|"meeting"|"call"|"site"}]}', `Current plan:\n${planText}\n\nInstruction: ${instruction}`)); },
  async insight(title, m) { return parseJSON(await aiComplete('Execution advisor to a CEO. Return ONLY JSON: {"read":string(2 sentences),"action":string}', `Work "${title}". Planning ${m.planning}%, execution ${m.execution}%. Behind: ${m.behind}. Stuck: ${m.stuck || "none"}.`)); },
  async score(work, activity, content) { return parseJSON(await aiComplete('Delivery quality reviewer. Score fit of a deliverable to the activity/work, 0-100. Return ONLY JSON: {"score":number,"verdict":string(<=6 words),"feedback":string}', `Work: "${work}". Activity: "${activity}". Deliverable:\n"""${content}"""`)); },
};

/* ---------- dates ---------- */
const MSD = 86400000;
const sod = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
const TODAY = sod(new Date());
const addDays = (d, n) => sod(new Date(sod(d).getTime() + n * MSD));
const iso = (d) => sod(d).toISOString().slice(0, 10);
const parseISO = (s) => (s ? sod(new Date(s + "T00:00:00")) : null);
const startOfWeek = (d) => addDays(d, -((sod(d).getDay() + 6) % 7));
const fmtFull = (d) => d.toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
const isOverdue = (a) => a.date && parseISO(a.date) < TODAY && a.status !== "executed";

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
  const deliv = a.filter((x) => x.deliverable).length;
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
function sufficiency(m, st) { if (st.overdue > 0) return ["Behind", "rose"]; if (m.planning < 70) return ["Under-planned", "amber"]; if (m.resultPct != null && m.execution - m.resultPct > 25) return ["Output ahead of result", "amber"]; return ["Sufficient", "emerald"]; }
const loadOf = (acts, uid) => acts.filter((a) => a.assigneeId === uid && a.status !== "executed" && a.status !== "cancelled").reduce((s, a) => s + a.plannedHrs, 0);
const RAG = { green: ["bg-emerald-500", "text-emerald-700", "bg-emerald-50", "On track"], amber: ["bg-amber-500", "text-amber-700", "bg-amber-50", "At risk"], red: ["bg-rose-500", "text-rose-700", "bg-rose-50", "Behind"] };

/* ---------- tree / roll-up helpers ---------- */
const LEVEL_LABEL = { objective: "Objective", initiative: "Initiative", work: "Work", subwork: "Sub-work", activity: "Activity" };
const CHILD_LABEL = { objective: "initiatives", initiative: "works", work: "sub-works", subwork: "activities" };
const CHILD_LEVEL = { objective: "initiative", initiative: "work", work: "subwork", subwork: "activity" };
// One distinct colour per level so you always know your altitude. Kept clear of the RAG greens/ambers/reds.
const LEVEL_THEME = {
  objective: { name: "Objective", bar: "bg-violet-500", chip: "bg-violet-100 text-violet-700", dot: "bg-violet-500", ring: "border-violet-200", soft: "bg-violet-50", text: "text-violet-700" },
  initiative: { name: "Initiative", bar: "bg-blue-500", chip: "bg-blue-100 text-blue-700", dot: "bg-blue-500", ring: "border-blue-200", soft: "bg-blue-50", text: "text-blue-700" },
  work: { name: "Work", bar: "bg-teal-500", chip: "bg-teal-100 text-teal-700", dot: "bg-teal-500", ring: "border-teal-200", soft: "bg-teal-50", text: "text-teal-700" },
  subwork: { name: "Sub-work", bar: "bg-orange-500", chip: "bg-orange-100 text-orange-700", dot: "bg-orange-500", ring: "border-orange-200", soft: "bg-orange-50", text: "text-orange-700" },
  activity: { name: "Activity", bar: "bg-slate-400", chip: "bg-slate-100 text-slate-600", dot: "bg-slate-400", ring: "border-slate-200", soft: "bg-slate-50", text: "text-slate-600" },
};
const HOME_LEVEL = { md: "objective", vp: "initiative", member: "subwork" };
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

/* ---------- ui atoms ---------- */
const inputCls = "w-full rounded-md border border-slate-200 px-3 py-2 text-sm outline-none focus:border-slate-400";
const btnDark = "inline-flex items-center justify-center gap-1.5 rounded-md bg-slate-900 px-3 py-2 text-sm font-medium text-white hover:bg-slate-700 disabled:opacity-50";
const btnLight = "inline-flex items-center justify-center gap-1.5 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-medium text-slate-600 hover:bg-slate-50 disabled:opacity-50";
const btnViolet = "inline-flex items-center justify-center gap-1.5 rounded-md bg-violet-600 px-3 py-2 text-sm font-medium text-white hover:bg-violet-700 disabled:opacity-50";
function Avatar({ id, size = 22 }) { return <span className="inline-flex items-center justify-center rounded-full bg-slate-200 font-medium text-slate-600" style={{ width: size, height: size, fontSize: size * 0.42 }}>{initials(id)}</span>; }
function Chip({ children, tone = "slate" }) { const t = { slate: "bg-slate-100 text-slate-600", rose: "bg-rose-50 text-rose-700", amber: "bg-amber-50 text-amber-800", emerald: "bg-emerald-50 text-emerald-700", violet: "bg-violet-100 text-violet-700", blue: "bg-blue-100 text-blue-700" }[tone]; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${t}`}>{children}</span>; }
function LevelChip({ level }) { const th = LEVEL_THEME[level] || LEVEL_THEME.activity; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${th.chip}`}><span className={`h-1.5 w-1.5 rounded-full ${th.bar}`} /> {th.name}</span>; }
function StatusPill({ rag }) { const rg = RAG[rag]; return <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium ${rg[2]} ${rg[1]}`}><span className={`h-1.5 w-1.5 rounded-full ${rg[0]}`} /> {rg[3]}</span>; }
function LevelLegend() { return <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-slate-400"><span className="font-medium text-slate-500">Levels:</span>{["objective", "initiative", "work", "subwork", "activity"].map((l) => { const th = LEVEL_THEME[l]; return <span key={l} className="inline-flex items-center gap-1"><span className={`h-2 w-2 rounded-full ${th.bar}`} />{th.name}</span>; })}<span className="text-slate-300">→ each level nests inside the one before it</span></div>; }
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
  const clearDoc = () => { setDocName(null); };
  return (
    <div>
      <div className="mb-2 flex gap-1">{[["type", "Type", TypeIcon], ["voice", "Voice", Mic], ["document", "Attach / doc", FileText]].map(([k, l, I]) => <button key={k} onClick={() => setMode(k)} className={`inline-flex items-center gap-1 rounded-md px-2.5 py-1 text-xs font-medium ${mode === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}><I size={12} /> {l}</button>)}</div>
      {mode === "voice" && <div className="mb-2 flex items-center gap-2">{working === "transcribing" ? <span className="inline-flex items-center gap-1.5 text-sm text-slate-500"><Loader2 size={14} className="animate-spin" /> Transcribing…</span> : !rec ? <button onClick={start} className={btnLight}><Mic size={14} /> Record</button> : <button onClick={stop} className="inline-flex items-center gap-1.5 rounded-md bg-rose-500 px-3 py-2 text-sm font-medium text-white"><Square size={13} /> Stop</button>}<span className="text-xs text-slate-400">{CAP.deepgram ? "Deepgram speech-to-text · transcript is editable below." : "Transcript is editable below."}</span></div>}
      {mode === "document" && <label className="mb-2 flex cursor-pointer items-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50"><Upload size={14} /> Attach a document — PDF, Word (.docx), .md or .txt (or paste below)<input type="file" accept=".pdf,application/pdf,.docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,.txt,.md,text/plain" className="hidden" onChange={(e) => readF(e.target.files && e.target.files[0])} /></label>}
      {docName && <div className="mb-2 flex items-center gap-2 rounded-md border border-blue-200 bg-blue-50 px-2.5 py-1.5 text-xs text-blue-700"><FileText size={13} /> <span className="min-w-0 flex-1 truncate">{docName}</span>{working === "extracting" ? <Loader2 size={13} className="animate-spin" /> : <button onClick={clearDoc} className="text-blue-400 hover:text-blue-700"><X size={13} /></button>}</div>}
      {err && <div className="mb-2 rounded-md border border-amber-200 bg-amber-50 px-2.5 py-1.5 text-xs text-amber-700">{err}</div>}
      <textarea value={value} onChange={(e) => onChange(e.target.value)} rows={mode === "type" ? 4 : 5} placeholder={placeholder} className={inputCls} />
    </div>
  );
}

/* ==================================================================== */
export default function App() {
  const [me, setMe] = useState(null);
  const [works, setWorks] = useState([]);
  const [acts, setActs] = useState([]);
  const [crs, setCrs] = useState([]);
  const [teams, setTeams] = useState([]);
  const [remarks, setRemarks] = useState([]);
  const [tab, setTab] = useState("portfolio");
  const [openId, setOpenId] = useState(null);
  const [busy, setBusy] = useState(null);
  const [note, setNote] = useState(null);
  const [capture, setCapture] = useState(false);
  const [portView, setPortView] = useState("cards");
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
  const unreadNudges = me ? remarks.filter((r) => r.toIds.includes(me.id) && !r.readBy.includes(me.id)).length : 0;

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
      ops.forEach((op) => { if (op.op === "add_subwork") nw.push({ id: nid("w"), parentId: node.id, level: "subwork", title: op.title, type: node.type || "general", ownerId: node.ownerId }); });
      const liveSubs = container.concat(nw);
      ops.forEach((op) => { if (op.op === "add_activity") { const sw = liveSubs.find((s) => s.title.toLowerCase().includes((op.subwork || "").toLowerCase())) || liveSubs[0]; if (sw) na.push({ id: nid("a"), workId: sw.id, title: op.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(op.estimateHrs) || 2, actualHrs: null, actType: op.type || "self", unplanned: true }); } });
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

  const handleLogin = async (user) => { setLoading(true); try { applySnap(await api.snapshot()); } catch { flash("Couldn't reach the Cadence service — is it running on port 4000?"); } setMe(user); setLoading(false); };
  const logout = () => { setMe(null); setTab("portfolio"); setOpenId(null); setWorks([]); setActs([]); setCrs([]); setTeams([]); setRemarks([]); };

  // probe service capabilities (voice mode) once
  useEffect(() => { api.health().then((h) => { CAP.deepgram = !!(h && h.deepgram); }).catch(() => {}); }, []);
  // poll the shared store so Teams changes appear here (and vice-versa)
  useEffect(() => {
    if (!me) return;
    const t = setInterval(() => { if (pendingWrites.current === 0) refresh(); }, 3500);
    return () => clearInterval(t);
  }, [me, refresh]);

  if (!me) return <Login onLogin={handleLogin} />;
  const tops = works.filter((w) => w.parentId === null);
  const open = openId ? works.find((w) => w.id === openId) : null;
  const tabs = [["portfolio", me.level === "member" ? "My work" : "Portfolio", LayoutGrid], ["myday", "My day", Calendar]];
  if (isOrg) tabs.push(["team", "Team", Users], ["approvals", "Approvals", ClipboardCheck]);

  return (
    <div className="w-full bg-stone-50 text-slate-800" style={{ minHeight: 720, fontFeatureSettings: '"tnum"' }}>
      <div className="flex items-center justify-between border-b border-slate-200 bg-white px-5 py-3">
        <div className="flex items-center gap-2.5"><div className="flex h-7 w-7 items-center justify-center rounded-md bg-slate-900 font-mono text-sm text-white">C</div><div><div className="text-sm font-medium leading-none text-slate-900">Cadence</div><div className="mt-0.5 text-xs text-slate-400">work &amp; initiative OS · prototype</div></div></div>
        <div className="flex items-center gap-3"><button onClick={openNudges} className={`${btnLight} relative`} title="Nudges"><MessageSquare size={14} />{unreadNudges > 0 && <span className="absolute -right-1.5 -top-1.5 rounded-full bg-rose-500 px-1 text-xs font-medium text-white">{unreadNudges}</span>}</button><div className="text-right"><div className="text-xs font-medium text-slate-700">{me.name}</div><div className="text-xs text-slate-400">{me.title} · {me.level}</div></div><button onClick={logout} className={btnLight}><LogOut size={14} /></button></div>
      </div>
      <div className="flex items-center gap-1 border-b border-slate-200 bg-white px-5">
        {tabs.map(([k, l, I]) => <button key={k} onClick={() => { setTab(k); setOpenId(null); }} className={`flex items-center gap-2 border-b-2 px-3 py-2.5 text-sm ${tab === k ? "border-slate-900 text-slate-900" : "border-transparent text-slate-400 hover:text-slate-600"}`}><I size={15} /><span className="font-medium">{l}</span>{k === "approvals" && pending.length > 0 && <span className="rounded-full bg-rose-500 px-1.5 text-xs font-medium text-white">{pending.length}</span>}</button>)}
        {isOrg && <div className="ml-auto flex items-center gap-2 py-1.5">
          <button onClick={() => { setTab("portfolio"); setOpenId(null); setReviewMode(true); }} className={btnLight}><Pencil size={14} /> Review &amp; update</button>
          <button onClick={() => (me.level === "md" ? setObjModal(true) : setCapture(true))} className={btnDark}><Plus size={14} /> {me.level === "md" ? "New objective" : "New initiative"}</button>
        </div>}
      </div>
      {note && <div className="mx-5 mt-3 rounded-md border border-slate-200 bg-white px-3 py-2 text-xs text-slate-600">{note}</div>}
      <div className="p-5">
        {tab === "portfolio" && !open && <Portfolio {...{ works, acts, crs, me, isOrg, onOpen: setOpenId, goApprovals, view: portView, setView: setPortView, reviewMode, setReviewMode, onRemark: setRemarkNode }} />}
        {tab === "portfolio" && open && <NodeView {...{ nodeId: openId, user: me, works, acts, crs, isOrg, busy, setBusy, flash, patchAct, patchWork, store, onOpen: setOpenId, onRemark: setRemarkNode, goApprovals }} />}
        {tab === "myday" && <MyDay {...{ me, works, acts, busy, setBusy, flash, patchAct, store }} />}
        {tab === "team" && isOrg && <TeamView {...{ user: me, teams, store, works, acts, onOpen: (id) => { setTab("portfolio"); setOpenId(id); } }} />}
        {tab === "approvals" && isOrg && <Approvals {...{ crs, store, works, me, flash }} />}
      </div>
      {capture && <Capture {...{ me, teams, store, works, busy, setBusy, flash, onClose: () => setCapture(false), onOpen: (id) => { setCapture(false); setTab("portfolio"); setOpenId(id); } }} />}
      {objModal && <QuickCreate {...{ me, parent: null, level: "objective", works, store, busy, setBusy, flash, onClose: () => setObjModal(false), onCreated: (id) => { setTab("portfolio"); setPortView("cards"); setOpenId(id); } }} />}
      {nudgeOpen && <NudgeInbox {...{ remarks, me, onOpen: (id) => { setNudgeOpen(false); setTab("portfolio"); setOpenId(id); }, onClose: () => setNudgeOpen(false) }} />}
      {remarkNode && <RemarkModal {...{ node: remarkNode, works, acts, busy, setBusy, flash, onSubmit: addRemark, onClose: () => setRemarkNode(null) }} />}
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
function Panel({ title, right, children }) { return <div className="rounded-xl border border-slate-200 bg-white p-4"><div className="mb-3 flex items-center justify-between"><div className="text-sm font-medium text-slate-700">{title}</div>{right}</div>{children}</div>; }

function InitiativeCard({ row, isOrg, onOpen, onJump, goApprovals, reviewMode, onRemark }) {
  const { w, m, st, rag, issue, childCount, blocked, crc } = row; const th = LEVEL_THEME[w.level] || LEVEL_THEME.activity;
  return (
    <div className={`overflow-hidden rounded-xl border bg-white ${reviewMode ? "border-violet-300 ring-1 ring-violet-200" : rag === "red" ? "border-rose-200" : "border-slate-200"}`}>
      <div className={`h-1.5 ${th.bar}`} />
      <div className="p-5">
        <div className="flex items-start justify-between gap-3">
          <button onClick={() => (reviewMode ? onRemark(w) : onOpen(w.id))} className="min-w-0 text-left">
            <LevelChip level={w.level} />
            <div className="mt-1.5 truncate text-base font-medium text-slate-900 hover:underline">{w.title}</div>
            <div className="mt-1 flex items-center gap-1.5 text-xs text-slate-400"><Avatar id={w.ownerId} size={16} /> {uName(w.ownerId)} · {fnOf(w.ownerId)}</div>
          </button>
          <div className="flex shrink-0 flex-col items-end gap-1.5">
            <StatusPill rag={rag} />
            {w.result && <div className="text-right"><div className={`font-mono text-xl font-medium ${th.text}`}>{m.resultPct != null ? Math.round(m.resultPct) + "%" : "—"}</div><div className="text-xs text-slate-400">{w.result.current}/{w.result.target} {w.result.unit}</div></div>}
          </div>
        </div>
        <div className="mt-4"><ProgressPair planning={m.planning} execution={m.execution} /></div>
        <div className="mt-3 flex items-center justify-between text-xs text-slate-500">
          <span>{childCount} {CHILD_LABEL[w.level]} · {st.done}/{st.total} activities done{st.overdue > 0 ? ` · ${st.overdue} overdue` : ""}{blocked > 0 ? ` · ${blocked} blocked` : ""}</span>
          {reviewMode ? <button onClick={() => onRemark(w)} className="inline-flex items-center gap-1 font-medium text-violet-600 hover:text-violet-800"><Pencil size={12} /> Remark &amp; update</button> : <button onClick={() => onOpen(w.id)} className="inline-flex items-center gap-1 font-medium text-slate-500 hover:text-slate-800">Open <ArrowRight size={12} /></button>}
        </div>
        {isOrg && crc > 0 && <button onClick={goApprovals} className="mt-3 inline-flex items-center gap-1 rounded-md bg-amber-100 px-2 py-1 text-xs font-medium text-amber-800 hover:bg-amber-200"><ClipboardCheck size={12} /> {crc} approval{crc > 1 ? "s" : ""} to review <ArrowRight size={11} /></button>}
        {issue && <button onClick={() => onJump(issue)} className="mt-3 flex w-full items-center gap-2 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-left hover:bg-rose-100"><AlertTriangle size={14} className="shrink-0 text-rose-500" /><span className="min-w-0 flex-1 text-xs text-rose-800"><span className="font-medium">{issue.blocked ? "Blocked" : "Overdue"} task deep inside this {LEVEL_LABEL[w.level].toLowerCase()}:</span> {issue.title} — this is why it shows red.</span><span className="shrink-0 text-xs font-medium text-rose-700">Go to it <ArrowRight size={11} className="inline" /></span></button>}
      </div>
    </div>
  );
}

function SufficiencyPanel({ rows }) {
  const objs = rows.filter((r) => r.w.result);
  return (
    <Panel title="Sufficiency & gap to target">
      <div className="space-y-3">
        {objs.length === 0 && <div className="text-xs text-slate-400">No result-bearing objectives in scope.</div>}
        {objs.map(({ w, m, st }) => { const [lbl, tone] = sufficiency(m, st); const gap = Math.round((w.result.target - w.result.current) * 100) / 100;
          return (
            <div key={w.id} className="rounded-lg border border-slate-100 p-3">
              <div className="flex items-center justify-between gap-2"><span className="min-w-0 truncate text-sm font-medium text-slate-800">{w.title}</span><Chip tone={tone}>{lbl}</Chip></div>
              <div className="mt-1 text-xs text-slate-500">{w.result.metric}: {w.result.current} of {w.result.target} {w.result.unit} · <span className="font-medium text-slate-700">gap {gap} {w.result.unit}</span></div>
              <div className="mt-2 grid grid-cols-2 gap-4 text-xs">
                <div><div className="mb-0.5 flex justify-between"><span className="text-slate-400">Result attained</span><span className="font-mono text-violet-700">{m.resultPct != null ? Math.round(m.resultPct) : "—"}%</span></div><div className="h-1.5 rounded-full bg-violet-100 overflow-hidden"><div className="h-full bg-violet-500" style={{ width: `${m.resultPct || 0}%` }} /></div></div>
                <div><div className="mb-0.5 flex justify-between"><span className="text-slate-400">Plan coverage</span><span className="font-mono text-teal-700">{m.planning}%</span></div><div className="h-1.5 rounded-full bg-teal-100 overflow-hidden"><div className="h-full bg-teal-500" style={{ width: `${m.planning}%` }} /></div></div>
              </div>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

function ByFunctionPanel({ initRows }) {
  const byFn = {};
  initRows.forEach((r) => { const f = fnOf(r.w.ownerId) || "—"; (byFn[f] = byFn[f] || []).push(r); });
  return (
    <Panel title="By function">
      <div className="space-y-2">{Object.entries(byFn).map(([f, rs]) => { const att = rs.filter((r) => r.m.resultPct != null); const avg = att.length ? Math.round(att.reduce((s, r) => s + r.m.resultPct, 0) / att.length) : null; const worst = rs.some((r) => r.rag === "red") ? "red" : rs.some((r) => r.rag === "amber") ? "amber" : "green";
        return <div key={f} className="flex items-center gap-3 rounded-lg border border-slate-100 px-3 py-2"><span className={`h-2.5 w-2.5 rounded-full ${RAG[worst][0]}`} /><span className="flex-1 text-sm text-slate-700">{f}</span><span className="text-xs text-slate-400">{rs.length} initiative{rs.length > 1 ? "s" : ""}</span><span className="font-mono text-sm text-violet-700">{avg != null ? avg + "%" : "—"}</span></div>;
      })}</div>
    </Panel>
  );
}

function AttentionPanel({ rows, works, acts, onOpen }) {
  const scope = new Set(rows.flatMap((r) => subtreeIds(works, r.w.id)));
  const items = [];
  acts.filter((a) => scope.has(a.workId) && isOverdue(a)).forEach((a) => items.push(["rose", Clock, a.title, "overdue " + fmtFull(parseISO(a.date)), a.workId]));
  acts.filter((a) => scope.has(a.workId) && a.blocked && a.status !== "executed").forEach((a) => items.push(["rose", AlertTriangle, a.title, "blocked — needs help", a.workId]));
  rows.filter((r) => r.m.stuck).forEach((r) => items.push(["amber", AlertTriangle, r.w.title, "stuck at " + r.m.stuck, r.w.id]));
  acts.filter((a) => scope.has(a.workId) && a.unplanned && a.status !== "executed").forEach((a) => items.push(["amber", Plus, a.title, "unplanned — reshaped the plan", a.workId]));
  const topOf = (wid) => { const w = works.find((x) => x.id === wid); return w && w.parentId ? w.parentId : wid; };
  return (
    <Panel title="Needs attention" right={<span className="text-xs text-slate-400">{items.length}</span>}>
      {items.length === 0 && <div className="text-xs text-slate-400">All clear.</div>}
      <div className="space-y-1.5">{items.slice(0, 8).map((it, i) => { const [tone, Icon, t, sub, wid] = it;
        return <button key={i} onClick={() => onOpen(wid)} className="flex w-full items-center gap-2 rounded-md border border-slate-100 px-3 py-1.5 text-left hover:bg-slate-50"><Icon size={13} className={tone === "rose" ? "text-rose-500" : "text-amber-500"} /><span className="min-w-0 flex-1 truncate text-sm text-slate-700">{t}</span><span className="shrink-0 text-xs text-slate-400">{sub}</span></button>;
      })}</div>
    </Panel>
  );
}

function CapacityPanel({ users, acts }) {
  const cap = 20;
  return (
    <Panel title="Team capacity (open work)">
      <div className="space-y-2">{users.map((u) => { const h = loadOf(acts, u.id); const pct = Math.min(100, (h / cap) * 100); const over = h > cap * 0.8;
        return <div key={u.id} className="flex items-center gap-3"><span className="w-32 shrink-0 truncate text-sm text-slate-600">{u.name.split(" ")[0]} <span className="text-xs text-slate-400">· {u.fn}</span></span><div className="h-2 flex-1 rounded-full bg-slate-100 overflow-hidden"><div className={`h-full ${over ? "bg-rose-400" : "bg-emerald-400"}`} style={{ width: `${pct}%` }} /></div><span className="w-10 text-right font-mono text-xs text-slate-500">{h}h</span></div>;
      })}</div>
    </Panel>
  );
}

function TeamSnapshot({ user, acts }) {
  const reports = USERS.filter((u) => u.reports_to === user.id);
  return (
    <Panel title="Your team">
      <div className="space-y-2">{reports.length === 0 && <div className="text-xs text-slate-400">No direct reports.</div>}{reports.map((u) => { const mine = acts.filter((a) => a.assigneeId === u.id && a.status !== "cancelled"); const done = mine.filter((a) => a.status === "executed").length; const over = mine.filter((a) => isOverdue(a)).length; const exec = mine.length ? Math.round((done / mine.length) * 100) : 0;
        return <div key={u.id} className="flex items-center gap-3"><Avatar id={u.id} size={28} /><div className="min-w-0 flex-1"><div className="text-sm text-slate-700">{u.name}</div><div className="text-xs text-slate-400">{mine.length} activities · {done} done{over ? ` · ${over} overdue` : ""}</div></div><span className="font-mono text-sm text-amber-700">{exec}%</span></div>;
      })}</div>
    </Panel>
  );
}

function Portfolio({ works, acts, crs, me, isOrg, onOpen, goApprovals, view, setView, reviewMode, setReviewMode, onRemark }) {
  const roots = homeNodes(works, acts, me);
  const rows = roots.map((w) => { const m = computeMeters(works, acts, w.id); const st = workStats(works, acts, w.id); const att = attentionCount(works, acts, w.id); return { w, m, st, rag: nodeRag(works, acts, w.id), issue: deepestIssue(works, acts, w.id), childCount: works.filter((x) => x.parentId === w.id).length, blocked: att.blocked, crc: crs.filter((c) => c.workId === w.id && c.status === "pending").length }; });
  return (
    <div>
      {reviewMode && isOrg && <div className="mb-3 flex items-center justify-between gap-2 rounded-lg border border-violet-200 bg-violet-50 px-3 py-2 text-xs text-violet-800"><span className="flex items-center gap-1.5"><Pencil size={13} /> Review &amp; update mode — {view === "tree" ? "click any card in the tree" : "use the Remark button on any card"} to leave a remark, re-assign, or shift dates. Owners get nudged.</span><button onClick={() => setReviewMode(false)} className="rounded-md border border-violet-200 bg-white px-2 py-1 font-medium text-violet-700 hover:bg-violet-100">Exit</button></div>}
      <div className="mb-4 flex items-center justify-end">
        <div className="inline-flex items-center gap-0.5 rounded-lg border border-slate-200 bg-white p-0.5">
          <button onClick={() => setView("cards")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "cards" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}><LayoutGrid size={13} /> Cards</button>
          <button onClick={() => setView("tree")} className={`inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-medium ${view === "tree" ? "bg-slate-900 text-white" : "text-slate-500 hover:bg-slate-50"}`}><ChevronRight size={13} className="rotate-90" /> Tree</button>
        </div>
      </div>
      {view === "tree" ? <BracketTree roots={roots} works={works} acts={acts} onOpen={onOpen} reviewMode={reviewMode} setReviewMode={setReviewMode} onRemark={onRemark} />
        : me.level === "member" ? <MyWork {...{ me, works, acts, rows, onOpen }} />
        : <LeaderDashboard {...{ me, works, acts, crs, rows, isOrg, onOpen, goApprovals, reviewMode, onRemark }} />}
    </div>
  );
}

function LeaderDashboard({ me, works, acts, crs, rows, isOrg, onOpen, goApprovals, reviewMode, onRemark }) {
  const [filter, setFilter] = useState("all");
  const md = me.level === "md";
  const homeLabel = md ? "Objectives" : "Initiatives";
  const initRows = scopeInitiatives(works, acts, me).map((w) => { const m = computeMeters(works, acts, w.id); const st = workStats(works, acts, w.id); return { w, m, st, rag: nodeRag(works, acts, w.id) }; });
  const scopeIds = new Set(rows.flatMap((r) => subtreeIds(works, r.w.id)));
  const pending = crs.filter((c) => c.status === "pending" && scopeIds.has(c.workId)).length;
  const overdueTotal = rows.reduce((s, r) => s + r.st.overdue, 0);
  const att = initRows.filter((r) => r.m.resultPct != null); const avg = att.length ? Math.round(att.reduce((s, r) => s + r.m.resultPct, 0) / att.length) : "—";
  const capUsers = md ? USERS.filter((u) => u.level !== "md") : USERS.filter((u) => u.reports_to === me.id);
  const shown = rows.filter((r) => filter === "all" ? true : r.rag !== "green");
  const onJump = (issue) => onOpen(issue.workId);
  return (
    <div>
      <div className="mb-1 text-lg font-medium text-slate-900">{md ? "Enterprise scorecard" : `Your function — ${me.fn}`}</div>
      <div className="mb-2 text-xs text-slate-400">{md ? "Your objectives, rolled up from every level below. A red flag means something is broken deeper down — click it to jump to the cause." : "Your initiatives, scoped to your function. Drill in to walk the tree down to activities."}</div>
      <div className="mb-4"><LevelLegend /></div>
      <div className="mb-5 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        <Tile label={homeLabel} value={rows.length} />
        <Tile label="On track" value={rows.filter((r) => r.rag === "green").length} tone="text-emerald-600" />
        <Tile label="At risk / behind" value={rows.filter((r) => r.rag !== "green").length} tone="text-amber-600" />
        <Tile label="Avg result" value={typeof avg === "number" ? avg + "%" : avg} tone="text-violet-700" />
        <Tile label="Overdue below" value={overdueTotal} tone={overdueTotal ? "text-rose-600" : undefined} />
        <Tile label="Approvals" value={pending} tone={pending ? "text-rose-600" : undefined} />
      </div>
      <div className="mb-5 grid grid-cols-1 gap-4 lg:grid-cols-2">
        <SufficiencyPanel rows={initRows} />
        <AttentionPanel rows={rows} works={works} acts={acts} onOpen={onOpen} />
        {md ? <ByFunctionPanel initRows={initRows} /> : <TeamSnapshot user={me} acts={acts} />}
        <CapacityPanel users={capUsers} acts={acts} />
      </div>
      <div className="mb-3 flex items-center justify-between">
        <div className="text-sm font-medium text-slate-700">{homeLabel}</div>
        <div className="flex gap-1">{[["all", "All"], ["risk", "Needs attention"]].map(([k, l]) => <button key={k} onClick={() => setFilter(k)} className={`rounded-md px-3 py-1.5 text-xs font-medium ${filter === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>{l}</button>)}</div>
      </div>
      <div className="space-y-4">
        {shown.map((row) => <InitiativeCard key={row.w.id} row={row} isOrg={isOrg} onOpen={onOpen} onJump={onJump} goApprovals={goApprovals} reviewMode={reviewMode} onRemark={onRemark} />)}
        {shown.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-400">Nothing needs attention.</div>}
      </div>
    </div>
  );
}

function MyWork({ me, works, acts, rows, onOpen }) {
  const mine = acts.filter((a) => a.assigneeId === me.id && a.status !== "cancelled");
  const done = mine.filter((a) => a.status === "executed").length; const overdue = mine.filter((a) => isOverdue(a)).length;
  const scored = mine.filter((a) => a.deliverable); const avg = scored.length ? Math.round(scored.reduce((s, a) => s + a.deliverable.score, 0) / scored.length) : "—";
  return (
    <div>
      <div className="mb-1 text-lg font-medium text-slate-900">My work</div>
      <div className="mb-2 text-xs text-slate-400">You see only what you own or are assigned — not the enterprise portfolio. {me.title}.</div>
      <div className="mb-4"><LevelLegend /></div>
      <div className="mb-5 grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Tile label="Open activities" value={mine.length - done} />
        <Tile label="Done" value={done} tone="text-emerald-600" />
        <Tile label="Overdue" value={overdue} tone={overdue ? "text-rose-600" : undefined} />
        <Tile label="Avg deliverable" value={typeof avg === "number" ? avg + "/100" : avg} tone="text-violet-700" />
      </div>
      <div className="mb-3 text-sm font-medium text-slate-700">Sub-works I'm on</div>
      <div className="grid grid-cols-1 gap-3 lg:grid-cols-2">
        {rows.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white p-8 text-center text-sm text-slate-400">Nothing assigned to you yet.</div>}
        {rows.map((row) => { const { w, m } = row; const ct = LEVEL_THEME[w.level]; const myActs = acts.filter((a) => subtreeIds(works, w.id).includes(a.workId) && a.assigneeId === me.id && a.status !== "cancelled"); const myDone = myActs.filter((a) => a.status === "executed").length;
          return (
            <div key={w.id} className={`overflow-hidden rounded-xl border bg-white ${row.rag === "red" ? "border-rose-200" : "border-slate-200"}`}>
              <div className={`h-1 ${ct.bar}`} />
              <button onClick={() => onOpen(w.id)} className="block w-full p-4 text-left hover:bg-slate-50">
                <div className="flex items-start justify-between gap-2"><LevelChip level={w.level} /><StatusPill rag={row.rag} /></div>
                <div className="mt-1.5 truncate text-sm font-medium text-slate-800">{w.title}</div>
                <div className="mt-0.5 text-xs text-slate-500">My part: {myDone}/{myActs.length} activities done</div>
                <div className="mt-3"><ProgressPair planning={m.planning} execution={m.execution} /></div>
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------- Work detail ---------- */
function BracketTree({ roots, works, acts, onOpen, reviewMode, setReviewMode, onRemark }) {
  const COL = 258, CARDW = 216, CARDH = 76, ROW = 98;
  const rootIds = roots.map((r) => r.id);
  const childIds = (id) => works.filter((w) => w.parentId === id).map((w) => w.id); // tree stops at sub-work
  const allWorkIds = [...new Set(rootIds.flatMap((id) => subtreeIds(works, id).filter((x) => works.find((w) => w.id === x))))];
  const init = () => { const s = new Set(rootIds); roots.forEach((r) => { const iss = deepestIssue(works, acts, r.id); if (iss) { let cur = works.find((w) => w.id === iss.workId); let g = 0; while (cur && g++ < 12) { s.add(cur.id); if (cur.id === r.id) break; cur = works.find((w) => w.id === cur.parentId); } } }); return s; };
  const [expanded, setExpanded] = useState(init);
  const toggle = (id) => setExpanded((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });

  // ----- layout pass: each visible node gets a column (depth) and a y-slot -----
  const nodes = {}; let leaf = 0; let maxDepth = 0;
  const place = (id, depth) => {
    maxDepth = Math.max(maxDepth, depth);
    const kids = expanded.has(id) ? childIds(id) : [];
    let y;
    if (kids.length === 0) { y = leaf * ROW; leaf++; }
    else { const ys = kids.map((k) => place(k, depth + 1)); y = (ys[0] + ys[ys.length - 1]) / 2; }
    nodes[id] = { depth, y };
    return y;
  };
  rootIds.forEach((r) => place(r, 0));
  const edges = [];
  Object.keys(nodes).forEach((id) => { if (!expanded.has(id)) return; childIds(id).forEach((cid) => { if (nodes[cid]) edges.push([id, cid]); }); });
  const W = maxDepth * COL + CARDW + 24; const H = Math.max(leaf, 1) * ROW + 8;

  return (
    <div>
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <LevelLegend />
        <div className="flex gap-1"><button onClick={() => setExpanded(new Set(allWorkIds))} className={btnLight}>Expand all</button><button onClick={() => setExpanded(new Set(rootIds))} className={btnLight}>Collapse</button></div>
      </div>
      {reviewMode && <div className="mb-3 flex items-center justify-between gap-2 rounded-lg border border-violet-200 bg-violet-50 px-3 py-2 text-xs text-violet-800"><span className="flex items-center gap-1.5"><Pencil size={13} /> Review mode — click any initiative, work or sub-work to leave a remark &amp; nudge its owners.</span><button onClick={() => setReviewMode(false)} className="rounded-md border border-violet-200 bg-white px-2 py-1 font-medium text-violet-700 hover:bg-violet-100">Exit</button></div>}
      <div className="overflow-x-auto rounded-xl border border-slate-200 bg-slate-50 p-4">
        <div className="relative" style={{ width: W, height: H }}>
          <svg className="absolute inset-0" width={W} height={H}>
            {edges.map(([p, c], i) => { const pn = nodes[p], cn = nodes[c]; const px = pn.depth * COL + CARDW, py = pn.y + CARDH / 2, cx = cn.depth * COL, cy = cn.y + CARDH / 2; const mx = px + (cx - px) / 2; const red = nodeRag(works, acts, c) === "red"; return <path key={i} d={`M ${px} ${py} C ${mx} ${py}, ${mx} ${cy}, ${cx} ${cy}`} fill="none" stroke={red ? "#fb7185" : "#cbd5e1"} strokeWidth={red ? 2 : 1.5} />; })}
          </svg>
          {Object.entries(nodes).map(([id, n]) => {
            const left = n.depth * COL, top = n.y;
            const w = works.find((x) => x.id === id); const th = LEVEL_THEME[w.level]; const rag = nodeRag(works, acts, id); const catt = attentionCount(works, acts, id); const m = computeMeters(works, acts, id); const hasKids = childIds(id).length > 0; const open = expanded.has(id); const canRemark = reviewMode && w.level !== "objective";
            return (
              <div key={id} style={{ position: "absolute", left, top, width: CARDW, height: CARDH }} onClick={() => { if (canRemark) onRemark(w); else if (hasKids) toggle(id); }} className={`relative overflow-hidden rounded-lg border bg-white ${rag === "red" ? "border-rose-200 ring-1 ring-rose-100" : "border-slate-200"} ${canRemark ? "cursor-pointer ring-1 ring-violet-200 hover:ring-violet-400" : hasKids ? "cursor-pointer hover:shadow-sm" : ""}`}>
                <span className={`absolute inset-y-0 left-0 w-1 ${th.bar}`} />
                <div className="flex h-full flex-col justify-center gap-1 pl-3 pr-2">
                  <div className="flex items-center gap-1.5">
                    <span className={`h-2 w-2 shrink-0 rounded-full ${th.bar}`} />
                    <span className="min-w-0 flex-1 truncate text-xs font-medium text-slate-800">{w.title}</span>
                    {w.recurring && <span className="shrink-0 rounded bg-violet-100 px-1 text-xs font-medium text-violet-700" title="recurring">↻ {w.recurring.cadence}</span>}
                    {hasKids && <span onClick={(e) => { e.stopPropagation(); toggle(id); }} className="shrink-0 cursor-pointer rounded p-0.5 hover:bg-slate-100"><ChevronRight size={13} className={`text-slate-400 transition-transform ${open ? "rotate-90" : ""}`} /></span>}
                  </div>
                  <div className="flex items-center gap-1 text-xs text-slate-500"><Avatar id={w.ownerId} size={14} /><span className="min-w-0 truncate">{uFirst(w.ownerId)}</span>{w.result && <span className={`shrink-0 font-mono ${th.text}`}>· {m.resultPct != null ? Math.round(m.resultPct) + "%" : "—"}</span>}</div>
                  <div className="flex items-center gap-1.5 text-xs">
                    <span className="text-teal-600">Plan {m.planning}%</span>
                    <span className="text-amber-600">Done {m.execution}%</span>
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
      <div className="mt-2 px-1 text-xs text-slate-400">Shows Objective → Initiative → Work → Sub-work. Click a card to branch it out · ↻ = recurring work · the arrow opens full detail · scroll sideways for deep branches.</div>
    </div>
  );
}

function NodeView({ nodeId, user, works, acts, crs, isOrg, busy, setBusy, flash, patchAct, patchWork, store, onOpen, onRemark, goApprovals }) {
  const node = works.find((w) => w.id === nodeId);
  const [insight, setInsight] = useState(null);
  const [modify, setModify] = useState(false); const [addUn, setAddUn] = useState(false); const [deliv, setDeliv] = useState(null); const [edit, setEdit] = useState(null); const [confirmDel, setConfirmDel] = useState(null); const [addChild, setAddChild] = useState(false); const [suggest, setSuggest] = useState(null);
  if (!node) return <button onClick={() => onOpen(null)} className="text-sm text-slate-500">← Portfolio</button>;
  const childLevel = CHILD_LEVEL[node.level]; const childLabel = childLevel === "activity" ? "task" : (LEVEL_LABEL[childLevel] || "").toLowerCase();
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
    try { const raw = await aiComplete('For each sub-work, 1-3 activities. Return ONLY JSON: {"fills":[{"title":<given>,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}', `Parent "${node.title}". Sub-works: ${JSON.stringify(empty.map((e) => e.title))}`); const fills = parseJSON(raw).fills || []; const add = []; fills.forEach((f) => { const lf = empty.find((l) => l.title === f.title) || empty[0]; (f.activities || []).forEach((ac) => add.push({ id: nid("a"), workId: lf.id, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }); if (add.length) store.addActs(add); flash("AI filled the missing branches."); } catch { flash("AI unavailable."); }
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

  return (
    <div>
      <div className="mb-4 flex flex-wrap items-center gap-1.5 text-sm">
        <button onClick={() => onOpen(null)} className="text-slate-500 hover:text-slate-800">Portfolio</button>
        {path.map((n, i) => { const t = LEVEL_THEME[n.level]; const last = i === path.length - 1; return <span key={n.id} className="flex items-center gap-1.5"><ChevronRight size={14} className="text-slate-300" /><span className={`h-2 w-2 rounded-full ${t.bar}`} />{last ? <span className="font-medium text-slate-800">{n.title}</span> : <button onClick={() => onOpen(n.id)} className="max-w-xs truncate text-slate-500 hover:text-slate-800">{n.title}</button>}</span>; })}
      </div>

      {/* ---- SECTION A: the level you're standing on ---- */}
      <div className={`overflow-hidden rounded-xl border ${th.ring} bg-white`}>
        <div className={`flex items-center justify-between gap-3 ${th.soft} px-5 py-3`}>
          <div className="flex items-center gap-2"><LevelChip level={node.level} />{node.recurring && <Chip tone="violet">↻ {node.recurring.cadence}</Chip>}{node.scope && <Chip tone={node.scope === "group" ? "blue" : "slate"}>{node.scope === "group" ? "group" : "individual"}</Chip>}</div>
          <div className="flex items-center gap-2">{canRemark && <button onClick={() => onRemark(node)} className="inline-flex items-center gap-1 rounded-md border border-violet-200 bg-white px-2.5 py-1 text-xs font-medium text-violet-700 hover:bg-violet-50"><Pencil size={12} /> Remark &amp; update</button>}<StatusPill rag={rag} /></div>
        </div>
        <div className="p-5">
          <h2 className="text-lg font-medium text-slate-900">{node.title}</h2>
          <div className="mt-1 flex flex-wrap items-center gap-1.5 text-xs text-slate-400"><Avatar id={node.ownerId} size={18} /> Owner {uName(node.ownerId)}{node.deadline && <span className="ml-2 inline-flex items-center gap-1"><Target size={12} /> deadline {fmtFull(parseISO(node.deadline))}</span>}{st.nextDue && <span className="ml-2 inline-flex items-center gap-1"><CalendarClock size={12} /> next {fmtFull(parseISO(st.nextDue))}</span>}</div>
          <div className="mt-4 max-w-md"><ProgressPair planning={m.planning} execution={m.execution} size="lg" /></div>
          <div className="mt-3 text-xs text-slate-500">{!leaf ? `${children.length} ${CHILD_LABEL[node.level]} inside · ` : ""}{st.done}/{st.total} activities done{st.overdue > 0 ? ` · ${st.overdue} overdue` : ""}</div>
        </div>
      </div>

      <div className="mt-4 space-y-4">
        {(att.blocked + att.overdue) > 0 && issue && (
          <div className="rounded-xl border border-rose-200 bg-rose-50 p-4">
            <div className="flex items-start gap-2.5">
              <AlertTriangle size={16} className="mt-0.5 shrink-0 text-rose-500" />
              <div className="min-w-0 flex-1">
                <div className="text-sm font-medium text-rose-800">Heads-up — {att.blocked > 0 ? `${att.blocked} blocked` : ""}{att.blocked > 0 && att.overdue > 0 ? " and " : ""}{att.overdue > 0 ? `${att.overdue} overdue` : ""} deeper down</div>
                <div className="mt-0.5 text-xs text-rose-700">That's why this {LEVEL_LABEL[node.level].toLowerCase()} shows red. Root cause: <span className="font-medium">{issue.title}</span> — {issue.blocked ? "a task that's blocked (someone can't move forward until it's cleared)." : "a task past its due date."}</div>
                {issue.workId !== node.id && <button onClick={() => onOpen(issue.workId)} className="mt-2 inline-flex items-center gap-1 rounded-md border border-rose-200 bg-white px-3 py-1.5 text-xs font-medium text-rose-700 hover:bg-rose-100">Go to the cause <ArrowRight size={12} /></button>}
              </div>
            </div>
          </div>
        )}
        {node.result && <ResultCard work={node} m={m} isOrg={isOrg} patchWork={patchWork} />}
        {isOrg && myCRs.length > 0 && <button onClick={goApprovals} className="flex w-full items-center justify-between rounded-xl border border-amber-200 bg-amber-50 p-4 text-left hover:bg-amber-100"><div><div className="text-sm font-medium text-amber-800">{myCRs.length} plan change{myCRs.length > 1 ? "s" : ""} awaiting your approval</div><div className="text-xs text-amber-700">Proposed by the team</div></div><span className="inline-flex items-center gap-1 rounded-md bg-white px-3 py-1.5 text-sm font-medium text-amber-800"><ClipboardCheck size={14} /> Review <ArrowRight size={13} /></span></button>}
        {isOrg && (
          <div className="rounded-xl border border-violet-200 bg-violet-50 p-4">
            <div className="mb-2 flex items-center gap-1.5 text-sm font-medium text-violet-800"><Sparkles size={15} /> AI &amp; plan tools</div>
            {node.level === "objective" ? (
              <>
                <div className="flex flex-wrap gap-2">
                  <button onClick={doSuggest} disabled={!!busy} className={btnLight}>{busy === "suggest" ? <Loader2 size={14} className="animate-spin" /> : <Wand2 size={14} />} Suggest initiatives</button>
                  <button onClick={doInsight} disabled={!!busy} className={btnLight}>{busy === "insight" ? <Loader2 size={14} className="animate-spin" /> : <Gauge size={14} />} Where does this stand</button>
                </div>
                <div className="mt-2 text-xs text-violet-700">At objective level AI shapes strategy — it proposes initiatives and reads the roll-up. Scheduling and assignment happen inside each initiative.</div>
                {suggest && <div className="mt-3 space-y-1">{suggest.length === 0 && <div className="text-sm text-slate-500">No new suggestions.</div>}{suggest.map((s, i) => <div key={i} className="flex items-center gap-2 rounded-md border border-violet-100 bg-white px-3 py-1.5 text-sm"><span className="min-w-0 flex-1 truncate text-slate-700">{s.title}</span><Chip tone="blue">{s.type}</Chip><button onClick={() => addSuggested(s)} className="shrink-0 rounded bg-slate-900 px-2 py-1 text-xs font-medium text-white hover:bg-slate-700">Add</button></div>)}</div>}
              </>
            ) : (
              <>
                <div className="flex flex-wrap gap-2">
                  <button onClick={autoAssign} disabled={!!busy} className={btnLight}><Wand2 size={14} /> Auto-assign</button>
                  <button onClick={expand} disabled={!!busy} className={btnLight}>{busy === "expand" ? <Loader2 size={14} className="animate-spin" /> : <Plus size={14} />} Fill missing</button>
                  <button onClick={doInsight} disabled={!!busy} className={btnLight}>{busy === "insight" ? <Loader2 size={14} className="animate-spin" /> : <Gauge size={14} />} Where do I stand</button>
                  <button onClick={() => setModify(true)} disabled={!!busy} className={btnLight}><Pencil size={14} /> Modify plan</button>
                  <button onClick={() => setAddUn(true)} disabled={!!busy} className={btnLight}><Plus size={14} /> Unplanned</button>
                </div>
                {insight && <div className="mt-3 space-y-1 text-sm"><p className="text-slate-700">{insight.read}</p><p className="text-violet-800"><span className="font-medium">Next: </span>{insight.action}</p></div>}
              </>
            )}
            {node.level === "objective" && insight && <div className="mt-3 space-y-1 text-sm"><p className="text-slate-700">{insight.read}</p><p className="text-violet-800"><span className="font-medium">Next: </span>{insight.action}</p></div>}
          </div>
        )}
      </div>

      {/* ---- divider: bifurcation into what's inside ---- */}
      <div className="mb-3 mt-6 flex items-center gap-2">
        <span className={`h-2.5 w-2.5 rounded-full ${leaf ? LEVEL_THEME.activity.bar : LEVEL_THEME[CHILD_LEVEL[node.level]].bar}`} />
        <span className="text-sm font-medium text-slate-700">Inside — {leaf ? `${nodeActs.length} ${nodeActs.length === 1 ? "activity" : "activities"}` : `${children.length} ${CHILD_LABEL[node.level]}`}</span>
        <span className="hidden text-xs text-slate-400 sm:inline">{leaf ? "the actual tasks to execute" : "click any to drill in one level"}</span>
        <div className="h-px flex-1 bg-slate-200" />
        {canCreate && <button onClick={() => setAddChild(true)} className={btnLight}><Plus size={14} /> Add {childLabel}</button>}
      </div>

      {/* ---- SECTION B: the level below ---- */}
      {!leaf && (
        <div className="grid grid-cols-1 gap-3 lg:grid-cols-2">
          {children.map((c) => {
            const ct = LEVEL_THEME[c.level]; const cm = computeMeters(works, acts, c.id); const cst = workStats(works, acts, c.id); const crag = nodeRag(works, acts, c.id); const cissue = deepestIssue(works, acts, c.id); const grand = works.filter((x) => x.parentId === c.id).length;
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
      )}

      {leaf && (
        <div className="space-y-2">
          {nodeActs.length === 0 && <div className="rounded-xl border border-dashed border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-400">No activities yet{isOrg ? " — use Auto-assign, Fill missing, or Unplanned above." : "."}</div>}
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
                  <div className="mt-1.5 flex flex-wrap items-center gap-1.5 pl-6 text-xs">
                    {a.assigneeId ? <span className="inline-flex items-center gap-1 text-slate-600"><Avatar id={a.assigneeId} size={16} /> {uFirst(a.assigneeId)}</span> : <span className="text-amber-700">Unassigned</span>}
                    <span className="text-slate-300">·</span>
                    <span className={over ? "text-rose-600" : "text-slate-500"}>{a.date ? fmtFull(parseISO(a.date)) : "No date"}</span>
                    <span className="text-slate-300">·</span><span className="text-slate-500">{a.actType}</span>
                    <span className="text-slate-300">·</span><span className="font-mono text-slate-400">{a.plannedHrs}h</span>
                    <button onClick={() => setDeliv(a)} className="ml-auto inline-flex items-center gap-1 rounded border border-slate-200 px-1.5 py-0.5 text-slate-500 hover:bg-slate-50">{a.deliverable ? <><Star size={10} className="text-amber-500" /> {a.deliverable.score}/100</> : <><FileText size={10} /> deliverable</>}</button>
                  </div>
                  {a.blocked && <div className="mt-1.5 pl-6 text-xs text-rose-600">Blocked — this task can't move forward until it's cleared. It's what's turning the levels above red.</div>}
                  {a.deliverable && <div className="mt-1 pl-6 text-xs text-slate-500">{a.deliverable.verdict} — {a.deliverable.feedback}</div>}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {modify && <ModifyPlan {...{ work: node, planText: planText(), subs, acts, store, busy, setBusy, flash, onClose: () => setModify(false) }} />}
      {addUn && <AddUnplanned {...{ subs, store, flash, onClose: () => setAddUn(false) }} />}
      {deliv && <Deliverable {...{ activity: deliv, work: node, store, busy, setBusy, flash, onClose: () => setDeliv(null) }} />}
      {edit && <ActivityEdit {...{ activity: edit, onSave: (p) => { patchAct(edit.id, p); setEdit(null); flash("Activity updated."); }, onClose: () => setEdit(null) }} />}
      {addChild && <QuickCreate {...{ me: user, parent: node, level: childLevel, works, store, busy, setBusy, flash, onClose: () => setAddChild(false), onCreated: (id) => { if (childLevel !== "activity") onOpen(id); } }} />}
    </div>
  );
}
function ResultCard({ work, m, isOrg, patchWork }) {
  const [edit, setEdit] = useState(false); const [val, setVal] = useState(work.result.current); const rt = work.result;
  return (
    <div className="rounded-xl border border-slate-200 bg-white p-4">
      <div className="mb-1 flex items-center justify-between"><span className="text-xs font-medium uppercase tracking-wide text-slate-400">Result — planned vs actual</span>{isOrg && <button onClick={() => setEdit(!edit)} className="text-xs text-violet-600">{edit ? "close" : "update actual"}</button>}</div>
      <div className="text-sm text-slate-700">{rt.metric}</div>
      <div className="mt-1 flex items-baseline gap-2"><span className="font-mono text-xl font-medium text-slate-900">{rt.current}</span><span className="text-xs text-slate-400">now · planned target {rt.target} {rt.unit} (from {rt.baseline})</span></div>
      {m.resultPct != null && <div className="mt-2 flex items-center gap-2"><div className="h-2 flex-1 rounded-full bg-violet-100 overflow-hidden"><div className="h-full rounded-full bg-violet-500" style={{ width: `${m.resultPct}%` }} /></div><span className="font-mono text-xs font-medium text-violet-700">{Math.round(m.resultPct)}%</span></div>}
      {edit && <div className="mt-3 flex items-center gap-2"><input type="number" value={val} onChange={(e) => setVal(e.target.value)} className="w-28 rounded-md border border-slate-200 px-2 py-1 text-sm" /><button onClick={() => { patchWork(work.id, { result: { ...rt, current: Number(val) } }); setEdit(false); }} className={btnDark}>Save</button></div>}
    </div>
  );
}
function ActivityEdit({ activity, onSave, onClose }) {
  const [assignee, setAssignee] = useState(activity.assigneeId || ""); const [date, setDate] = useState(activity.date || ""); const [type, setType] = useState(activity.actType); const [hrs, setHrs] = useState(activity.plannedHrs);
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Edit activity</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 rounded-md bg-slate-50 px-3 py-2 text-sm text-slate-700">{activity.title}</div>
      <div className="space-y-3">
        <div><label className="mb-1 block text-xs text-slate-500">Assign to</label><select value={assignee} onChange={(e) => setAssignee(e.target.value)} className={inputCls}><option value="">Unassigned</option>{USERS.map((u) => <option key={u.id} value={u.id}>{u.name} — {u.title}</option>)}</select></div>
        <div className="grid grid-cols-3 gap-2">
          <div className="col-span-2"><label className="mb-1 block text-xs text-slate-500">Date</label><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} /></div>
          <div><label className="mb-1 block text-xs text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div>
        </div>
        <div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
      </div>
      <button onClick={() => onSave({ assigneeId: assignee || null, date: date || null, actType: type, plannedHrs: hrs })} className={`${btnDark} mt-4 w-full`}>Save changes</button>
    </Modal>
  );
}

/* ---------- My day (real calendar) ---------- */
function MyDay({ me, works, acts, busy, setBusy, flash, patchAct, store }) {
  const [anchor, setAnchor] = useState(startOfWeek(TODAY));
  const [sel, setSel] = useState(iso(TODAY));
  const [quick, setQuick] = useState(false); const [deliv, setDeliv] = useState(null); const [propose, setPropose] = useState(null);
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
                      <button onClick={() => setDeliv(a)} className="inline-flex items-center gap-1 text-xs text-slate-500 hover:text-slate-800"><FileText size={11} /> {a.deliverable ? `deliverable ${a.deliverable.score}/100` : "attach deliverable"}</button>
                      <button onClick={() => setPropose(a)} className="inline-flex items-center gap-1 text-xs text-violet-600 hover:text-violet-800"><Pencil size={11} /> propose change</button>
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
      {deliv && <Deliverable {...{ activity: deliv, work: works.find((w) => w.id === deliv.workId), store, busy, setBusy, flash, onClose: () => setDeliv(null) }} />}
      {propose && <ProposeChange {...{ activity: propose, me, store, flash, onClose: () => setPropose(null) }} />}
    </div>
  );
}
function QuickAdd({ me, works, date, store, flash, onClose }) {
  const leaves = works.filter((w) => !works.some((x) => x.parentId === w.id));
  const [wid, setWid] = useState(leaves[0] ? leaves[0].id : ""); const [title, setTitle] = useState(""); const [hrs, setHrs] = useState(1); const [type, setType] = useState("self");
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Add an activity</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Adds to your day on {fmtFull(parseISO(date))}, assigned to you.</p>
      <label className="mb-1 block text-xs text-slate-500">Under work</label><select value={wid} onChange={(e) => setWid(e.target.value)} className={`${inputCls} mb-3`}>{leaves.map((w) => <option key={w.id} value={w.id}>{w.title}</option>)}</select>
      <label className="mb-1 block text-xs text-slate-500">Activity</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder="What are you doing?" />
      <div className="mb-4 grid grid-cols-2 gap-2"><div><label className="mb-1 block text-xs text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
      <button onClick={() => { if (!title.trim() || !wid) return; store.addActs([{ id: nid("a"), workId: wid, title: title.trim(), assigneeId: me.id, date, status: "planned", plannedHrs: hrs, actualHrs: null, actType: type }]); flash("Added to your day."); onClose(); }} disabled={!title.trim()} className={`${btnDark} w-full`}>Add</button>
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
  const stats = (uid) => { const mine = acts.filter((a) => a.assigneeId === uid && a.status !== "cancelled"); const done = mine.filter((a) => a.status === "executed"); const over = mine.filter((a) => isOverdue(a)); const scored = mine.filter((a) => a.deliverable); return { assigned: mine.length, done: done.length, over: over.length, exec: mine.length ? Math.round((done.length / mine.length) * 100) : 0, avg: scored.length ? Math.round(scored.reduce((s, a) => s + a.deliverable.score, 0) / scored.length) : null }; };
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
                    {isOpen && <div className="bg-slate-50 px-4 py-3"><div className="space-y-1.5">{mine.length === 0 && <div className="text-xs text-slate-400">No activities.</div>}{mine.map((a) => <div key={a.id} className="flex items-center gap-2 rounded-md border border-slate-100 bg-white px-3 py-1.5 text-sm"><button onClick={() => { const top = works.find((x) => x.id === a.workId); onOpen(top && top.parentId ? top.parentId : a.workId); }} className="min-w-0 flex-1 truncate text-left text-slate-700 hover:text-slate-900">{a.title}</button><span className="hidden shrink-0 truncate text-xs text-slate-400 sm:inline" style={{ maxWidth: 180 }}>{wt(a.workId)}</span><span className="shrink-0 text-xs text-slate-400">{a.date ? fmtFull(parseISO(a.date)) : "—"}</span>{a.deliverable && <span className="shrink-0 font-mono text-xs text-amber-700">{a.deliverable.score}</span>}<Chip tone={a.status === "executed" ? "blue" : isOverdue(a) ? "rose" : "slate"}>{a.status === "executed" ? "done" : isOverdue(a) ? "overdue" : a.status}</Chip></div>)}</div></div>}
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
  const build = (topTitle, ty, subs) => {
    const topId = nid("w"); const tpl = METRIC_BY_TYPE[ty] || METRIC_BY_TYPE.general; const memberIds = team ? team.memberIds : [];
    const nw = [{ id: topId, parentId: parentObj || null, level: "initiative", title: topTitle, type: ty, ownerId: me.id, teamId: teamId || null, scope: memberIds.length > 1 ? "group" : "individual", objective: objective || null, deadline: deadline || null, result: { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 } }];
    const na = []; const load = {}; memberIds.forEach((id) => (load[id] = 0)); let ai = 0;
    const dl = deadline ? parseISO(deadline) : null; const span = dl ? Math.max(1, Math.round((dl - TODAY) / MSD)) : 5;
    (subs || []).forEach((sw) => { const sid = nid("w"); nw.push({ id: sid, parentId: topId, level: "subwork", title: sw.title, type: ty, ownerId: me.id }); (sw.activities || []).forEach((ac) => { let assignee = null, date = null; if (memberIds.length) { assignee = memberIds.reduce((a, b) => (load[a] <= load[b] ? a : b)); load[assignee] += Number(ac.estimateHrs) || 2; date = iso(addDays(TODAY, ai % span)); ai++; } na.push({ id: nid("a"), workId: sid, title: ac.title, assigneeId: assignee, date, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" }); }); });
    store.addWorks(nw); if (na.length) store.addActs(na); return topId;
  };
  const run = async () => {
    if (!title.trim() && !text.trim()) return; setBusy("cap");
    try { const ctx = `${title.trim()}${objective ? ". Objective: " + objective : ""}${text ? ". Plan shared by the lead: " + text : ""}`.slice(0, 1500); const out = await AI.decompose(ctx, type); const id = build(title.trim() || text.slice(0, 50), type, out.subworks); flash(team ? `AI drafted the plan and assigned it to ${team.name}.` : "AI drafted the plan."); onOpen(id); }
    catch { const id = build(title.trim() || "New work", type, [{ title: "Plan", activities: [{ title: "Define scope & owners", estimateHrs: 2, type: "self" }] }]); flash("AI unavailable — created a starter you can edit."); onOpen(id); }
    setBusy(null);
  };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-4 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Capture &amp; plan a work</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 grid grid-cols-3 gap-2"><div className="col-span-2"><label className="mb-1 block text-xs text-slate-500">Title</label><input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. Roll out MFA to all field staff" className={inputCls} /></div><div><label className="mb-1 block text-xs text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{Object.keys(METRIC_BY_TYPE).map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
      <div className="mb-3"><label className="mb-1 block text-xs text-slate-500">Sits under objective</label><select value={parentObj} onChange={(e) => setParentObj(e.target.value)} className={inputCls}><option value="">None (top-level initiative)</option>{objectives.map((o) => <option key={o.id} value={o.id}>{o.title}</option>)}</select></div>
      <div className="mb-3 grid grid-cols-3 gap-2">
        <div className="col-span-2"><label className="mb-1 block text-xs text-slate-500">Assign to team</label><div className="flex gap-2"><select value={teamId} onChange={(e) => setTeamId(e.target.value)} className={inputCls}><option value="">No team (leave unassigned)</option>{teams.map((t) => <option key={t.id} value={t.id}>{t.name} ({t.memberIds.length})</option>)}</select><button onClick={() => setTeamModal(true)} className={btnLight}><Users size={14} /> New / merge</button></div></div>
        <div><label className="mb-1 block text-xs text-slate-500">Deadline</label><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className={inputCls} /></div>
      </div>
      {team && <div className="mb-3 flex flex-wrap items-center gap-1.5">{team.memberIds.map((id) => <span key={id} className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2 py-0.5 text-xs text-slate-600"><Avatar id={id} size={16} /> {uFirst(id)} <span className="text-slate-400">{fnOf(id)}</span></span>)}</div>}
      <label className="mb-1 block text-xs text-slate-500">Objective (optional)</label><input value={objective} onChange={(e) => setObjective(e.target.value)} className={`${inputCls} mb-3`} placeholder="The result this should move" />
      <label className="mb-1 block text-xs text-slate-500">Plan from the head — type, speak, or attach (the AI uses this to build the sub-works)</label>
      <MultiModalInput value={text} onChange={setText} placeholder="Paste or speak the plan the lead shared, or attach the note…" />
      <button onClick={run} disabled={busy || (!title.trim() && !text.trim())} className={`${btnViolet} mt-3 w-full`}>{busy ? <><Loader2 size={15} className="animate-spin" /> Drafting &amp; assigning…</> : <><Sparkles size={15} /> Draft plan &amp; assign to team</>}</button>
      <p className="mt-2 text-center text-xs text-slate-400">AI builds objective → work → sub-works → activities, and distributes them across the team by load.</p>
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
  const opLabel = (op) => op.op === "add_activity" ? `Add task “${op.title}” to ${op.subwork || container[0].title}` : op.op === "add_subwork" ? `Add sub-work “${op.title}”` : `Retype “${op.match}” → ${op.type}`;
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Remark &amp; update plan</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <div className="mb-3 flex items-center gap-2 rounded-md bg-slate-50 px-3 py-2"><span className={`h-2.5 w-2.5 rounded-full ${th.bar}`} /><span className={`rounded px-1.5 py-0.5 text-xs font-medium ${th.chip}`}>{LEVEL_LABEL[node.level]}</span><span className="min-w-0 truncate text-sm font-medium text-slate-800">{node.title}</span></div>
      <label className="mb-1 block text-xs font-medium text-slate-600">Your remark</label>
      <textarea value={text} onChange={(e) => setText(e.target.value)} rows={3} className={`${inputCls} mb-3`} placeholder="e.g. Pull the vendor demo forward; loop in Finance before the PO goes out; HOD sign-off should be a meeting." />
      <div className="mb-3 grid grid-cols-2 gap-3">
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Re-assign owner (optional)</label><select value={owner} onChange={(e) => setOwner(e.target.value)} className={inputCls}>{USERS.map((u) => <option key={u.id} value={u.id}>{u.name} · {u.fn}</option>)}</select></div>
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Shift dates + deadline (days)</label><input type="number" value={shift} onChange={(e) => setShift(Number(e.target.value) || 0)} className={inputCls} placeholder="e.g. -3 or 5" /></div>
      </div>
      <div className="mb-3 rounded-lg border border-violet-200 bg-violet-50 p-3">
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-1.5 text-xs font-medium text-violet-800"><Sparkles size={14} /> Turn my remark into plan changes for the team below</div>
          <button onClick={draftOps} disabled={busy || !text.trim()} className={btnLight}>{busy === "remarkops" ? <Loader2 size={13} className="animate-spin" /> : <Wand2 size={13} />} Draft</button>
        </div>
        {ops && <div className="mt-2 space-y-1">{ops.length === 0 && <div className="text-xs text-slate-500">AI didn't propose concrete changes — the remark alone will be sent.</div>}{ops.map((op, i) => <div key={i} className="flex items-center gap-2 rounded-md bg-white px-3 py-1.5 text-sm text-slate-700"><span className="min-w-0 flex-1 truncate">{opLabel(op)}</span><button onClick={() => setOps(ops.filter((_, j) => j !== i))} className="shrink-0 text-slate-300 hover:text-rose-500"><X size={13} /></button></div>)}{ops.length > 0 && <div className="text-xs text-violet-700">These apply to the plan beneath this {LEVEL_LABEL[node.level].toLowerCase()} when you send.</div>}</div>}
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
      {mine.length === 0 && <div className="rounded-md border border-dashed border-slate-200 px-3 py-8 text-center text-sm text-slate-400">No nudges yet. When a leader remarks on your initiative, work, or sub-work, it lands here.</div>}
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
      else if (level === "initiative") { sys = 'Draft an initiative that fulfils the objective. Return ONLY JSON: {"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","subworks":[{"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}'; usr = `Objective: "${parent.title}". Note:\n"""${text}"""`; }
      else if (level === "work") { sys = 'Draft a work (a phase of execution). Return ONLY JSON: {"title":string,"type":"procurement"|"cost"|"onboarding"|"compliance"|"general","subworks":[{"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}]}'; usr = `Initiative: "${parent.title}". Note:\n"""${text}"""`; }
      else if (level === "subwork") { sys = 'Draft a sub-work with 1-4 concrete activities. Return ONLY JSON: {"title":string,"activities":[{"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}]}'; usr = `Parent: "${parent.title}". Note:\n"""${text}"""`; }
      else { sys = 'Turn the note into ONE task. Return ONLY JSON: {"title":string,"estimateHrs":number,"type":"self"|"meeting"|"call"|"site"}'; usr = `Sub-work: "${parent.title}". Note:\n"""${text}"""`; }
      const d = parseJSON(await aiComplete(sys, usr, pdf ? pdf.data : undefined));
      if (d.title) setTitle(d.title);
      if (d.type && (isObj || isContainer)) setType(d.type);
      if (isObj && d.metric) { setHasMetric(true); setMetric(d.metric); setUnit(d.unit || "%"); setTarget(d.target || 100); }
      if (isAct) { if (d.estimateHrs) setHrs(Number(d.estimateHrs) || 2); if (d.type) setActType(d.type); }
      if (level === "subwork") setSubs(d.activities || []);
      if (isContainer) setSubs(d.subworks || []);
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
    if (level === "subwork") { subs.forEach((ac) => na.push({ id: nid("a"), workId: topId, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }
    else if (isContainer) { subs.forEach((sw) => { const sid = nid("w"); nw.push({ id: sid, parentId: topId, level: "subwork", title: sw.title, ownerId: me.id, type }); (sw.activities || []).forEach((ac) => na.push({ id: nid("a"), workId: sid, title: ac.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || "self" })); }); }
    store.addWorks(nw); if (na.length) store.addActs(na);
    flash(`${label} created.`); onCreated && onCreated(topId); onClose();
  };
  return (
    <Modal onClose={onClose} wide>
      <div className="mb-1 flex items-center justify-between"><h3 className="flex items-center gap-2 text-sm font-medium text-slate-900"><span className={`h-2.5 w-2.5 rounded-full ${th.bar}`} /> New {label.toLowerCase()}</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">{parent ? <>Under <span className="font-medium text-slate-500">{parent.title}</span>. </> : null}Fill the fields, or describe / dictate / attach a PDF below and let AI fill them.</p>

      <label className="mb-1 block text-xs font-medium text-slate-600">{label} name</label>
      <input value={title} onChange={(e) => setTitle(e.target.value)} className={`${inputCls} mb-3`} placeholder={isObj ? "e.g. Become the lowest-cost clinker producer in the region" : isAct ? "e.g. Call the vendor to confirm delivery dates" : `Name this ${label.toLowerCase()}`} />

      {(isObj || isContainer) && <div className="mb-3 grid grid-cols-2 gap-3">
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{Object.keys(METRIC_BY_TYPE).map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
        <div><label className="mb-1 block text-xs font-medium text-slate-600">Deadline</label><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className={inputCls} /></div>
      </div>}

      {isObj && <div className="mb-3">
        <label className="mb-1 flex items-center gap-2 text-xs font-medium text-slate-600"><input type="checkbox" checked={hasMetric} onChange={(e) => setHasMetric(e.target.checked)} /> North-star metric</label>
        {hasMetric && <div className="grid grid-cols-3 gap-2 text-xs"><div><label className="mb-1 block text-slate-500">Metric</label><input value={metric} onChange={(e) => setMetric(e.target.value)} className={inputCls} placeholder="Cost / tonne" /></div><div><label className="mb-1 block text-slate-500">Target</label><input type="number" value={target} onChange={(e) => setTarget(e.target.value)} className={inputCls} /></div><div><label className="mb-1 block text-slate-500">Unit</label><input value={unit} onChange={(e) => setUnit(e.target.value)} className={inputCls} /></div></div>}
      </div>}

      {isAct && <div className="mb-3 grid grid-cols-3 gap-2"><div><label className="mb-1 block text-xs font-medium text-slate-600">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-xs font-medium text-slate-600">Type</label><select value={actType} onChange={(e) => setActType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div><div><label className="mb-1 block text-xs font-medium text-slate-600">Date</label><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={inputCls} /></div></div>}

      <div className="mb-3 rounded-lg border border-violet-200 bg-violet-50 p-3">
        <div className="mb-2 flex items-center gap-1.5 text-xs font-medium text-violet-800"><Sparkles size={14} /> Describe it, dictate, or attach a PDF — AI fills the fields{isContainer ? " and a starter breakdown" : ""}</div>
        <MultiModalInput value={text} onChange={setText} onPdf={setPdf} placeholder="Speak or paste the note, or attach the doc…" />
        <button onClick={draftIt} disabled={busy || (!text.trim() && !pdf)} className={`${btnViolet} mt-2 w-full`}>{busy === "draft" ? <><Loader2 size={15} className="animate-spin" /> Reading…</> : <><Sparkles size={15} /> Draft with AI</>}</button>
      </div>

      {subs.length > 0 && <div className="mb-3"><div className="mb-1 text-xs font-medium text-slate-500">{level === "subwork" ? "Activities" : "Sub-works"} AI drafted ({subs.length})</div><div className="max-h-40 space-y-1 overflow-y-auto">{subs.map((c, i) => <div key={i} className="flex items-center gap-2 rounded-md bg-slate-50 px-3 py-1.5 text-sm text-slate-700"><span className="min-w-0 flex-1 truncate">{c.title}</span>{c.activities ? <span className="shrink-0 text-xs text-slate-400">{c.activities.length} activities</span> : null}<button onClick={() => setSubs(subs.filter((_, j) => j !== i))} className="shrink-0 text-slate-300 hover:text-rose-500"><X size={13} /></button></div>)}</div></div>}

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
    const nw = []; (ops || []).forEach((op) => { if (op.op === "add_subwork") nw.push({ id: nid("w"), parentId: work.id, level: "subwork", title: op.title, type: work.type, ownerId: work.ownerId }); });
    const liveSubs = subs.concat(nw); const na = [];
    (ops || []).forEach((op) => { if (op.op === "add_activity") { const sw = liveSubs.find((s) => s.title.toLowerCase().includes((op.subwork || "").toLowerCase())) || liveSubs[0]; if (sw) na.push({ id: nid("a"), workId: sw.id, title: op.title, assigneeId: null, date: null, status: "planned", plannedHrs: Number(op.estimateHrs) || 2, actualHrs: null, actType: op.type || "self", unplanned: true }); } });
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
      {ops && <div className="mt-3"><div className="mb-2 text-xs font-medium text-slate-500">Proposed changes</div><div className="space-y-1">{ops.length === 0 && <div className="text-sm text-slate-400">No changes proposed.</div>}{ops.map((op, i) => <div key={i} className="rounded-md bg-slate-50 px-3 py-1.5 text-sm text-slate-700">{op.op === "add_activity" ? `Add “${op.title}” to ${op.subwork}` : op.op === "add_subwork" ? `Add sub-work “${op.title}”` : `Retype “${op.match}” → ${op.type}`}</div>)}</div><div className="mt-3 flex gap-2"><button onClick={apply} className={`${btnDark} flex-1`}><Check size={14} /> Apply</button><button onClick={() => setOps(null)} className={btnLight}>Redo</button></div></div>}
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
      <div className="mb-4 grid grid-cols-3 gap-2 text-xs"><div><label className="mb-1 block text-slate-500">Under</label><select value={sw} onChange={(e) => setSw(e.target.value)} className={inputCls}>{subs.map((s) => <option key={s.id} value={s.id}>{s.title.slice(0, 18)}</option>)}</select></div><div><label className="mb-1 block text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div><div><label className="mb-1 block text-slate-500">Type</label><select value={type} onChange={(e) => setType(e.target.value)} className={inputCls}>{ACT_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</select></div></div>
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
    store.addCr({ id: nid("cr"), workId: activity.workId, subworkId: activity.workId, proposerId: me.id, kind, desc: desc || "(no note)", status: "pending", payload });
    flash("Proposed. It now needs organizer approval."); onClose();
  };
  const kinds = [["add_activity", "Add follow-up"], ["extend", "Needs more time"], ["blocked", "Blocked / help"], ["reassign", "Change owner"]];
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Propose a plan change</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">On “{activity.title}”. Your organizer approves before it changes the plan.</p>
      <div className="mb-3 flex flex-wrap gap-1">{kinds.map(([k, l]) => <button key={k} onClick={() => pick(k)} className={`rounded-md px-2.5 py-1.5 text-xs font-medium ${kind === k ? "bg-slate-900 text-white" : "border border-slate-200 text-slate-500 hover:bg-slate-50"}`}>{l}</button>)}</div>
      {kind === "add_activity" && <div className="mb-3 grid grid-cols-3 gap-2 text-xs"><div className="col-span-2"><label className="mb-1 block text-slate-500">New activity</label><input value={title} onChange={(e) => setTitle(e.target.value)} className={inputCls} placeholder="e.g. Re-run for missed depots" /></div><div><label className="mb-1 block text-slate-500">Hours</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={inputCls} /></div></div>}
      {kind === "extend" && <div className="mb-3 text-xs"><label className="mb-1 block text-slate-500">Extra hours needed</label><input type="number" value={hrs} onChange={(e) => setHrs(Number(e.target.value))} className={`${inputCls} w-28`} /></div>}
      {kind === "reassign" && <div className="mb-3 text-xs"><label className="mb-1 block text-slate-500">Hand over to</label><select value={to} onChange={(e) => setTo(e.target.value)} className={inputCls}><option value="">Pick a person</option>{USERS.filter((u) => u.id !== me.id).map((u) => <option key={u.id} value={u.id}>{u.name} — {u.title}</option>)}</select></div>}
      {kind === "blocked" && <div className="mb-3 rounded-md bg-rose-50 px-3 py-2 text-xs text-rose-700">Flags the activity as blocked so your organizer sees it in “Needs attention”.</div>}
      <label className="mb-1 block text-xs text-slate-500">Why</label><textarea value={desc} onChange={(e) => setDesc(e.target.value)} rows={2} className={`${inputCls} mb-3`} placeholder="Reason for the change…" />
      <button onClick={submit} disabled={kind === "reassign" && !to} className={`${btnDark} w-full`}>Submit for approval</button>
    </Modal>
  );
}
function Deliverable({ activity, work, store, busy, setBusy, flash, onClose }) {
  const [name, setName] = useState(activity.deliverable ? activity.deliverable.name : ""); const [content, setContent] = useState(activity.deliverable ? activity.deliverable.content : ""); const [result, setResult] = useState(activity.deliverable || null); const [reading, setReading] = useState(false);
  const readF = async (f) => {
    if (!f) return; setName(f.name); const n = f.name.toLowerCase();
    if (n.endsWith(".txt") || n.endsWith(".md") || f.type.startsWith("text")) { const r = new FileReader(); r.onload = () => setContent(String(r.result || "").slice(0, 8000)); r.readAsText(f); }
    else if (n.endsWith(".pdf") || n.endsWith(".docx") || n.endsWith(".doc")) { setReading(true); try { const b64 = await fileToB64(f); const text = await api.aiExtract(b64, f.name); setContent(text || ""); } catch (e) { flash(e.message || "Couldn't read that document."); } setReading(false); }
  };
  const score = async () => { setBusy("score"); let out; try { out = await AI.score(work ? work.title : "", activity.title, content || name); } catch { out = { score: content.length > 200 ? 72 : 55, verdict: "Reasonable draft", feedback: "AI scoring unavailable; provisional score." }; } const d = { name: name || "deliverable", content, ...out }; setResult(d); store.patchAct(activity.id, { deliverable: d }); setBusy(null); flash(`Deliverable scored ${out.score}/100.`); };
  return (
    <Modal onClose={onClose}>
      <div className="mb-3 flex items-center justify-between"><h3 className="text-sm font-medium text-slate-900">Deliverable — “{activity.title}”</h3><button onClick={onClose} className="text-slate-400"><X size={18} /></button></div>
      <p className="mb-3 text-xs text-slate-400">Attach the output (PO, email, PPT, doc). AI checks fit to the work. PDF / Word / .md / .txt are read automatically; for other files paste the content or a summary.</p>
      <label className="mb-2 flex cursor-pointer items-center gap-2 rounded-md border border-dashed border-slate-300 px-3 py-2 text-xs text-slate-500 hover:bg-slate-50"><Upload size={14} /> {reading ? <span className="inline-flex items-center gap-1"><Loader2 size={13} className="animate-spin" /> Reading {name}…</span> : (name || "Choose a file — PDF, Word, .md, .txt")}<input type="file" accept=".pdf,.docx,.doc,.txt,.md,text/plain,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document" className="hidden" onChange={(e) => readF(e.target.files && e.target.files[0])} /></label>
      <label className="mb-1 block text-xs text-slate-500">Content / summary</label><textarea value={content} onChange={(e) => setContent(e.target.value)} rows={4} className={`${inputCls} mb-3`} placeholder="Paste the deliverable text or a summary…" />
      <button onClick={score} disabled={busy || (!content.trim() && !name)} className={`${btnViolet} w-full`}>{busy === "score" ? <><Loader2 size={15} className="animate-spin" /> Scoring…</> : <><Sparkles size={15} /> Score with AI</>}</button>
      {result && <div className="mt-4 rounded-lg border border-slate-200 bg-slate-50 p-3"><div className="flex items-center gap-2"><div className="font-mono text-2xl font-medium text-violet-700">{result.score}<span className="text-sm text-slate-400">/100</span></div><div className="text-sm font-medium text-slate-700">{result.verdict}</div></div><div className="mt-1 text-sm text-slate-600">{result.feedback}</div></div>}
    </Modal>
  );
}
