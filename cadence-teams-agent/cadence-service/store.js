// In-memory store for the shared Cadence service — the single source of truth
// that BOTH the Teams agent and the portal read AND write. Holds the raw 5-level
// tree; the portal renders it directly (via /api/snapshot) while the bot reads
// the summarized shapes below. Swap this module for Postgres/Prisma in
// production without changing the route layer.
const seed = require('./seed');
const M = require('./model');

let uid = 5000;
const nid = (p) => `${p}${++uid}`;

// deep-copy seeds so mutations don't touch the module-level constants
const db = {
  users: JSON.parse(JSON.stringify(seed.USERS)),
  teams: JSON.parse(JSON.stringify(seed.TEAMS)),
  works: JSON.parse(JSON.stringify(seed.WORKS)),
  acts: JSON.parse(JSON.stringify(seed.ACTIVITIES)),
  crs: JSON.parse(JSON.stringify(seed.CRS)),
  remarks: JSON.parse(JSON.stringify(seed.REMARKS)),
};
M.setUsers(db.users); // let model.fnOf resolve functions

const user = (id) => db.users.find((u) => u.id === id);
const uName = (id) => user(id)?.name || 'Unassigned';
const team = (id) => db.teams.find((t) => t.id === id);
const METRIC_BY_TYPE = {
  procurement: { metric: 'Units issued', unit: 'count' },
  cost: { metric: 'Cost saved', unit: '₹/unit' },
  onboarding: { metric: 'Cycle time', unit: 'days' },
  compliance: { metric: 'Coverage', unit: '%' },
  general: { metric: 'Completion', unit: '%' },
};

/* ---------- auth ---------- */
function login(username, password) {
  const u = db.users.find((x) => x.username === (username || '').trim() && x.password === password);
  if (!u) return null;
  const { password: _pw, ...safe } = u;
  return safe;
}

/* ---------- raw snapshot (portal reads the whole tree, scopes client-side) ---------- */
function snapshot() {
  return {
    users: db.users.map(({ password, ...u }) => u),
    teams: db.teams,
    works: db.works,
    acts: db.acts,
    crs: db.crs,
    remarks: db.remarks,
  };
}

/* ---------- shapers (bot cards) ---------- */
function summarizeInitiative(top) {
  const m = M.computeMeters(db.works, db.acts, top.id);
  const st = M.workStats(db.works, db.acts, top.id);
  return {
    id: top.id, title: top.title, type: top.type, level: top.level,
    ownerId: top.ownerId, ownerName: uName(top.ownerId), fn: M.fnOf(top.ownerId),
    teamId: top.teamId || null, teamName: top.teamId ? team(top.teamId)?.name : null,
    scope: top.scope || null, objective: top.objective || null, deadline: top.deadline || null,
    result: top.result || null,
    resultPct: m.resultPct == null ? null : Math.round(m.resultPct),
    planning: m.planning, execution: m.execution, stuck: m.stuck,
    rag: M.nodeRag(db.works, db.acts, top.id), sufficiency: M.sufficiency(m, st),
    stats: { done: st.done, total: st.total, scheduled: st.scheduled, unscheduled: st.unscheduled, overdue: st.overdue, blocked: st.blocked, deliverables: st.deliv, nextDue: st.nextDue, teamSize: st.team.length },
    gap: top.result ? Math.round((top.result.target - top.result.current) * 100) / 100 : null,
  };
}

// For the bot's initiative card: flatten every sub-work in the initiative's
// subtree (works hold sub-works which hold activities in the 5-level model).
function detailInitiative(top) {
  const base = summarizeInitiative(top);
  const ids = M.subtreeIds(db.works, top.id);
  const subs = db.works.filter((w) => ids.includes(w.id) && w.level === 'subwork').map((s) => {
    const m = M.computeMeters(db.works, db.acts, s.id);
    const activities = db.acts.filter((a) => a.workId === s.id).map((a) => ({
      id: a.id, title: a.title, assigneeId: a.assigneeId, assigneeName: a.assigneeId ? uName(a.assigneeId) : null,
      date: a.date, status: a.status, plannedHrs: a.plannedHrs, actType: a.actType,
      overdue: M.isOverdue(a), blocked: !!a.blocked, unplanned: !!a.unplanned,
      deliverable: a.deliverable ? { score: a.deliverable.score, verdict: a.deliverable.verdict } : null,
    }));
    return { id: s.id, title: s.title, ownerId: s.ownerId, ownerName: uName(s.ownerId), planning: m.planning, execution: m.execution, activities };
  });
  return { ...base, subworks: subs };
}

/* ---------- reads (scoped, for the bot) ---------- */
function getPortfolio(userId) {
  const u = user(userId) || user('u_vik');
  const inis = M.scopeInitiatives(db.works, db.acts, u);
  const initiatives = inis.map(summarizeInitiative);
  const attain = initiatives.filter((i) => i.resultPct != null);
  const iniIds = new Set(inis.map((t) => t.id));
  const tiles = {
    scope: u.level === 'md' ? 'Enterprise' : u.level === 'vp' ? `Function — ${u.fn}` : 'My work',
    initiatives: initiatives.length,
    onTrack: initiatives.filter((i) => i.rag === 'green').length,
    atRisk: initiatives.filter((i) => i.rag !== 'green').length,
    avgResult: attain.length ? Math.round(attain.reduce((s, i) => s + i.resultPct, 0) / attain.length) : null,
    overdueActivities: initiatives.reduce((s, i) => s + i.stats.overdue, 0),
    approvalsPending: db.crs.filter((c) => c.status === 'pending' && iniIds.has(c.workId)).length,
  };
  return { tiles, initiatives };
}

function getAttention(userId) {
  const u = user(userId) || user('u_vik');
  const homes = M.homeNodes(db.works, db.acts, u);
  const scope = new Set(homes.flatMap((t) => M.subtreeIds(db.works, t.id)));
  const items = [];
  db.acts.filter((a) => scope.has(a.workId) && M.isOverdue(a)).forEach((a) => items.push({ kind: 'overdue', title: a.title, detail: `overdue ${a.date}`, assignee: uName(a.assigneeId), activityId: a.id }));
  db.acts.filter((a) => scope.has(a.workId) && a.blocked && a.status !== 'executed').forEach((a) => items.push({ kind: 'blocked', title: a.title, detail: 'blocked — needs help', assignee: uName(a.assigneeId), activityId: a.id }));
  M.scopeInitiatives(db.works, db.acts, u).map(summarizeInitiative).filter((i) => i.stuck).forEach((i) => items.push({ kind: 'stuck', title: i.title, detail: `stuck at ${i.stuck}` }));
  return { items };
}

function getCapacity(userId) {
  const u = user(userId) || user('u_vik');
  const pool = u.level === 'md' ? db.users.filter((x) => x.level !== 'md') : db.users.filter((x) => x.reports_to === u.id || x.id === u.id);
  return { capacity: pool.map((x) => ({ id: x.id, name: x.name, fn: x.fn, openHours: M.loadOf(db.acts, x.id) })) };
}

function getApprovals(userId) {
  const u = user(userId) || user('u_vik');
  const inis = M.scopeInitiatives(db.works, db.acts, u);
  const ids = new Set(inis.map((t) => t.id));
  const KL = { add_activity: 'add follow-up', extend: 'needs more time', blocked: 'flag blocked', reassign: 'change owner', retype: 'reclassify' };
  const pending = db.crs.filter((c) => c.status === 'pending' && (u.level === 'md' || ids.has(c.workId))).map((c) => ({
    id: c.id, proposer: uName(c.proposerId), kind: c.kind, kindLabel: KL[c.kind] || c.kind, desc: c.desc,
    initiative: db.works.find((w) => w.id === c.workId)?.title, payload: c.payload,
  }));
  return { pending };
}

/* ---------- resolvers (name -> entity) ---------- */
const norm = (s) => (s || '').toLowerCase().trim();
const STOP = new Set(['the', 'a', 'an', 'to', 'of', 'for', 'and', 'in', 'on', 'our', 'my', 'this', 'that', 'all', 'status']);
const tokens = (s) => norm(s).replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter((w) => w.length >= 3 && !STOP.has(w));
// Best fuzzy match by shared significant words (used when exact / substring fail),
// so the agent's paraphrase ("laptop refresh") still finds "Replace retired laptops".
function bestTokenMatch(list, q, titleOf) {
  const qt = tokens(q); if (!qt.length) return null;
  let best = null, bestScore = 0;
  for (const item of list) {
    const it = tokens(titleOf(item));
    // a query word hits if it shares a stem with any title word (handles plurals
    // and paraphrase: "laptop" ↔ "laptops", "logistics" ↔ "logistic")
    const score = qt.reduce((s, w) => s + (it.some((t) => t === w || t.includes(w) || w.includes(t)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = item; }
  }
  return bestScore > 0 ? best : null;
}
function resolveUser(q) { const n = norm(q); return db.users.find((u) => norm(u.name) === n) || db.users.find((u) => norm(u.name).includes(n) || norm(u.email) === n || norm(u.title).includes(n)) || bestTokenMatch(db.users, q, (u) => u.name); }
function resolveTeam(q) { const n = norm(q); return db.teams.find((t) => norm(t.name) === n) || db.teams.find((t) => norm(t.name).includes(n)) || bestTokenMatch(db.teams, q, (t) => t.name); }
function resolveInitiative(q) { const n = norm(q); const inis = db.works.filter((w) => w.level === 'initiative'); return inis.find((w) => norm(w.title) === n) || inis.find((w) => norm(w.title).includes(n)) || db.works.find((w) => w.id === q) || bestTokenMatch(inis, q, (w) => w.title); }
function resolveActivity(q) { const n = norm(q); return db.acts.find((a) => norm(a.title) === n) || db.acts.find((a) => norm(a.title).includes(n)) || db.acts.find((a) => a.id === q) || bestTokenMatch(db.acts, q, (a) => a.title); }

/* ---------- generic writes (portal primitives) ---------- */
// Adds accept client-provided ids so the portal's optimistic update and the
// server stay in sync with no duplication.
function addWorks(works) { const list = Array.isArray(works) ? works : [works]; list.forEach((w) => { if (!db.works.some((x) => x.id === w.id)) db.works.push(w); }); return list; }
function patchWork(id, patch) { const w = db.works.find((x) => x.id === id); if (!w) return null; Object.assign(w, patch); return w; }
function deleteWork(id) { const ids = M.subtreeIds(db.works, id); db.works = db.works.filter((w) => !ids.includes(w.id)); db.acts = db.acts.filter((a) => !ids.includes(a.workId)); return { removed: ids }; }
function addActs(acts) { const list = Array.isArray(acts) ? acts : [acts]; list.forEach((a) => { if (!db.acts.some((x) => x.id === a.id)) db.acts.push(a); }); return list; }
function patchAct(id, patch) { const a = db.acts.find((x) => x.id === id); if (!a) return null; Object.assign(a, patch); return a; }
function deleteAct(id) { const before = db.acts.length; db.acts = db.acts.filter((a) => a.id !== id); return before !== db.acts.length; }
function addCr(cr) { if (!db.crs.some((x) => x.id === cr.id)) db.crs.push(cr); return cr; }
function patchCr(id, patch) { const c = db.crs.find((x) => x.id === id); if (!c) return null; Object.assign(c, patch); return c; }
function addTeam(t) { if (!db.teams.some((x) => x.id === t.id)) db.teams.push(t); return t; }
function patchTeam(id, patch) { const t = db.teams.find((x) => x.id === id); if (!t) return null; Object.assign(t, patch); return t; }
function addRemark(r) { db.remarks.unshift(r); return r; }
function markRemarksRead(userId) { db.remarks.forEach((r) => { if (r.toIds.includes(userId) && !r.readBy.includes(userId)) r.readBy.push(userId); }); return db.remarks; }

/* ---------- higher-level writes (bot uses these) ---------- */
// An initiative must live under an objective so it shows in the objective-
// organized MD/VP portfolio. When the bot doesn't name one, home it under a
// single reusable "New initiatives" objective (owned by the MD).
function ensureHomeObjective() {
  let obj = db.works.find((w) => w.level === 'objective' && w.title === 'New initiatives');
  if (!obj) { obj = { id: nid('w'), parentId: null, level: 'objective', title: 'New initiatives', type: 'general', ownerId: 'u_vik' }; db.works.push(obj); }
  return obj.id;
}
function createInitiative({ ownerId, title, type = 'general', objective, deadline, teamId, parentId, subworks = [] }) {
  const topId = nid('w');
  const tpl = METRIC_BY_TYPE[type] || METRIC_BY_TYPE.general;
  const t = teamId ? team(teamId) : null;
  const memberIds = t ? t.memberIds : [];
  const parentObj = (parentId && db.works.some((w) => w.id === parentId && w.level === 'objective')) ? parentId : ensureHomeObjective();
  const top = { id: topId, parentId: parentObj, level: 'initiative', title, type, ownerId: ownerId || 'u_vik', teamId: teamId || null, scope: memberIds.length > 1 ? 'group' : 'individual', objective: objective || null, deadline: deadline || null, result: { metric: tpl.metric, unit: tpl.unit, baseline: 0, target: 100, current: 0 } };
  db.works.push(top);
  const newActs = [];
  subworks.forEach((sw) => {
    const sid = nid('w');
    db.works.push({ id: sid, parentId: topId, level: 'subwork', title: sw.title, type, ownerId: ownerId || 'u_vik' });
    (sw.activities || []).forEach((ac) => {
      newActs.push({ id: nid('a'), workId: sid, title: ac.title, assigneeId: null, date: null, status: 'planned', plannedHrs: Number(ac.estimateHrs) || 2, actualHrs: null, actType: ac.type || 'self' });
    });
  });
  M.assignToTeam(db.acts, newActs, memberIds, deadline);
  db.acts.push(...newActs);
  return detailInitiative(top);
}

function scheduleActivity(activityId, { assigneeId, date }) {
  const a = db.acts.find((x) => x.id === activityId);
  if (!a) return null;
  if (assigneeId !== undefined) a.assigneeId = assigneeId;
  if (date !== undefined) a.date = date;
  return a;
}

function reassignActivity(activityId, assigneeId) {
  const a = db.acts.find((x) => x.id === activityId);
  if (!a) return null;
  a.assigneeId = assigneeId;
  return a;
}

function decideApproval(crId, { approve, remark, spinoff, approverId }) {
  const cr = db.crs.find((c) => c.id === crId);
  if (!cr) return null;
  if (approve) {
    if (cr.kind === 'add_activity') db.acts.push({ id: nid('a'), workId: cr.subworkId, title: cr.payload.title, assigneeId: null, date: null, status: 'planned', plannedHrs: cr.payload.hrs, actualHrs: null, actType: cr.payload.type || 'self', unplanned: true });
    if (cr.kind === 'extend') { const a = db.acts.find((x) => x.id === cr.payload.activityId); if (a) a.plannedHrs += cr.payload.hrs || 1; }
    if (cr.kind === 'blocked') { const a = db.acts.find((x) => x.id === cr.payload.activityId); if (a) a.blocked = true; }
    if (cr.kind === 'reassign') { const a = db.acts.find((x) => x.id === cr.payload.activityId); if (a) a.assigneeId = cr.payload.to; }
    if (cr.kind === 'retype') { const a = db.acts.find((x) => x.id === cr.payload.activityId); if (a) a.actType = cr.payload.type; }
    let spun = null;
    if (spinoff && remark && remark.trim()) {
      // spin off a follow-up WORK under the initiative the CR belongs to
      let cur = db.works.find((w) => w.id === cr.workId); let g = 0; let ini = null;
      while (cur && g++ < 12) { if (cur.level === 'initiative') { ini = cur; break; } cur = db.works.find((w) => w.id === cur.parentId); }
      spun = { id: nid('w'), parentId: ini ? ini.id : null, level: 'work', title: remark.trim(), type: (ini && ini.type) || 'general', ownerId: (ini && ini.ownerId) || approverId || 'u_vik' };
      db.works.push(spun);
    }
    cr.status = 'approved'; cr.remark = remark || '';
    return { cr, spun };
  }
  cr.status = 'rejected'; cr.remark = remark || '';
  return { cr, spun: null };
}

function createTeam(name, memberIds) { const t = { id: nid('t'), name, memberIds }; db.teams.push(t); return t; }

module.exports = {
  db, nid, user, uName, team, METRIC_BY_TYPE,
  login, snapshot,
  summarizeInitiative, detailInitiative,
  getPortfolio, getAttention, getCapacity, getApprovals,
  resolveUser, resolveTeam, resolveInitiative, resolveActivity,
  addWorks, patchWork, deleteWork, addActs, patchAct, deleteAct, addCr, patchCr, addTeam, patchTeam, addRemark, markRemarksRead,
  createInitiative, scheduleActivity, reassignActivity, decideApproval, createTeam,
};
