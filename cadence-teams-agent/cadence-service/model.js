// Pure domain logic for Cadence — meters, RAG, sufficiency, scoping, capacity,
// roll-ups, and team auto-assignment. This is a VERBATIM port of the portal
// prototype's logic so the portal and the Teams agent compute results
// identically off the same 5-level tree.
const { addDays, iso, TODAY, MSD } = require('./seed');

const parseISO = (s) => (s ? new Date(s + 'T00:00:00') : null);
const isOverdue = (a) => a.date && parseISO(a.date) < TODAY && a.status !== 'executed';

/* ---------- meters + stats + rag ---------- */
function subtreeIds(works, id) {
  const out = [id];
  works.filter((w) => w.parentId === id).forEach((k) => out.push(...subtreeIds(works, k.id)));
  return out;
}
function computeMeters(works, acts, topId) {
  const ids = subtreeIds(works, topId);
  const nodes = works.filter((w) => ids.includes(w.id));
  const leaves = nodes.filter((n) => !works.some((w) => w.parentId === n.id));
  const my = acts.filter((a) => ids.includes(a.workId) && a.status !== 'cancelled');
  const top = works.find((w) => w.id === topId);
  const definition = ([top.title, top.ownerId, top.type, top.result].filter(Boolean).length / 4) * 100;
  const decomposition = leaves.length ? (leaves.filter((l) => my.some((a) => a.workId === l.id)).length / leaves.length) * 100 : 0;
  const scheduling = my.length ? (my.filter((a) => a.assigneeId && a.date).length / my.length) * 100 : 0;
  const executed = my.filter((a) => a.status === 'executed');
  const plannedTot = my.reduce((s, a) => s + a.plannedHrs, 0);
  const completion = plannedTot ? (executed.reduce((s, a) => s + a.plannedHrs, 0) / plannedTot) * 100 : 0;
  const planning = Math.round((definition + decomposition + scheduling) / 3);
  const execution = Math.round(completion);
  const behind = my.some((a) => isOverdue(a));
  const childWorks = works.filter((w) => w.parentId === topId);
  let stuck = null, low = 101;
  childWorks.forEach((c) => { const m = computeMeters(works, acts, c.id); if (m.completion < 100 && m.completion < low) { low = m.completion; stuck = c.title; } });
  let resultPct = null;
  if (top.result && top.result.target !== top.result.baseline) resultPct = Math.max(0, Math.min(100, ((top.result.current - top.result.baseline) / (top.result.target - top.result.baseline)) * 100));
  return { definition, decomposition, scheduling, completion, planning, execution, behind, stuck, resultPct };
}
function workStats(works, acts, topId) {
  const ids = subtreeIds(works, topId);
  const a = acts.filter((x) => ids.includes(x.workId) && x.status !== 'cancelled');
  const done = a.filter((x) => x.status === 'executed').length;
  const scheduled = a.filter((x) => x.assigneeId && x.date).length;
  const overdue = a.filter((x) => isOverdue(x)).length;
  const unscheduled = a.length - scheduled;
  const team = [...new Set(a.map((x) => x.assigneeId).filter(Boolean))];
  const deliv = a.filter((x) => x.deliverable).length;
  const blocked = a.filter((x) => x.blocked && x.status !== 'executed').length;
  const nextDue = a.filter((x) => x.date && x.status !== 'executed').map((x) => x.date).sort()[0] || null;
  return { total: a.length, done, scheduled, unscheduled, overdue, blocked, team, deliv, nextDue };
}
function ragOf(m, st) { if (st.overdue > 0) return 'red'; if (m.stuck || m.planning - m.execution > 40) return 'amber'; return 'green'; }
function subtreeAssignees(works, acts, topId) { const ids = subtreeIds(works, topId); return [...new Set(acts.filter((a) => ids.includes(a.workId) && a.assigneeId).map((a) => a.assigneeId))]; }

const USERS_REF = { list: [] }; // set by store so fnOf works without circular seed import
const setUsers = (users) => { USERS_REF.list = users; };
const fnOf = (id) => USERS_REF.list.find((u) => u.id === id)?.fn;

function visibleTops(works, acts, user) {
  const tops = works.filter((w) => !w.parentId);
  if (!user || user.level === 'md') return tops;
  if (user.level === 'vp') return tops.filter((w) => w.ownerId === user.id || fnOf(w.ownerId) === user.fn || subtreeAssignees(works, acts, w.id).some((uid) => fnOf(uid) === user.fn));
  return tops.filter((w) => w.ownerId === user.id || subtreeIds(works, w.id).some((id) => acts.some((a) => a.workId === id && a.assigneeId === user.id)));
}
// sufficiency returns just the label (the portal keeps a colour too; the API
// only needs the verdict string).
function sufficiency(m, st) {
  if (st.overdue > 0) return 'Behind';
  if (m.planning < 70) return 'Under-planned';
  if (m.resultPct != null && m.execution - m.resultPct > 25) return 'Output ahead of result';
  return 'Sufficient';
}
const loadOf = (acts, uid) => acts.filter((a) => a.assigneeId === uid && a.status !== 'executed' && a.status !== 'cancelled').reduce((s, a) => s + a.plannedHrs, 0);

/* ---------- tree / roll-up helpers ---------- */
const HOME_LEVEL = { md: 'objective', vp: 'initiative', member: 'subwork' };
function subtreeActs(works, acts, id) { const ids = subtreeIds(works, id); return acts.filter((a) => ids.includes(a.workId) && a.status !== 'cancelled'); }
function nodeRag(works, acts, id) { const sa = subtreeActs(works, acts, id); if (sa.some((a) => isOverdue(a) || a.blocked)) return 'red'; const m = computeMeters(works, acts, id); if (m.stuck || m.planning - m.execution > 40) return 'amber'; return 'green'; }
function deepestIssue(works, acts, id) { const sa = subtreeActs(works, acts, id); return sa.find((a) => a.blocked) || sa.find((a) => isOverdue(a)) || null; }
function attentionCount(works, acts, id) { const sa = subtreeActs(works, acts, id); return { blocked: sa.filter((a) => a.blocked).length, overdue: sa.filter((a) => isOverdue(a)).length }; }
function homeNodes(works, acts, user) {
  const lvl = HOME_LEVEL[user.level]; const nodes = works.filter((w) => w.level === lvl);
  if (user.level === 'md') return nodes;
  if (user.level === 'vp') return nodes.filter((w) => fnOf(w.ownerId) === user.fn || w.ownerId === user.id || subtreeAssignees(works, acts, w.id).some((uid) => fnOf(uid) === user.fn));
  return nodes.filter((w) => w.ownerId === user.id || subtreeIds(works, w.id).some((id) => acts.some((a) => a.workId === id && a.assigneeId === user.id)));
}
function scopeInitiatives(works, acts, user) { const homes = homeNodes(works, acts, user); const ids = new Set(homes.flatMap((h) => subtreeIds(works, h.id))); return works.filter((w) => w.level === 'initiative' && ids.has(w.id)); }

/* ---------- team auto-assignment (used by bot-side plan_initiative) ---------- */
// Distribute a fresh set of activities across a team, balanced by current open
// load, with dates spread from today toward the deadline. Mutates `activities`
// by setting assigneeId + date; returns them.
function assignToTeam(acts, activities, memberIds, deadline) {
  if (!memberIds || !memberIds.length) return activities;
  const load = {};
  memberIds.forEach((id) => (load[id] = loadOf(acts, id)));
  const dl = deadline ? parseISO(deadline) : null;
  const span = dl ? Math.max(1, Math.round((dl - TODAY) / MSD)) : 5;
  let i = 0;
  for (const act of activities) {
    const pick = memberIds.reduce((a, b) => (load[a] <= load[b] ? a : b));
    load[pick] += act.plannedHrs || 2;
    act.assigneeId = pick;
    act.date = iso(addDays(TODAY, i % span));
    i++;
  }
  return activities;
}

module.exports = {
  parseISO, isOverdue, subtreeIds, computeMeters, workStats, ragOf, sufficiency,
  subtreeAssignees, visibleTops, loadOf, nodeRag, deepestIssue, attentionCount,
  subtreeActs, homeNodes, scopeInitiatives, fnOf, setUsers, assignToTeam, HOME_LEVEL,
};
