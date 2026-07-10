// Cadence shared service — the REST contract that BOTH the Teams agent and the
// portal read from and write to. Single source of truth. Start: `npm start`
// (PORT 4000).
require('dotenv/config');
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const store = require('./store');
const ai = require('./ai');

// Safety net: never let a stray async rejection (e.g. a malformed /api/messages
// body) take down the whole merged process — log it and keep serving.
process.on('unhandledRejection', (reason) => console.error('Unhandled rejection:', reason));

const app = express();
app.use(cors());
app.use(express.json({ limit: '30mb' })); // room for base64 docs / audio

// light request log
app.use((req, _res, next) => { console.log(`${req.method} ${req.url}`); next(); });

// Optional shared-secret guard (set CADENCE_API_KEY to enable). Only applies to
// /api/* — the portal's own static files are public. Login, health, AI, and the
// bot's /api/messages (which carries its own Bot Framework auth) stay open.
app.use((req, res, next) => {
  const key = process.env.CADENCE_API_KEY;
  if (!key) return next();
  if (!req.path.startsWith('/api')) return next();
  if (['/api/health', '/api/login', '/api/messages'].includes(req.path) || req.path.startsWith('/api/ai')) return next();
  if (req.get('x-api-key') === key) return next();
  return res.status(401).json({ error: 'unauthorized' });
});

const ok = (res, data) => res.json(data);
const bad = (res, msg, code = 400) => res.status(code).json({ error: msg });
// Every write returns the fresh full snapshot so the portal reconciles to truth.
const withSnap = (res, extra = {}) => res.json({ ...extra, snapshot: store.snapshot() });

app.get('/api/health', (_req, res) => ok(res, { status: 'ok', foundry: ai.hasFoundry, deepgram: !!process.env.DEEPGRAM_API_KEY, time: new Date().toISOString() }));

/* ---------- auth ---------- */
app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  const u = store.login(username, password);
  if (!u) return bad(res, 'Wrong username or password.', 401);
  ok(res, { user: u });
});

/* ---------- raw snapshot (portal) ---------- */
app.get('/api/snapshot', (_req, res) => ok(res, store.snapshot()));

/* ---------- directory ---------- */
app.get('/api/users', (_req, res) => ok(res, { users: store.snapshot().users }));
app.get('/api/teams', (_req, res) => ok(res, { teams: store.db.teams.map((t) => ({ ...t, members: t.memberIds.map(store.uName) })) }));
app.post('/api/teams', (req, res) => {
  const { name, memberIds, id } = req.body || {};
  if (!name || !Array.isArray(memberIds) || !memberIds.length) return bad(res, 'name and memberIds[] required');
  const team = id ? store.addTeam({ id, name, memberIds }) : store.createTeam(name, memberIds);
  withSnap(res, { team });
});
app.patch('/api/teams/:id', (req, res) => { const t = store.patchTeam(req.params.id, req.body || {}); if (!t) return bad(res, 'team not found', 404); withSnap(res, { team: t }); });

/* ---------- scoped reads (bot) ---------- */
app.get('/api/portfolio', (req, res) => ok(res, store.getPortfolio(req.query.userId)));
app.get('/api/attention', (req, res) => ok(res, store.getAttention(req.query.userId)));
app.get('/api/capacity', (req, res) => ok(res, store.getCapacity(req.query.userId)));
app.get('/api/approvals', (req, res) => ok(res, store.getApprovals(req.query.userId)));
app.get('/api/member-status', (req, res) => { const r = store.getMemberStatus(req.query.userId, req.query.q); if (r && r.error) return bad(res, r.error, 400); ok(res, r); });

/* ---------- level-aware detail + interactive card actions (bot) ---------- */
app.get('/api/objectives', (req, res) => ok(res, store.getObjectives(req.query.userId)));
app.get('/api/objectives/:id', (req, res) => { const d = store.detailObjective(req.params.id); if (!d) return bad(res, 'objective not found', 404); ok(res, d); });
app.get('/api/works/:id/detail', (req, res) => { const d = store.detailWork(req.params.id); if (!d) return bad(res, 'work not found', 404); ok(res, d); });
app.post('/api/works/:id/deliverables/:did/toggle', (req, res) => { const w = store.toggleDeliverable(req.params.id, req.params.did); if (!w) return bad(res, 'not found', 404); withSnap(res, { work: store.detailWork(w.id) }); });
app.post('/api/works/:id/auto-assign', (req, res) => { const r = store.autoAssignWork(req.params.id); if (!r.work) return bad(res, 'work not found', 404); withSnap(res, { assigned: r.assigned, work: store.detailWork(r.work.id) }); });
app.post('/api/works/:id/deliverables/suggest', async (req, res) => {
  const w = store.getWork(req.params.id); if (!w) return bad(res, 'work not found', 404);
  const acts = store.db.acts.filter((a) => a.workId === w.id).map((a) => a.title);
  let items = [];
  try { const out = await ai.suggestDeliverables(w.title, acts); items = out.deliverables || []; } catch (e) { return aiErr(res, e); }
  const r = store.addWorkDeliverables(w.id, items);
  withSnap(res, { added: r.added, work: store.detailWork(w.id) });
});

app.get('/api/resolve', (req, res) => {
  const { kind, q } = req.query;
  let e = null;
  if (kind === 'user') e = store.resolveUser(q);
  else if (kind === 'team') e = store.resolveTeam(q);
  else if (kind === 'initiative') e = store.resolveInitiative(q);
  else if (kind === 'activity') e = store.resolveActivity(q);
  else if (kind === 'work') e = store.resolveWork(q);
  else if (kind === 'objective') e = store.resolveObjective(q);
  ok(res, { match: e ? { id: e.id, name: e.name || e.title } : null });
});

app.get('/api/initiatives', (req, res) => ok(res, { initiatives: store.getPortfolio(req.query.userId).initiatives }));
app.get('/api/initiatives/:id', (req, res) => {
  const top = store.resolveInitiative(req.params.id);
  if (!top) return bad(res, 'initiative not found', 404);
  ok(res, store.detailInitiative(top));
});
app.post('/api/initiatives', (req, res) => {
  const { ownerId, title, type, objective, deadline, teamId, parentId, works, subworks } = req.body || {};
  if (!title) return bad(res, 'title required');
  const initiative = store.createInitiative({ ownerId, title, type, objective, deadline, teamId, parentId, works: works || subworks || [] });
  withSnap(res, { initiative });
});

/* ---------- generic work / activity writes (portal primitives) ---------- */
app.post('/api/works', (req, res) => { const works = store.addWorks(req.body?.works || req.body); withSnap(res, { works }); });
app.patch('/api/works/:id', (req, res) => { const w = store.patchWork(req.params.id, req.body || {}); if (!w) return bad(res, 'work not found', 404); withSnap(res, { work: w }); });
app.delete('/api/works/:id', (req, res) => { withSnap(res, store.deleteWork(req.params.id)); });

/* ---------- deliverables (work-level checklist; bot logs files here) ---------- */
// Read a work's checklist by id or title.
app.get('/api/deliverables', (req, res) => { const d = store.workDeliverables(req.query.work); if (!d) return bad(res, `No work matches "${req.query.work}".`, 404); ok(res, d); });
// Attach (and optionally AI-score) a file against one checklist item. Extraction
// + scoring run server-side so the bot only forwards the raw file.
app.post('/api/works/:id/deliverables/attach', async (req, res) => {
  const { label, deliverableId, kind, fileBase64, fileName, score, create } = req.body || {};
  const w = store.getWork(req.params.id);
  if (!w) return bad(res, 'work not found', 404);
  let content = '', scored = null;
  if (fileBase64) { try { content = await ai.extract(fileBase64, fileName); } catch { /* unreadable file — still record the attachment */ } }
  if (score !== false && content && content.trim()) { try { scored = await ai.scoreDeliverable(w.title, label, content); } catch { /* scoring optional */ } }
  const r = store.attachDeliverable(w.id, { label, deliverableId, kind, create, file: fileName ? { name: fileName } : null, score: scored ? scored.score : undefined, verdict: scored ? scored.verdict : undefined, feedback: scored ? scored.feedback : undefined });
  if (!r) return bad(res, 'work not found', 404);
  if (r.unmatched) return bad(res, `No deliverable on "${w.title}" matches "${label || ''}". Available: ${r.available.length ? r.available.join(', ') : 'none — pass create=true to add it'}.`);
  withSnap(res, { work: r.work, item: r.item, scored: !!scored });
});
// Mark a work (and its open activities) complete.
app.post('/api/works/:id/complete', (req, res) => { const r = store.completeWork(req.params.id); if (!r) return bad(res, 'work not found', 404); withSnap(res, r); });

app.post('/api/activities', (req, res) => { const acts = store.addActs(req.body?.activities || req.body); withSnap(res, { activities: acts }); });
app.patch('/api/activities/:id', (req, res) => { const a = store.patchAct(req.params.id, req.body || {}); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });
app.delete('/api/activities/:id', (req, res) => { const okd = store.deleteAct(req.params.id); if (!okd) return bad(res, 'activity not found', 404); withSnap(res, { deleted: req.params.id }); });

// bot-friendly aliases
app.post('/api/activities/:id/schedule', (req, res) => { const a = store.scheduleActivity(req.params.id, req.body || {}); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });
app.post('/api/activities/:id/reassign', (req, res) => { const { assigneeId } = req.body || {}; if (!assigneeId) return bad(res, 'assigneeId required'); const a = store.reassignActivity(req.params.id, assigneeId); if (!a) return bad(res, 'activity not found', 404); withSnap(res, { activity: a }); });

/* ---------- Outlook → Cadence import (reverse calendar sync) ----------
   The portal reads the signed-in user's week of meetings (client-side, with the
   Graph token) and posts them here. We skip anything already imported (by event
   id), ask the AI to file each new meeting under the best-matching work, and
   auto-create activities (unmatched ones land in the user's "unsorted" inbox). */
app.post('/api/calendar/import', async (req, res) => {
  try {
    const { userId, events } = req.body || {};
    if (!Array.isArray(events)) return bad(res, 'events[] required');
    const u = store.user(userId) || store.user('u_vik');
    const fresh = events.filter((ev) => ev && ev.id && !ev.isCancelled && !store.db.acts.some((a) => a.outlookEventId === ev.id));
    const matches = {};
    if (fresh.length && ai.hasFoundry) {
      const works = store.db.works.filter((w) => w.level === 'work' && !w.inbox).map((w) => {
        const ini = store.db.works.find((x) => x.id === w.parentId);
        return { id: w.id, title: w.title, initiative: ini ? ini.title : null };
      });
      try {
        const r = await ai.categorizeMeetings(fresh, works);
        (r.matches || []).forEach((m) => { if (m && typeof m.i === 'number' && fresh[m.i] && m.workId && works.some((w) => w.id === m.workId)) matches[fresh[m.i].id] = m.workId; });
      } catch { /* AI off / bad JSON → everything falls through as unsorted */ }
    }
    const imported = store.importOutlookEvents(u.id, events, matches);
    withSnap(res, { imported });
  } catch (e) { bad(res, e.message || String(e)); }
});

/* ---------- change requests / approvals ---------- */
app.post('/api/crs', (req, res) => { const cr = store.addCr(req.body || {}); withSnap(res, { cr }); });
app.patch('/api/crs/:id', (req, res) => { const c = store.patchCr(req.params.id, req.body || {}); if (!c) return bad(res, 'change request not found', 404); withSnap(res, { cr: c }); });
app.post('/api/approvals/:id/decide', (req, res) => { const out = store.decideApproval(req.params.id, req.body || {}); if (!out) return bad(res, 'change request not found', 404); withSnap(res, out); });

/* ---------- remarks / nudges (portal) ---------- */
app.post('/api/remarks', (req, res) => { const r = store.addRemark(req.body || {}); withSnap(res, { remark: r }); });
app.post('/api/remarks/read', (req, res) => { const { userId } = req.body || {}; store.markRemarksRead(userId); withSnap(res, {}); });

/* ---------- AI (shared Foundry model) ---------- */
const aiErr = (res, e) => { if (e && e.notConfigured) return res.status(501).json({ error: e.message, notConfigured: true }); return res.status(500).json({ error: e.message || String(e) }); };

app.post('/api/ai/complete', async (req, res) => {
  try { const text = await ai.complete(req.body?.system, req.body?.user); ok(res, { text }); }
  catch (e) { aiErr(res, e); }
});
app.post('/api/ai/extract', async (req, res) => {
  try { const text = await ai.extract(req.body?.data, req.body?.name); ok(res, { text, name: req.body?.name || 'document' }); }
  catch (e) { aiErr(res, e); }
});
app.post('/api/ai/transcribe', async (req, res) => {
  try { const text = await ai.transcribe(req.body?.audio, req.body?.mimetype); ok(res, { text }); }
  catch (e) { aiErr(res, e); }
});

/* ---------- single-App-Service extras: mount the bot + serve the portal ----------
   In production one Azure App Service runs this process, which serves the REST
   API (above), the Teams bot's /api/messages, AND the built portal SPA. Both
   extras are optional so local `npm run dev` (API only) still works: the bot is
   mounted only if it's been compiled, and the portal only if it's been built. */

// Mount the Teams bot in-process (needs teams-bot to be built → dist/host.js).
try {
  const { registerBot } = require('../teams-bot/dist/host');
  registerBot(app);
  console.log('Teams bot mounted at /api/messages');
} catch (e) {
  console.log(`Teams bot not mounted (build the bot to enable /api/messages): ${e.message}`);
}

// Serve the built portal (SPA). Must come AFTER the API routes so /api/* wins.
const PORTAL_DIST = path.join(__dirname, '..', '..', 'cadence-portal', 'dist');
if (fs.existsSync(PORTAL_DIST)) {
  app.use(express.static(PORTAL_DIST));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) return next(); // let unknown API paths 404 as JSON
    res.sendFile(path.join(PORTAL_DIST, 'index.html'));
  });
  console.log('Serving portal from', PORTAL_DIST);
} else {
  console.log('Portal build not found (run the portal build to serve the SPA from here).');
}

const PORT = process.env.PORT || 4000;
if (require.main === module) {
  app.listen(PORT, () => console.log(`Cadence server on http://localhost:${PORT}  (Foundry: ${ai.hasFoundry ? 'on' : 'off'})`));
}
module.exports = app;
