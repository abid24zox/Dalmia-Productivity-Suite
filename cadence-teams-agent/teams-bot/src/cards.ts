// Adaptive Card builders — the agent replies with these instead of walls of
// text. Cards are plain JSON (schema 1.4) wrapped by CardFactory in bot.ts.
//
// Design system (keep every card consistent & scannable):
//   • header(title, sub)   — bold title + subtle one-line context
//   • kpiStrip([...])       — a tinted band of "big number + small label" columns
//   • sectionHeader(text)   — an accent, separated section divider
//   • RAG is always colour-coded: green = on track, amber = at risk, red = behind
const PORTAL = process.env.PORTAL_URL || 'https://cadence-zoxima-productivity-suite-gubndqggbkgfbpda.centralindia-01.azurewebsites.net';

const ragColor = (rag: string) => (rag === 'red' ? 'Attention' : rag === 'amber' ? 'Warning' : 'Good');
const ragWord = (rag: string) => (rag === 'red' ? 'Behind' : rag === 'amber' ? 'At risk' : 'On track');
const plural = (n: number, one: string, many = one + 's') => `${n} ${n === 1 ? one : many}`;

const card = (body: any[], actions: any[] = []) => ({
  type: 'AdaptiveCard', $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', version: '1.4', body, ...(actions.length ? { actions } : {}),
});

function header(t: string, sub?: string) {
  return {
    type: 'Container', spacing: 'None', items: [
      { type: 'TextBlock', text: t, weight: 'Bolder', size: 'Large', wrap: true },
      ...(sub ? [{ type: 'TextBlock', text: sub, isSubtle: true, size: 'Small', spacing: 'None', wrap: true }] : []),
    ],
  };
}

// A tinted band of KPIs — each a big value over a small label.
function kpiStrip(kpis: { label: string; value: string; color?: string }[]) {
  return {
    type: 'Container', style: 'emphasis', spacing: 'Medium', items: [
      {
        type: 'ColumnSet', columns: kpis.map((k) => ({
          type: 'Column', width: 'stretch', items: [
            { type: 'TextBlock', text: k.value, size: 'Large', weight: 'Bolder', color: k.color || 'Default', horizontalAlignment: 'Center', spacing: 'None', wrap: false },
            { type: 'TextBlock', text: k.label, size: 'Small', isSubtle: true, horizontalAlignment: 'Center', spacing: 'None', wrap: true },
          ],
        })),
      },
    ],
  };
}

const sectionHeader = (t: string) => ({ type: 'TextBlock', text: t.toUpperCase(), weight: 'Bolder', size: 'Small', color: 'Accent', spacing: 'Medium', separator: true });
const runT = (text: string, o: any = {}) => ({ type: 'TextRun', text, ...o });
const rich = (inlines: any[], spacing = 'Small') => ({ type: 'RichTextBlock', spacing, inlines });
const note = (text: string) => ({ type: 'TextBlock', text, size: 'Small', isSubtle: true, wrap: true, spacing: 'Small' });
const levelTag = (t: string) => ({ type: 'TextBlock', text: t.toUpperCase(), size: 'Small', weight: 'Bolder', color: 'Accent', spacing: 'None' });

// A tappable row that drills to the level below (the whole row is a Submit).
function drillRow(title: string, metaInline: any, data: any, rightText?: string | null, rightColor?: string) {
  return {
    type: 'ColumnSet', spacing: 'Small', separator: true, selectAction: { type: 'Action.Submit', data },
    columns: [
      { type: 'Column', width: 'stretch', items: [
        { type: 'TextBlock', text: title, weight: 'Bolder', wrap: true, spacing: 'None' },
        metaInline,
      ] },
      { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [
        ...(rightText ? [{ type: 'TextBlock', text: rightText, color: rightColor || 'Default', weight: 'Bolder', horizontalAlignment: 'Right', spacing: 'None' }] : []),
        { type: 'TextBlock', text: '›', color: 'Accent', size: 'Large', weight: 'Bolder', horizontalAlignment: 'Right', spacing: 'None' },
      ] },
    ],
  };
}

// ---------- portfolio ----------
export function portfolioCard(p: any) {
  const t = p.tiles;
  const body: any[] = [
    header('Portfolio', `${t.scope} · ${plural(t.initiatives, 'initiative')}`),
    kpiStrip([
      { label: 'On track', value: `${t.onTrack}`, color: 'Good' },
      { label: 'At risk', value: `${t.atRisk}`, color: t.atRisk ? 'Attention' : 'Default' },
      { label: 'Avg result', value: t.avgResult == null ? '—' : `${t.avgResult}%` },
    ]),
    note(`${plural(t.overdueActivities, 'activity', 'activities')} overdue  ·  ${plural(t.approvalsPending, 'approval')} pending`),
    sectionHeader('Initiatives'),
  ];
  p.initiatives.forEach((i: any) => {
    body.push({
      type: 'Container', spacing: 'Small', separator: true, items: [
        { type: 'ColumnSet', columns: [
          { type: 'Column', width: 'stretch', items: [
            { type: 'TextBlock', text: i.title, weight: 'Bolder', wrap: true, spacing: 'None' },
            { type: 'TextBlock', text: `${i.ownerName} · ${i.fn}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true },
          ] },
          { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [
            { type: 'TextBlock', text: i.resultPct == null ? '—' : `${i.resultPct}%`, color: ragColor(i.rag), weight: 'Bolder', size: 'Large', horizontalAlignment: 'Right', spacing: 'None' },
            { type: 'TextBlock', text: 'result', isSubtle: true, size: 'Small', horizontalAlignment: 'Right', spacing: 'None' },
          ] },
        ] },
        rich([
          runT(ragWord(i.rag), { color: ragColor(i.rag), weight: 'Bolder' }),
          runT(`   Exec ${i.execution}%  ·  ${i.stats.done}/${i.stats.total} done`, { isSubtle: true }),
          ...(i.stats.overdue ? [runT(`  ·  ${i.stats.overdue} overdue`, { color: 'Attention' })] : []),
          ...(i.stuck ? [runT(`  ·  stuck at ${i.stuck}`, { isSubtle: true })] : []),
        ]),
      ],
    });
  });
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

// ---------- one initiative (detail) ----------
export function initiativeCard(i: any) {
  const body: any[] = [
    levelTag('Initiative'),
    header(i.title, `${i.ownerName} · ${i.fn}${i.teamName ? ' · ' + i.teamName : ''}`),
    kpiStrip([
      { label: 'Result', value: i.resultPct == null ? '—' : `${i.resultPct}%`, color: ragColor(i.rag) },
      { label: 'Execution', value: `${i.execution}%` },
      { label: 'Planning', value: `${i.planning}%` },
    ]),
    rich([runT(ragWord(i.rag), { color: ragColor(i.rag), weight: 'Bolder' }), runT(`   ${i.sufficiency}`, { isSubtle: true })], 'Small'),
    { type: 'FactSet', spacing: 'Small', facts: [
      ...(i.result ? [{ title: i.result.metric, value: `${i.result.current} of ${i.result.target} ${i.result.unit}  ·  gap ${i.gap} ${i.result.unit}` }] : []),
      { title: 'Timeline', value: `${i.startDate || '?'} → ${i.endDate || i.deadline || '?'}` },
      { title: 'Effort', value: i.effort ? `${i.effort.actual} / ${i.effort.planned} h  (actual / est)` : '—' },
    ] },
    sectionHeader('Work breakdown — tap to open'),
  ];
  (i.subworks || []).forEach((s: any) => {
    const d = s.deliverables;
    const ov = (s.activities || []).filter((a: any) => a.overdue).length;
    const word = s.completedAt ? 'Complete' : ov ? 'Behind' : 'In progress';
    const col = s.completedAt ? 'Good' : ov ? 'Attention' : 'Warning';
    body.push(drillRow(
      s.title,
      rich([
        runT(word, { color: col, weight: 'Bolder' }),
        runT(`   Exec ${s.execution}%${d && d.total ? ` · deliv ${d.done}/${d.total}` : ''}`, { isSubtle: true }),
        ...(ov ? [runT(`  ·  ${ov} overdue`, { color: 'Attention' })] : []),
      ], 'None'),
      { action: 'drill', kind: 'work', id: s.id },
    ));
  });
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open in portal', url: PORTAL }]);
}

// ---------- objectives report (tappable) ----------
export function objectivesReportCard(r: any) {
  const body: any[] = [
    header('Objectives', `${r.scope} · ${plural(r.objectives.length, 'objective')}`),
    kpiStrip([
      { label: 'On track', value: `${r.onTrack}`, color: 'Good' },
      { label: 'To watch', value: `${r.atRisk}`, color: r.atRisk ? 'Attention' : 'Default' },
      { label: 'Overdue', value: `${r.overdue}`, color: r.overdue ? 'Attention' : 'Default' },
    ]),
    sectionHeader('Objectives — tap to open'),
  ];
  r.objectives.forEach((o: any) => body.push(drillRow(
    o.title,
    rich([
      runT(ragWord(o.rag), { color: ragColor(o.rag), weight: 'Bolder' }),
      runT(`   ${o.initiativesCount} initiatives${o.avgResult != null ? ` · avg ${o.avgResult}%` : ''}`, { isSubtle: true }),
      ...(o.overdue ? [runT(`  ·  ${o.overdue} overdue`, { color: 'Attention' })] : []),
    ], 'None'),
    { action: 'drill', kind: 'objective', id: o.id },
  )));
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

// ---------- one objective (detail) ----------
export function objectiveCard(o: any) {
  const body: any[] = [
    levelTag('Objective'),
    header(o.title, o.ownerName),
    kpiStrip([
      { label: 'Status', value: ragWord(o.rag), color: ragColor(o.rag) },
      { label: 'Avg result', value: o.avgResult == null ? '—' : `${o.avgResult}%` },
      { label: 'Overdue', value: `${o.overdue}`, color: o.overdue ? 'Attention' : 'Default' },
    ]),
    sectionHeader(`Initiatives · ${o.initiativesCount} — tap to open`),
  ];
  (o.initiatives || []).forEach((i: any) => body.push(drillRow(
    i.title,
    rich([
      runT(ragWord(i.rag), { color: ragColor(i.rag), weight: 'Bolder' }),
      runT(`   ${i.ownerName} · Exec ${i.execution}%`, { isSubtle: true }),
      ...(i.overdue ? [runT(`  ·  ${i.overdue} overdue`, { color: 'Attention' })] : []),
    ], 'None'),
    { action: 'drill', kind: 'initiative', id: i.id },
    i.resultPct == null ? null : `${i.resultPct}%`, ragColor(i.rag),
  )));
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open in portal', url: PORTAL }]);
}

// ---------- one work (detail) — activities + interactive deliverables ----------
export function workCard(w: any) {
  const dn = w.deliverables.done, tt = w.deliverables.total, done = !!w.completedAt;
  const word = done ? 'Complete' : w.rag === 'red' ? 'Behind' : w.rag === 'amber' ? 'At risk' : 'On track';
  const body: any[] = [
    levelTag('Work'),
    header(w.title, `${w.ownerName}${w.parentTitle ? ' · ' + w.parentTitle : ''}`),
    kpiStrip([
      { label: 'Planned', value: `${w.planning}%` },
      { label: 'Done', value: `${w.execution}%`, color: done ? 'Good' : w.execution ? 'Default' : 'Warning' },
      { label: 'Overdue', value: `${w.overdue}`, color: w.overdue ? 'Attention' : 'Default' },
    ]),
    rich([runT(word, { color: done ? 'Good' : ragColor(w.rag), weight: 'Bolder' }), runT(`   ${w.doneActs}/${w.totalActs} activities done`, { isSubtle: true })], 'Small'),
    sectionHeader(`Activities · ${w.totalActs}`),
  ];
  w.activities.forEach((a: any) => body.push(rich([
    runT(a.overdue || a.blocked ? '●  ' : '○  ', { color: a.overdue ? 'Attention' : a.blocked ? 'Warning' : 'Default' }),
    runT(a.title),
    runT(`  —  ${a.assigneeName || 'Unassigned'} · ${a.date || 'no date'} · ${a.status}`, { isSubtle: true }),
    ...(a.overdue ? [runT('  · overdue', { color: 'Attention' })] : a.blocked ? [runT('  · blocked', { color: 'Warning' })] : []),
  ], 'Small')));
  body.push(sectionHeader(`Deliverables · ${dn}/${tt} delivered${w.deliverables.avgScore != null ? ` · avg ${w.deliverables.avgScore}` : ''}  —  tap to tick`));
  if (!w.deliverables.items.length) body.push({ type: 'TextBlock', text: 'None yet — use “Suggest deliverables”.', isSubtle: true, wrap: true, spacing: 'Small' });
  w.deliverables.items.forEach((it: any) => body.push({
    type: 'ColumnSet', spacing: 'Small', separator: true, selectAction: { type: 'Action.Submit', data: { action: 'deliv_toggle', workId: w.id, did: it.id } },
    columns: [
      { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [{ type: 'TextBlock', text: it.done ? '✅' : '⬜' }] },
      { type: 'Column', width: 'stretch', items: [
        { type: 'TextBlock', text: `${KIND_ICON[it.kind] || '📎'}  ${it.label}`, wrap: true, spacing: 'None', isSubtle: it.done },
        ...(it.file || it.verdict ? [{ type: 'TextBlock', text: `${it.file || ''}${it.file && it.verdict ? ' · ' : ''}${it.verdict || ''}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true }] : []),
      ] },
      { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [{ type: 'TextBlock', text: it.score != null ? `${it.score}` : '', color: 'Good', weight: 'Bolder', horizontalAlignment: 'Right' }] },
    ],
  }));
  const actions: any[] = [];
  if (!done) {
    actions.push({ type: 'Action.Submit', title: '✨ Suggest deliverables', data: { action: 'suggest_deliv', workId: w.id } });
    actions.push({ type: 'Action.Submit', title: 'Auto-assign', data: { action: 'auto_assign', workId: w.id } });
    actions.push({ type: 'Action.Submit', title: '✓ Mark complete', style: 'positive', data: { action: 'work_complete', workId: w.id } });
  }
  actions.push({ type: 'Action.OpenUrl', title: 'Open in portal', url: PORTAL });
  return card(body, actions);
}

// Inline confirm before completing a work (replaces the card in place).
export function confirmCompleteCard(w: any) {
  const open = w.totalActs - w.doneActs;
  return card([
    levelTag('Work'),
    header(w.title, w.parentTitle || ''),
    { type: 'Container', style: 'warning', spacing: 'Medium', items: [
      { type: 'TextBlock', text: `Mark this work${open > 0 ? ` and its ${open} open ${open === 1 ? 'activity' : 'activities'}` : ''} complete?`, wrap: true, weight: 'Bolder' },
      { type: 'TextBlock', text: `${w.deliverables.done}/${w.deliverables.total} deliverables are in.`, wrap: true, isSubtle: true, spacing: 'None' },
    ] },
  ], [
    { type: 'Action.Submit', title: '✓ Yes, complete it', style: 'positive', data: { action: 'work_complete_go', workId: w.id } },
    { type: 'Action.Submit', title: 'Cancel', data: { action: 'drill', kind: 'work', id: w.id } },
  ]);
}

// ---------- initiative created & assigned ----------
export function createdCard(i: any) {
  const subs = i.works || i.subworks || [];
  const nAct = subs.reduce((s: number, sw: any) => s + (sw.activities?.length || 0), 0);
  const body: any[] = [
    header('Initiative created & assigned ✅', i.title),
    kpiStrip([
      { label: 'Works', value: `${subs.length}` },
      { label: 'Activities', value: `${nAct}` },
    ]),
    note(`Team ${i.teamName || '—'}  ·  type ${i.type}  ·  deadline ${i.deadline || '—'}`),
    sectionHeader('The plan'),
  ];
  subs.forEach((s: any) => {
    body.push({
      type: 'Container', spacing: 'Small', separator: true, items: [
        { type: 'TextBlock', text: s.title, weight: 'Bolder', wrap: true, spacing: 'None' },
        ...(s.activities || []).map((a: any) => ({ type: 'TextBlock', size: 'Small', isSubtle: true, wrap: true, spacing: 'None', text: `○  ${a.title} — ${a.assigneeName || 'unassigned'} · ${a.date || 'tbd'}` })),
      ],
    });
  });
  return card(body, [{ type: 'Action.OpenUrl', title: 'View in portal', url: PORTAL }]);
}

// ---------- needs attention ----------
export function attentionCard(items: any[]) {
  if (!items.length) return card([header('Needs attention'), { type: 'TextBlock', text: '✅  All clear — nothing overdue, blocked, or stuck in your scope.', color: 'Good', wrap: true, spacing: 'Small' }]);
  const n = (k: string) => items.filter((i) => i.kind === k).length;
  const body: any[] = [
    header('Needs attention', `${plural(items.length, 'item')} to look at`),
    kpiStrip([
      { label: 'Overdue', value: `${n('overdue')}`, color: n('overdue') ? 'Attention' : 'Default' },
      { label: 'Blocked', value: `${n('blocked')}`, color: n('blocked') ? 'Warning' : 'Default' },
      { label: 'Stuck', value: `${n('stuck')}`, color: n('stuck') ? 'Warning' : 'Default' },
    ]),
  ];
  items.forEach((it) => {
    const col = it.kind === 'stuck' ? 'Warning' : 'Attention';
    body.push(rich([
      runT(`${String(it.kind).toUpperCase()}  `, { color: col, weight: 'Bolder', size: 'Small' }),
      runT(it.title, { weight: 'Bolder' }),
      runT(`  —  ${it.detail}${it.assignee ? `  (${it.assignee})` : ''}`, { isSubtle: true }),
    ], 'Small'));
  });
  return card(body);
}

// ---------- team capacity ----------
export function capacityCard(rows: any[]) {
  const body: any[] = [header('Team capacity', 'Open (not-yet-done) work per person')];
  [...rows].sort((a, b) => a.openHours - b.openHours).forEach((r) => {
    const tone = r.openHours === 0 ? 'Good' : r.openHours >= 30 ? 'Warning' : 'Default';
    body.push({
      type: 'ColumnSet', spacing: 'Small', separator: true, columns: [
        { type: 'Column', width: 'stretch', items: [
          { type: 'TextBlock', text: r.name, weight: 'Bolder', wrap: true, spacing: 'None' },
          { type: 'TextBlock', text: r.fn, isSubtle: true, size: 'Small', spacing: 'None' },
        ] },
        { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [
          { type: 'TextBlock', text: `${r.openHours}h`, weight: 'Bolder', color: tone, horizontalAlignment: 'Right', spacing: 'None' },
          { type: 'TextBlock', text: r.openHours === 0 ? 'free' : r.openHours >= 30 ? 'heavy' : 'open', isSubtle: true, size: 'Small', horizontalAlignment: 'Right', spacing: 'None' },
        ] },
      ],
    });
  });
  return card(body);
}

// ---------- one person's performance ----------
export function memberStatusCard(r: any) {
  if (!r || r.error) return textCard('Performance', r?.error || 'No matching person.');
  const u = r.user, a = r.activities, d = r.deliverables;
  const body: any[] = [
    header(u.name, `${u.title} · ${u.fn}`),
    kpiStrip([
      { label: 'Execution', value: `${a.execPct}%` },
      { label: 'Overdue', value: `${a.overdue}`, color: a.overdue ? 'Attention' : 'Default' },
      { label: 'Blocked', value: `${a.blocked}`, color: a.blocked ? 'Warning' : 'Default' },
      { label: 'Deliv. avg', value: d.avgScore != null ? `${d.avgScore}` : '—' },
    ]),
    note(`${a.done}/${a.total} activities done  ·  ${d.delivered}/${d.total} deliverables in`),
    sectionHeader('Owns'),
  ];
  if (!(r.owned || []).length) body.push({ type: 'TextBlock', text: 'No initiatives or works owned directly.', isSubtle: true, wrap: true, spacing: 'Small' });
  (r.owned || []).forEach((w: any) => {
    body.push({
      type: 'ColumnSet', spacing: 'Small', separator: true, columns: [
        { type: 'Column', width: 'stretch', items: [
          { type: 'TextBlock', text: w.title, weight: 'Bolder', wrap: true, spacing: 'None' },
          rich([runT(ragWord(w.rag), { color: ragColor(w.rag), weight: 'Bolder' }), runT(`   ${w.level} · Exec ${w.execution}%`, { isSubtle: true })], 'None'),
        ] },
        ...(w.resultPct != null ? [{ type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [
          { type: 'TextBlock', text: `${w.resultPct}%`, color: ragColor(w.rag), weight: 'Bolder', horizontalAlignment: 'Right', spacing: 'None' },
          { type: 'TextBlock', text: 'result', isSubtle: true, size: 'Small', horizontalAlignment: 'Right', spacing: 'None' },
        ] }] : []),
      ],
    });
  });
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

// ---------- pending approvals (interactive) ----------
export function approvalsCard(pending: any[]) {
  if (!pending.length) return card([header('Approvals'), { type: 'TextBlock', text: '✅  No pending approvals in your scope.', color: 'Good', wrap: true, spacing: 'Small' }]);
  const body: any[] = [header('Pending approvals', `${plural(pending.length, 'request')} awaiting your decision`)];
  for (const c of pending) {
    body.push({
      type: 'Container', separator: true, spacing: 'Medium', style: 'emphasis', items: [
        { type: 'TextBlock', text: String(c.kindLabel || '').toUpperCase(), weight: 'Bolder', color: 'Accent', size: 'Small', spacing: 'None' },
        { type: 'TextBlock', text: c.desc, wrap: true, spacing: 'None' },
        { type: 'TextBlock', text: `${c.proposer}  ·  ${c.initiative}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true },
        { type: 'Input.Text', id: `remark_${c.id}`, placeholder: 'Optional remark…', isMultiline: true, spacing: 'Small' },
        { type: 'Input.Toggle', id: `spinoff_${c.id}`, title: 'Create a follow-up from my remark', value: 'false' },
        { type: 'ActionSet', actions: [
          { type: 'Action.Submit', title: '✓ Approve', style: 'positive', data: { action: 'decide_approval', id: c.id, approve: true } },
          { type: 'Action.Submit', title: '✕ Reject', style: 'destructive', data: { action: 'decide_approval', id: c.id, approve: false } },
        ] },
      ],
    });
  }
  return card(body);
}

const KIND_ICON: Record<string, string> = { document: '📄', spreadsheet: '📊', email: '✉️', slides: '📽️', other: '📎' };

// ---------- a work's deliverables checklist ----------
export function deliverablesCard(d: any) {
  if (!d) return textCard('Deliverables', 'No matching work.');
  const body: any[] = [
    header(d.workTitle, 'Deliverables checklist'),
    kpiStrip([
      { label: 'Delivered', value: `${d.done}/${d.total}`, color: d.total && d.done === d.total ? 'Good' : 'Default' },
      { label: 'Avg score', value: d.avgScore != null ? `${d.avgScore}` : '—' },
    ]),
    ...(d.completedAt ? [{ type: 'TextBlock', text: '✔  Work marked complete', color: 'Good', size: 'Small', spacing: 'Small', wrap: true }] : []),
  ];
  if (!d.items.length) { body.push({ type: 'TextBlock', text: 'No deliverables defined on this work yet.', isSubtle: true, wrap: true, spacing: 'Medium' }); return card(body, [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]); }
  body.push(sectionHeader('Items'));
  d.items.forEach((it: any) => body.push({
    type: 'ColumnSet', spacing: 'Small', separator: true, columns: [
      { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [{ type: 'TextBlock', text: it.done ? '✅' : '⬜' }] },
      { type: 'Column', width: 'stretch', items: [
        { type: 'TextBlock', text: `${KIND_ICON[it.kind] || '📎'}  ${it.label}`, wrap: true, spacing: 'None' },
        ...(it.file || it.verdict ? [{ type: 'TextBlock', text: `${it.file || ''}${it.file && it.verdict ? ' · ' : ''}${it.verdict || ''}`, isSubtle: true, size: 'Small', spacing: 'None', wrap: true }] : []),
      ] },
      { type: 'Column', width: 'auto', verticalContentAlignment: 'Center', items: [{ type: 'TextBlock', text: it.score != null ? `${it.score}` : '', color: 'Good', weight: 'Bolder', horizontalAlignment: 'Right' }] },
    ],
  }));
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open portal', url: PORTAL }]);
}

// ---------- a deliverable just logged from a chat file ----------
export function deliverableLoggedCard(work: any, item: any, scored: boolean) {
  const body: any[] = [
    header('Deliverable logged ✅', `${KIND_ICON[item.kind] || '📎'}  ${item.label}`),
    { type: 'FactSet', spacing: 'Small', facts: [
      { title: 'Work', value: work.title },
      { title: 'File', value: item.file ? item.file.name : '—' },
      ...(scored && typeof item.score === 'number' ? [{ title: 'AI score', value: `${item.score}/100  ·  ${item.verdict || ''}` }] : []),
    ] },
    ...(scored && item.feedback ? [note(item.feedback)] : []),
  ];
  return card(body, [{ type: 'Action.OpenUrl', title: 'Open in portal', url: PORTAL }]);
}

// ---------- plain text fallback ----------
export function textCard(heading: string, msg: string) {
  return card([header(heading), { type: 'TextBlock', text: msg, wrap: true, spacing: 'Small' }]);
}
