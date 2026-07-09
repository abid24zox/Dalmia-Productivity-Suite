// The bridge between the LLM and Cadence. `tools` are the function schemas the
// Foundry model sees; `makeDispatch` executes a chosen tool against the Cadence
// API and returns { data } (fed back to the model) and an optional { card }
// (shown to the user).
import { cadence, needTeam, needUser, needActivity, needWork, needObjective } from './cadenceClient';
import * as cards from './cards';

export type ChatFile = { name: string; base64: string };

export const tools = [
  { type: 'function', function: { name: 'get_portfolio', description: "Show the caller's portfolio/scorecard, scoped to their level (the CEO sees the enterprise; a VP sees their function; a member sees their own work). Use for 'how are we doing', 'what's the portfolio', 'what's at risk overall'.", parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'get_initiative_status', description: 'Get the detailed status of one initiative: result vs target, gap, sufficiency, and its work breakdown (each work tappable to drill in). Use when the user names or refers to a specific initiative.', parameters: { type: 'object', properties: { initiative: { type: 'string', description: 'Initiative title or a distinctive part of it, e.g. "laptops" or "logistics cost".' } }, required: ['initiative'] } } },
  { type: 'function', function: { name: 'get_objectives', description: "The objective-level report — every objective the caller can see, with its RAG, avg result, initiative count and overdue count, each tappable to drill into its initiatives. Use for 'which objectives are overdue', 'objective report', 'how are the objectives doing', 'show me the objectives'.", parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'get_objective', description: 'Open ONE objective: its rollup plus the initiatives under it, each tappable. Use when the user names an objective (e.g. "the dealer experience objective").', parameters: { type: 'object', properties: { objective: { type: 'string', description: 'Objective title or a distinctive part.' } }, required: ['objective'] } } },
  { type: 'function', function: { name: 'get_work', description: 'Open ONE work package: its activities and its deliverables checklist, with interactive buttons (tick a deliverable done, auto-assign, suggest deliverables, mark complete). Use when the user names a work (e.g. "open process mapping").', parameters: { type: 'object', properties: { work: { type: 'string', description: 'Work title or a distinctive part.' } }, required: ['work'] } } },
  { type: 'function', function: { name: 'plan_initiative', description: 'Create a new initiative from a decomposition YOU produce, and assign it to a team. Break the goal into 3-6 works (phases of execution), each with 1-4 concrete activities (each with an hour estimate and a type of self/meeting/call/site). The service distributes activities across the team balanced by load and spreads dates to the deadline. Always confirm the plan with the user before calling this.', parameters: { type: 'object', properties: { title: { type: 'string' }, type: { type: 'string', enum: ['procurement', 'cost', 'onboarding', 'compliance', 'general'] }, objective: { type: 'string' }, deadline: { type: 'string', description: 'ISO date YYYY-MM-DD' }, team: { type: 'string', description: 'Team name to assign to, e.g. "IT Ops".' }, works: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, activities: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, estimateHrs: { type: 'number' }, type: { type: 'string', enum: ['self', 'meeting', 'call', 'site'] } }, required: ['title'] } } }, required: ['title', 'activities'] } } }, required: ['title', 'type', 'team', 'works'] } } },
  { type: 'function', function: { name: 'schedule_activity', description: 'Assign and/or date an existing activity. Use to place an unscheduled activity onto a person and a day.', parameters: { type: 'object', properties: { activity: { type: 'string', description: 'Activity title or distinctive part.' }, assignee: { type: 'string', description: 'Person name (optional).' }, date: { type: 'string', description: 'ISO date YYYY-MM-DD (optional).' } }, required: ['activity'] } } },
  { type: 'function', function: { name: 'reassign_activity', description: 'Move an activity to a different owner.', parameters: { type: 'object', properties: { activity: { type: 'string' }, assignee: { type: 'string' } }, required: ['activity', 'assignee'] } } },
  { type: 'function', function: { name: 'list_attention', description: "What needs attention in the caller's scope: overdue, blocked, and stuck items.", parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'list_approvals', description: 'List pending plan-change approvals proposed by the team, with Approve/Reject buttons.', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'decide_approval', description: 'Approve or reject a pending change by its id. Optionally add a remark and spin off a follow-up initiative from it.', parameters: { type: 'object', properties: { change_id: { type: 'string' }, approve: { type: 'boolean' }, remark: { type: 'string' }, spinoff: { type: 'boolean' } }, required: ['change_id', 'approve'] } } },
  { type: 'function', function: { name: 'team_capacity', description: 'Show open-work load per person, to see who is free or overloaded before assigning.', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'get_member_status', description: "One person's performance scorecard: the initiatives/works they own with RAG status and result-vs-target, their activity completion %, overdue and blocked counts, and deliverable quality. Scoped to the caller (CEO sees anyone; a VP sees their own function or reports; a member sees only themselves). Use for 'how is Rajeev doing', 'show Neha's performance', 'how's my supply-chain VP tracking'.", parameters: { type: 'object', properties: { person: { type: 'string', description: "Person's name or a distinctive part, e.g. 'Rajeev' or 'Neha'." } }, required: ['person'] } } },
  { type: 'function', function: { name: 'list_teams', description: 'List the teams and their members (for choosing who to assign to, or to confirm a merge).', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'list_deliverables', description: "Show a work package's deliverables checklist — the concrete outputs it must produce, which are delivered/scored, and whether the work is complete. Use before logging a file so you know the exact item labels.", parameters: { type: 'object', properties: { work: { type: 'string', description: 'Work package title or a distinctive part, e.g. "process mapping" or "demand".' } }, required: ['work'] } } },
  { type: 'function', function: { name: 'log_deliverable', description: "Log the file the user ATTACHED to their current message against a deliverable on a work, marking it delivered and (by default) AI-scoring it. Only call this when the user actually attached a file. If unsure of the exact deliverable label, call list_deliverables first.", parameters: { type: 'object', properties: { work: { type: 'string', description: 'Work package title or distinctive part.' }, deliverable: { type: 'string', description: 'The checklist item label to log against (fuzzy-matched), e.g. "vendor quotes".' }, score: { type: 'boolean', description: 'Whether to AI-score the file (default true).' }, create: { type: 'boolean', description: 'Set true to add the deliverable as a new checklist item when it does not already exist.' } }, required: ['work', 'deliverable'] } } },
  { type: 'function', function: { name: 'complete_work', description: 'Mark a work package complete — sets all its open activities to executed and stamps it done. Use when the user says a work is finished, ideally once its deliverables are in.', parameters: { type: 'object', properties: { work: { type: 'string', description: 'Work package title or distinctive part.' } }, required: ['work'] } } },
];

export function makeDispatch(ctx: { userId: string; file?: ChatFile | null }) {
  return async function dispatch(name: string, args: any): Promise<{ data: any; card?: any }> {
    try {
      switch (name) {
        case 'get_portfolio': {
          const p = await cadence.portfolio(ctx.userId);
          return { data: p, card: cards.portfolioCard(p) };
        }
        case 'get_initiative_status': {
          const i = await cadence.initiative(args.initiative);
          return { data: i, card: cards.initiativeCard(i) };
        }
        case 'get_objectives': {
          const r = await cadence.objectives(ctx.userId);
          return { data: r, card: cards.objectivesReportCard(r) };
        }
        case 'get_objective': {
          const o = await needObjective(args.objective);
          const d = await cadence.objective(o.id);
          return { data: d, card: cards.objectiveCard(d) };
        }
        case 'get_work': {
          const w = await needWork(args.work);
          const d = await cadence.workDetail(w.id);
          return { data: d, card: cards.workCard(d) };
        }
        case 'plan_initiative': {
          const team = await needTeam(args.team);
          const res = await cadence.createInitiative({ ownerId: ctx.userId, title: args.title, type: args.type, objective: args.objective, deadline: args.deadline, teamId: team.id, works: args.works });
          return { data: res.initiative, card: cards.createdCard(res.initiative) };
        }
        case 'schedule_activity': {
          const act = await needActivity(args.activity);
          const body: any = {};
          if (args.assignee) body.assigneeId = (await needUser(args.assignee)).id;
          if (args.date) body.date = args.date;
          const res = await cadence.scheduleActivity(act.id, body);
          return { data: { ok: true, activity: res.activity }, card: cards.textCard('Activity scheduled', `“${act.name}” → ${args.assignee || 'same owner'}${args.date ? ' on ' + args.date : ''}.`) };
        }
        case 'reassign_activity': {
          const act = await needActivity(args.activity);
          const to = await needUser(args.assignee);
          await cadence.reassignActivity(act.id, to.id);
          return { data: { ok: true }, card: cards.textCard('Activity reassigned', `“${act.name}” → ${to.name}.`) };
        }
        case 'list_attention': {
          const a = await cadence.attention(ctx.userId);
          return { data: a, card: cards.attentionCard(a.items) };
        }
        case 'list_approvals': {
          const a = await cadence.approvals(ctx.userId);
          return { data: a, card: cards.approvalsCard(a.pending) };
        }
        case 'decide_approval': {
          const res = await cadence.decideApproval(args.change_id, { approve: args.approve, remark: args.remark, spinoff: args.spinoff, approverId: ctx.userId });
          const msg = res.cr.status === 'approved' ? `Approved.${res.spun ? ' Follow-up initiative created: “' + res.spun.title + '”.' : ''}` : 'Rejected.';
          return { data: res, card: cards.textCard('Decision recorded', msg) };
        }
        case 'team_capacity': {
          const c = await cadence.capacity(ctx.userId);
          return { data: c, card: cards.capacityCard(c.capacity) };
        }
        case 'get_member_status': {
          const r = await cadence.memberStatus(ctx.userId, args.person);
          return { data: r, card: cards.memberStatusCard(r) };
        }
        case 'list_teams': {
          const t = await cadence.teams();
          return { data: t };
        }
        case 'list_deliverables': {
          const d = await cadence.deliverables(args.work);
          return { data: d, card: cards.deliverablesCard(d) };
        }
        case 'log_deliverable': {
          if (!ctx.file) return { data: { error: 'No file is attached to this message. Ask the user to attach the file to the same message and try again.' } };
          const w = await needWork(args.work);
          const res = await cadence.attachDeliverable(w.id, { label: args.deliverable, fileBase64: ctx.file.base64, fileName: ctx.file.name, score: args.score !== false, create: !!args.create });
          return { data: { ok: true, item: res.item, scored: res.scored }, card: cards.deliverableLoggedCard(res.work, res.item, res.scored) };
        }
        case 'complete_work': {
          const w = await needWork(args.work);
          const res = await cadence.completeWork(w.id);
          return { data: res, card: cards.textCard('Work marked complete', `“${res.work.title}” — ${res.activitiesCompleted} of ${res.totalActivities} ${res.totalActivities === 1 ? 'activity' : 'activities'} set to executed.`) };
        }
        default:
          return { data: { error: `unknown tool ${name}` } };
      }
    } catch (e: any) {
      // Tool errors are returned to the model so it can ask a clarifying question.
      return { data: { error: e.message || String(e) } };
    }
  };
}
