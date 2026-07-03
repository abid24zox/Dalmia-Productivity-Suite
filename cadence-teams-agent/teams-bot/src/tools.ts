// The bridge between the LLM and Cadence. `tools` are the function schemas the
// Foundry model sees; `makeDispatch` executes a chosen tool against the Cadence
// API and returns { data } (fed back to the model) and an optional { card }
// (shown to the user).
import { cadence, needTeam, needUser, needActivity } from './cadenceClient';
import * as cards from './cards';

export const tools = [
  { type: 'function', function: { name: 'get_portfolio', description: "Show the caller's portfolio/scorecard, scoped to their level (MD sees the enterprise; a VP sees their function; a member sees their own work). Use for 'how are we doing', 'what's the portfolio', 'what's at risk overall'.", parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'get_initiative_status', description: 'Get the detailed status of one initiative: result vs target, gap, sufficiency, and the work/activity breakdown. Use when the user names or refers to a specific initiative.', parameters: { type: 'object', properties: { initiative: { type: 'string', description: 'Initiative title or a distinctive part of it, e.g. "laptops" or "logistics cost".' } }, required: ['initiative'] } } },
  { type: 'function', function: { name: 'plan_initiative', description: 'Create a new initiative from a decomposition YOU produce, and assign it to a team. Break the goal into 3-6 works (phases), each with 1-4 concrete activities (each with an hour estimate and a type of self/meeting/call/site). The service distributes activities across the team balanced by load and spreads dates to the deadline. Always confirm the plan with the user before calling this.', parameters: { type: 'object', properties: { title: { type: 'string' }, type: { type: 'string', enum: ['procurement', 'cost', 'onboarding', 'compliance', 'general'] }, objective: { type: 'string' }, deadline: { type: 'string', description: 'ISO date YYYY-MM-DD' }, team: { type: 'string', description: 'Team name to assign to, e.g. "IT Ops".' }, works: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, activities: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, estimateHrs: { type: 'number' }, type: { type: 'string', enum: ['self', 'meeting', 'call', 'site'] } }, required: ['title'] } } }, required: ['title', 'activities'] } } }, required: ['title', 'type', 'team', 'works'] } } },
  { type: 'function', function: { name: 'schedule_activity', description: 'Assign and/or date an existing activity. Use to place an unscheduled activity onto a person and a day.', parameters: { type: 'object', properties: { activity: { type: 'string', description: 'Activity title or distinctive part.' }, assignee: { type: 'string', description: 'Person name (optional).' }, date: { type: 'string', description: 'ISO date YYYY-MM-DD (optional).' } }, required: ['activity'] } } },
  { type: 'function', function: { name: 'reassign_activity', description: 'Move an activity to a different owner.', parameters: { type: 'object', properties: { activity: { type: 'string' }, assignee: { type: 'string' } }, required: ['activity', 'assignee'] } } },
  { type: 'function', function: { name: 'list_attention', description: "What needs attention in the caller's scope: overdue, blocked, and stuck items.", parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'list_approvals', description: 'List pending plan-change approvals proposed by the team, with Approve/Reject buttons.', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'decide_approval', description: 'Approve or reject a pending change by its id. Optionally add a remark and spin off a follow-up initiative from it.', parameters: { type: 'object', properties: { change_id: { type: 'string' }, approve: { type: 'boolean' }, remark: { type: 'string' }, spinoff: { type: 'boolean' } }, required: ['change_id', 'approve'] } } },
  { type: 'function', function: { name: 'team_capacity', description: 'Show open-work load per person, to see who is free or overloaded before assigning.', parameters: { type: 'object', properties: {} } } },
  { type: 'function', function: { name: 'list_teams', description: 'List the teams and their members (for choosing who to assign to, or to confirm a merge).', parameters: { type: 'object', properties: {} } } },
];

export function makeDispatch(ctx: { userId: string }) {
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
        case 'list_teams': {
          const t = await cadence.teams();
          return { data: t };
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
