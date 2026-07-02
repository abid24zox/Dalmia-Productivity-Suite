// Cadence Teams bot. Handles the MD's planning conversation (LLM + tools),
// renders Adaptive Cards, and processes card button actions (Approve/Reject).
import { TeamsActivityHandler, TurnContext, MessageFactory, CardFactory } from 'botbuilder';
import type { ChatCompletionMessageParam } from 'openai/resources';
import { runConversation } from './foundry';
import { tools, makeDispatch } from './tools';
import { getCadenceUserId } from './userMap';
import { cadence } from './cadenceClient';
import * as cards from './cards';

const HISTORY_TURNS = 12;
const histories = new Map<string, ChatCompletionMessageParam[]>();

const SYSTEM = (scopeNote: string) => `You are the Cadence agent inside Microsoft Teams — an execution assistant for enterprise leaders. You help the signed-in leader PLAN and OPERATE work that lives in the Cadence portal.

Cadence object model: Objective → Initiative → Work → Sub-work → Activity. An Activity is the schedulable unit (owner + date + hours + type). Results are tracked against a target; each initiative shows planning%, execution%, a RAG, and a "sufficiency" verdict (whether the plan is enough to close the gap to target).

How to behave:
- Everything the user asks about state, use a tool — never invent numbers. ${scopeNote}
- When the user wants to plan something new: YOU decompose it into 3-6 sub-works, each with 1-4 concrete activities (title, estimateHrs, type = self/meeting/call/site). Briefly show the proposed plan and the target team, ask for a yes, THEN call plan_initiative. The service assigns activities across the team by load and dates them to the deadline.
- Before assigning people, you may check team_capacity so you don't overload someone.
- Keep replies short and executive. A card usually accompanies your answer, so don't repeat every number in prose — give the headline and the recommendation.
- If a tool returns an error asking which team/person/initiative, ask the user that one question.
- Today's date is ${new Date().toISOString().slice(0, 10)}.`;

export class CadenceBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      // Adaptive Card button (Action.Submit) arrives as a message with .value
      const value: any = context.activity.value;
      if (value && value.action) {
        await this.handleCardAction(context, value);
        await next();
        return;
      }

      const text = (context.activity.text || '').trim();
      if (!text) { await next(); return; }

      const userId = await getCadenceUserId(context);
      const convoId = context.activity.conversation.id;
      const history = histories.get(convoId) || [];
      await context.sendActivity({ type: 'typing' });

      try {
        const scopeNote = 'Results are already scoped to the caller by the API — just present what comes back.';
        const dispatch = makeDispatch({ userId });
        const { text: reply, cards: attachCards, messages } = await runConversation(SYSTEM(scopeNote), history, text, dispatch, tools as any);

        // keep a trimmed rolling history (drop system; keep last N)
        histories.set(convoId, messages.filter((m) => m.role !== 'system').slice(-HISTORY_TURNS));

        const activity = MessageFactory.text(reply && reply.trim() ? reply : ' ');
        if (attachCards.length) activity.attachments = attachCards.map((c) => CardFactory.adaptiveCard(c));
        await context.sendActivity(activity);
      } catch (e: any) {
        await context.sendActivity(`Something went wrong reaching Cadence: ${e.message || e}. Check the service is up and credentials are set.`);
      }
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      for (const m of context.activity.membersAdded || []) {
        if (m.id !== context.activity.recipient.id) {
          await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(cards.textCard(
            'Cadence in Teams',
            "Hi — I'm your Cadence agent. Ask me things like:\n\n• *What's at risk across the portfolio?*\n• *Status of the laptop refresh?*\n• *Plan an MFA rollout for field staff, assign it to IT Ops, deadline in two weeks.*\n• *Who has capacity this week?*\n• *Show pending approvals.*",
          ))));
        }
      }
      await next();
    });
  }

  private async handleCardAction(context: TurnContext, value: any) {
    if (value.action === 'decide_approval') {
      const userId = await getCadenceUserId(context);
      const remark = value[`remark_${value.id}`] || value.remark || '';
      const spinoff = String(value[`spinoff_${value.id}`]) === 'true';
      try {
        const res = await cadence.decideApproval(value.id, { approve: !!value.approve, remark, spinoff, approverId: userId });
        const msg = res.cr.status === 'approved'
          ? `Approved.${res.spun ? ` Follow-up initiative created: “${res.spun.title}”.` : ''}`
          : 'Rejected.';
        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(cards.textCard('Decision recorded', msg))));
      } catch (e: any) {
        await context.sendActivity(`Couldn't record that decision: ${e.message || e}`);
      }
    }
  }
}
