// Cadence Teams bot. Handles the MD's planning conversation (LLM + tools),
// renders Adaptive Cards, and processes card button actions (Approve/Reject).
import { TeamsActivityHandler, TurnContext, MessageFactory, CardFactory } from 'botbuilder';
import type { ChatCompletionMessageParam } from 'openai/resources';
import { runConversation } from './foundry';
import { tools, makeDispatch, type ChatFile } from './tools';
import { getCadenceUserId } from './userMap';
import { cadence } from './cadenceClient';
import * as cards from './cards';

const HISTORY_TURNS = 8;

// A user can drop a file in chat to log it as a deliverable. Teams delivers it as
// a "file download info" attachment with a pre-authenticated downloadUrl (no auth
// header needed). Fetch the first usable one and hand it to the tools as base64.
const FILE_DL = 'application/vnd.microsoft.teams.file.download.info';
async function extractChatFile(context: TurnContext): Promise<ChatFile | null> {
  const atts: any[] = context.activity.attachments || [];
  for (const a of atts) {
    let url: string | undefined;
    if (a.contentType === FILE_DL && a.content?.downloadUrl) url = a.content.downloadUrl;
    else if (a.contentUrl && a.contentType && !a.contentType.startsWith('text/html') && a.contentType !== 'application/vnd.microsoft.card.adaptive') url = a.contentUrl;
    if (!url) continue;
    try {
      const r = await fetch(url);
      if (!r.ok) continue;
      const buf = Buffer.from(await r.arrayBuffer());
      if (buf.length && buf.length < 25 * 1024 * 1024) return { name: a.name || 'file', base64: buf.toString('base64') };
    } catch { /* try the next attachment */ }
  }
  return null;
}
const histories = new Map<string, ChatCompletionMessageParam[]>();
// Trim the rolling history to the last N conversation turns — but ONLY on user-
// message boundaries. A turn is a group: user → assistant(tool_calls) → tool(s)
// → assistant(reply). A blind tail-slice can cut through that group, orphaning a
// `tool` message at the head (no preceding `tool_calls`), which makes the very
// next request fail with a 400 and never recover. Starting the retained history
// at a user message keeps every tool_calls/tool pair intact.
function trimHistory(messages: ChatCompletionMessageParam[]): ChatCompletionMessageParam[] {
  const msgs = messages.filter((m) => m.role !== 'system');
  const userIdx: number[] = [];
  msgs.forEach((m, i) => { if (m.role === 'user') userIdx.push(i); });
  const start = userIdx.length > HISTORY_TURNS ? userIdx[userIdx.length - HISTORY_TURNS] : 0;
  return msgs.slice(start);
}

const SYSTEM = (scopeNote: string) => `You are the Cadence agent inside Microsoft Teams — an execution assistant for enterprise leaders. You help the signed-in leader PLAN and OPERATE work that lives in the Cadence portal.

Cadence object model: Objective → Initiative → Work → Activity. An Activity is the schedulable unit (owner + date + hours + type). Results are tracked against a target; each initiative shows planning%, execution%, a RAG, and a "sufficiency" verdict (whether the plan is enough to close the gap to target).

How to behave:
- Everything the user asks about state, use a tool — never invent numbers. ${scopeNote}
- When the user wants to plan something new: YOU decompose it into 3-6 works (phases), each with 1-4 concrete activities (title, estimateHrs, type = self/meeting/call/site). Briefly show the proposed plan and the target team, ask for a yes, THEN call plan_initiative. The service assigns activities across the team by load and dates them to the deadline.
- Before assigning people, you may check team_capacity so you don't overload someone.
- The cards are INTERACTIVE and drill down. For anything at the OBJECTIVE level ("which objectives are overdue", "objective report", "how are the objectives") use get_objectives; get_objective opens one objective; get_work opens one work. Each card's rows are tappable to drill in, so you rarely need to fetch children yourself — return the right top card and let the user drill.
- Each Work carries a DELIVERABLES checklist. get_work shows it with tick-boxes; the user can mark items done, auto-assign, suggest deliverables, or mark complete right on the card. When the user ATTACHES a file, log it against a deliverable with log_deliverable (it marks the item delivered and AI-scores the file). When a work's outputs are in, complete_work marks it and its activities done.
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
      const file = await extractChatFile(context);
      if (!text && !file) { await next(); return; }

      const userId = await getCadenceUserId(context);
      const convoId = context.activity.conversation.id;
      const history = histories.get(convoId) || [];
      await context.sendActivity({ type: 'typing' });

      try {
        const scopeNote = 'Results are already scoped to the caller by the API — just present what comes back.';
        const dispatch = makeDispatch({ userId, file });
        // Tell the model a file is present so it can offer to log it as a deliverable.
        const userText = file
          ? `${text || 'I attached a file.'}\n\n[The user attached a file named "${file.name}". If they want it recorded against a work's deliverable, call log_deliverable (list_deliverables first if the exact item is unclear).]`
          : text;
        const { text: reply, cards: attachCards, messages } = await runConversation(SYSTEM(scopeNote), history, userText, dispatch, tools as any);

        // keep a trimmed rolling history (drop system; keep last N whole turns)
        histories.set(convoId, trimHistory(messages));

        // One message per turn: when a tool produced a card, that card IS the
        // answer — send it alone (no duplicate text bubble restating the same
        // numbers). Otherwise (clarifying questions, plain answers) send the text.
        if (attachCards.length) {
          await context.sendActivity(MessageFactory.list(attachCards.map((c) => CardFactory.adaptiveCard(c))));
        } else {
          await context.sendActivity(MessageFactory.text(reply && reply.trim() ? reply : ' '));
        }
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
            "Hi — I'm your Cadence agent. Ask me things like:\n\n• *What's at risk across the portfolio?*\n• *Status of the laptop refresh?*\n• *Plan an MFA rollout for field staff, assign it to IT Ops, deadline in two weeks.*\n• *Show the deliverables for process mapping.*\n• *(attach a file)* *Log this against the vendor quotes deliverable on demand & selection.*\n• *Who has capacity this week?*",
          ))));
        }
      }
      await next();
    });
  }

  // Refresh the SAME card message in place (updateActivity) so drilling and
  // ticking deliverables never spawns a new message. Falls back to a new
  // message if Teams didn't give us the id to edit.
  private async refresh(context: TurnContext, cardJson: any) {
    const activity: any = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
    const id = context.activity.replyToId;
    if (id) { activity.id = id; await context.updateActivity(activity); }
    else await context.sendActivity(activity);
  }

  private async buildCard(kind: string, id: string) {
    if (kind === 'objective') return cards.objectiveCard(await cadence.objective(id));
    if (kind === 'initiative') return cards.initiativeCard(await cadence.initiative(id));
    return cards.workCard(await cadence.workDetail(id));
  }

  private async handleCardAction(context: TurnContext, value: any) {
    try {
      switch (value.action) {
        case 'decide_approval': {
          const userId = await getCadenceUserId(context);
          const remark = value[`remark_${value.id}`] || value.remark || '';
          const spinoff = String(value[`spinoff_${value.id}`]) === 'true';
          const res = await cadence.decideApproval(value.id, { approve: !!value.approve, remark, spinoff, approverId: userId });
          const msg = res.cr.status === 'approved' ? `Approved.${res.spun ? ` Follow-up initiative created: “${res.spun.title}”.` : ''}` : 'Rejected.';
          await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard(cards.textCard('Decision recorded', msg))));
          return;
        }
        // drill up/down — re-render this card as the target level
        case 'drill':
          return this.refresh(context, await this.buildCard(value.kind, value.id));
        // tick / untick a deliverable in place
        case 'deliv_toggle': {
          const r = await cadence.toggleDeliverable(value.workId, value.did);
          return this.refresh(context, cards.workCard(r.work));
        }
        case 'auto_assign': {
          const r = await cadence.autoAssignWork(value.workId);
          return this.refresh(context, cards.workCard(r.work));
        }
        case 'suggest_deliv': {
          const r = await cadence.suggestDeliverables(value.workId);
          return this.refresh(context, cards.workCard(r.work));
        }
        case 'work_complete': {
          const w = await cadence.workDetail(value.workId);
          return this.refresh(context, cards.confirmCompleteCard(w));
        }
        case 'work_complete_go': {
          await cadence.completeWork(value.workId);
          return this.refresh(context, cards.workCard(await cadence.workDetail(value.workId)));
        }
      }
    } catch (e: any) {
      await context.sendActivity(`Couldn't do that: ${e.message || e}`);
    }
  }
}
