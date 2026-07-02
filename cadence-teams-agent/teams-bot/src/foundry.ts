// Azure AI Foundry model client (Azure OpenAI-compatible chat completions with
// tool calling). Credentials come from env — nothing hardcoded. The loop runs
// tool calls until the model returns a final message.
import { AzureOpenAI } from 'openai';
import type { ChatCompletionMessageParam, ChatCompletionTool } from 'openai/resources';

const client = new AzureOpenAI({
  endpoint: process.env.AZURE_OPENAI_ENDPOINT,
  apiKey: process.env.AZURE_OPENAI_API_KEY,
  apiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-08-01-preview',
  deployment: process.env.AZURE_OPENAI_DEPLOYMENT,
});
const MODEL = process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o';

export type DispatchFn = (name: string, args: any) => Promise<{ data: any; card?: any }>;

export async function runConversation(
  system: string,
  history: ChatCompletionMessageParam[],
  userText: string,
  dispatch: DispatchFn,
  tools: ChatCompletionTool[],
): Promise<{ text: string; cards: any[]; messages: ChatCompletionMessageParam[] }> {
  const messages: ChatCompletionMessageParam[] = [
    { role: 'system', content: system },
    ...history,
    { role: 'user', content: userText },
  ];
  const cards: any[] = [];

  for (let step = 0; step < 6; step++) {
    const resp = await client.chat.completions.create({
      model: MODEL,
      messages,
      tools,
      tool_choice: 'auto',
      // gpt-5.x minis require max_completion_tokens (not max_tokens) and only
      // support the default temperature.
      max_completion_tokens: 1600,
    });
    const msg = resp.choices[0].message;
    messages.push(msg as ChatCompletionMessageParam);

    if (msg.tool_calls && msg.tool_calls.length) {
      for (const tc of msg.tool_calls) {
        let args: any = {};
        try { args = JSON.parse(tc.function.arguments || '{}'); } catch { /* leave empty */ }
        const { data, card } = await dispatch(tc.function.name, args);
        if (card) cards.push(card);
        messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(data).slice(0, 6000) });
      }
      continue; // let the model read the tool results and respond
    }
    return { text: (msg.content as string) || '', cards, messages };
  }
  return { text: "I couldn't complete that in a few steps — try narrowing the request.", cards, messages };
}
