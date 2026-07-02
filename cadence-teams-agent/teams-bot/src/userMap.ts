// Map the signed-in Teams user to a Cadence user. Production: match the AAD
// email/UPN to a Cadence account. Demo fallback: treat the caller as the MD so
// the conversation works even before identities are linked.
import { TurnContext, TeamsInfo } from 'botbuilder';
import { cadence } from './cadenceClient';

const DEFAULT_USER_ID = process.env.DEFAULT_CADENCE_USER_ID || 'u_vik';

export async function getCadenceUserId(context: TurnContext): Promise<string> {
  try {
    const member: any = await TeamsInfo.getMember(context, context.activity.from!.id);
    const email = member?.email || member?.userPrincipalName;
    if (email) {
      const match = await cadence.resolve('user', email);
      if (match) return match.id;
    }
  } catch {
    // getMember can fail in personal scope / local emulator — fall through.
  }
  return DEFAULT_USER_ID;
}
