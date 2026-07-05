// Express-mountable Teams bot host. The SAME CloudAdapter + CadenceBot the
// standalone restify server (index.ts) uses, factored out so the merged
// single-App-Service process can serve the bot's /api/messages endpoint
// alongside the Cadence API and the portal static files.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
} from 'botbuilder';
import { CadenceBot } from './bot';

// Build the adapter + bot from env (MICROSOFT_APP_ID / _PASSWORD / _TYPE /
// _TENANT_ID). Shared by the standalone host and the in-process mount.
export function buildBot() {
  const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MICROSOFT_APP_ID,
    MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD,
    MicrosoftAppType: process.env.MICROSOFT_APP_TYPE || 'MultiTenant',
    MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
  });
  const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({}, credentialsFactory);
  const adapter = new CloudAdapter(botFrameworkAuthentication);
  adapter.onTurnError = async (context: TurnContext, error: Error) => {
    console.error('[onTurnError]', error);
    try { await context.sendActivity('The Cadence agent hit an error. Please try again.'); } catch { /* channel may be gone */ }
  };
  return { adapter, bot: new CadenceBot() };
}

// Mount POST /api/messages on an existing Express app (used by the merged server).
// The handler MUST catch — adapter.process() rejects on malformed activities
// (empty/garbage bodies from health probes or internet scanners), and an
// unhandled rejection would crash the whole merged process (portal + API too).
export function registerBot(app: any) {
  const { adapter, bot } = buildBot();
  app.post('/api/messages', async (req: any, res: any) => {
    try {
      await adapter.process(req, res, (context: any) => bot.run(context));
    } catch (err: any) {
      console.error('[/api/messages] rejected:', err?.message || err);
      if (!res.headersSent) res.status(400).json({ error: 'invalid bot activity' });
    }
  });
  return { adapter, bot };
}
