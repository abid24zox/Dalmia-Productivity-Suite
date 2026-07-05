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
export function registerBot(app: any) {
  const { adapter, bot } = buildBot();
  app.post('/api/messages', (req: any, res: any) => adapter.process(req, res, (context: any) => bot.run(context)));
  return { adapter, bot };
}
