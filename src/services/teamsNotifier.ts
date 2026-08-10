import { InvocationContext } from "@azure/functions";
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
  TurnContext,
  Activity,
  CardFactory,
  ConversationReference,
} from "botbuilder";
import { getConversationRef } from "./cosmosDbService";

let adapter: CloudAdapter | null = null;

export function getCloudAdapter(): CloudAdapter {
  if (!adapter) {
    const appId = process.env.BOT_APP_ID;
    const appPassword = process.env.BOT_APP_PASSWORD;
    if (!appId || !appPassword) {
      throw new Error("BOT_APP_ID or BOT_APP_PASSWORD is not set");
    }
    // Azure has deprecated multi-tenant bot creation — default to SingleTenant
    const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
      MicrosoftAppId: appId,
      MicrosoftAppPassword: appPassword,
      MicrosoftAppType: process.env.BOT_APP_TYPE ?? "SingleTenant",
      MicrosoftAppTenantId: process.env.AZURE_TENANT_ID,
    });
    const botAuth = new ConfigurationBotFrameworkAuthentication({}, credentialsFactory);
    adapter = new CloudAdapter(botAuth);
  }
  return adapter;
}

/**
 * Sends an Adaptive Card to a Teams channel proactively.
 * Prefers a conversation reference captured from real bot traffic (correct
 * regional serviceUrl); falls back to a constructed reference using
 * BOT_SERVICE_URL. Requires the bot to be installed in the target team.
 */
export async function sendTeamsAlert(
  channelId: string,
  card: object,
  context: InvocationContext
): Promise<void> {
  if (!channelId) {
    context.warn("TEAMS_CHANNEL_ID not set — skipping Teams alert");
    return;
  }

  const stored = await getConversationRef(channelId).catch(() => null);
  const reference = (stored ?? {
    channelId: "msteams",
    serviceUrl: process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/teams/",
    conversation: {
      id: channelId,
      name: "",
      isGroup: true,
      conversationType: "channel",
      tenantId: process.env.AZURE_TENANT_ID ?? "",
    },
    bot: { id: process.env.BOT_APP_ID ?? "", name: "ScopeGuardian" },
  }) as Partial<ConversationReference>;

  try {
    await getCloudAdapter().continueConversationAsync(
      process.env.BOT_APP_ID ?? "",
      reference,
      async (turnContext: TurnContext) => {
        const activity: Partial<Activity> = {
          type: "message",
          attachments: [CardFactory.adaptiveCard(card)],
        };
        await turnContext.sendActivity(activity);
      }
    );
  } catch (err) {
    context.error("Failed to send Teams alert:", err);
    // Non-fatal — violation is still persisted in Cosmos DB
  }
}
