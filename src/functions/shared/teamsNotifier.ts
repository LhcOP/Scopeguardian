import { InvocationContext } from "@azure/functions";
import { BotFrameworkAdapter, TurnContext, Activity, CardFactory } from "botbuilder";

let adapter: BotFrameworkAdapter | null = null;

function getAdapter(): BotFrameworkAdapter {
  if (!adapter) {
    const appId = process.env.BOT_APP_ID;
    const appPassword = process.env.BOT_APP_PASSWORD;
    if (!appId || !appPassword) {
      throw new Error("BOT_APP_ID or BOT_APP_PASSWORD is not set");
    }
    adapter = new BotFrameworkAdapter({ appId, appPassword });
  }
  return adapter;
}

export async function sendTeamsAlert(
  teamId: string,
  channelId: string,
  card: object,
  context: InvocationContext
): Promise<void> {
  const serviceUrl = process.env.BOT_SERVICE_URL ?? "https://smba.trafficmanager.net/teams/";
  const botAdapter = getAdapter();

  const conversationRef = {
    serviceUrl,
    channelId: "msteams",
    conversation: {
      id: channelId,
      name: "",
      isGroup: true,
      conversationType: "channel",
      tenantId: process.env.AZURE_TENANT_ID ?? "",
    },
    bot: { id: process.env.BOT_APP_ID!, name: "ScopeGuardian" },
  };

  try {
    await botAdapter.continueConversation(conversationRef, async (turnContext: TurnContext) => {
      const adaptiveCardAttachment = CardFactory.adaptiveCard(card);
      const activity: Partial<Activity> = {
        type: "message",
        attachments: [adaptiveCardAttachment],
      };
      await turnContext.sendActivity(activity);
    });
  } catch (err) {
    context.error("Failed to send Teams alert:", err);
    // Non-fatal — violation is still persisted in Cosmos DB
  }
}
