import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { BotFrameworkAdapter, TurnContext, WebRequest, WebResponse } from "botbuilder";
import { BotActivityPayload, FeedbackAction } from "../models/FeedbackLog";
import { updateViolationStatus, insertFeedbackLog, listViolations, getScopeSummary } from "../services/cosmosDbService";
import { buildConfirmationCard, buildScopeStatusCard } from "../utils/adaptiveCardBuilder";
import { v4 as uuidv4 } from "uuid";

let adapter: BotFrameworkAdapter | null = null;

function getAdapter(): BotFrameworkAdapter {
  if (!adapter) {
    const appId = process.env.BOT_APP_ID;
    const appPassword = process.env.BOT_APP_PASSWORD;
    if (!appId || !appPassword) throw new Error("BOT_APP_ID or BOT_APP_PASSWORD is not set");
    adapter = new BotFrameworkAdapter({ appId, appPassword });
  }
  return adapter;
}

const FEEDBACK_ACTION_TO_STATUS: Partial<Record<FeedbackAction, "acknowledged" | "dismissed" | "escalated">> = {
  scope_violation_acknowledged: "acknowledged",
  false_positive_reported: "dismissed",
  scope_violation_escalated: "escalated",
};

/** Adapts Azure Functions HttpRequest into the Bot Framework WebRequest shape. */
async function toWebRequest(req: HttpRequest): Promise<WebRequest> {
  const body = await req.text();
  return {
    body,
    headers: Object.fromEntries(req.headers.entries()),
    method: req.method,
    originalUrl: req.url,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as unknown as WebRequest;
}

/**
 * HTTP trigger — receives Bot Framework activity from Teams (Adaptive Card submits and messages).
 */
async function handleTeamsFeedbackHandler(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const botAdapter = getAdapter();
  let statusCode = 200;

  const webReq = await toWebRequest(req);
  const webRes: WebResponse = {
    status(code: number) { statusCode = code; return this; },
    send() { return this; },
    end() { return; },
  } as unknown as WebResponse;

  await botAdapter.processActivity(webReq, webRes, async (turnContext: TurnContext) => {
    const activity = turnContext.activity as unknown as BotActivityPayload;

    // Handle Adaptive Card submit actions
    if (activity.type === "message" && activity.value) {
      const value = activity.value as {
        action?: FeedbackAction;
        violationId?: string;
        projectId?: string;
        comment?: string;
      };

      if (value.action && value.violationId && value.projectId) {
        await handleViolationFeedback(turnContext, value as Required<typeof value>, context);
        return;
      }
    }

    // Handle plain text commands
    if (activity.type === "message" && activity.text) {
      await handleTextCommand(turnContext, activity.text.trim().toLowerCase(), activity, context);
      return;
    }

    await turnContext.sendActivity("I received your message but wasn't sure how to handle it. Try `/status <projectId>`.");
  });

  return { status: statusCode };
}

async function handleViolationFeedback(
  ctx: TurnContext,
  value: { action: FeedbackAction; violationId: string; projectId: string; comment?: string },
  context: InvocationContext
): Promise<void> {
  const { action, violationId, projectId, comment } = value;
  const activity = ctx.activity as unknown as BotActivityPayload;

  const newStatus = FEEDBACK_ACTION_TO_STATUS[action];
  if (!newStatus) {
    await ctx.sendActivity("Unknown action. No changes made.");
    return;
  }

  try {
    await updateViolationStatus(violationId, projectId, {
      status: newStatus,
      acknowledgedBy: activity.from?.name ?? "Unknown",
      acknowledgedAt: new Date().toISOString(),
    });

    await insertFeedbackLog({
      feedbackId: uuidv4(),
      projectId,
      violationId,
      action,
      actorId: activity.from?.aadObjectId ?? activity.from?.id ?? "",
      actorName: activity.from?.name ?? "Unknown",
      actorEmail: "",
      comment,
      occurredAt: new Date().toISOString(),
      teamsConversationId: activity.conversation?.id,
      teamsActivityId: activity.id,
    });

    context.log(`Violation ${violationId} marked as ${newStatus} by ${activity.from?.name}`);
    const card = buildConfirmationCard(`Violation ${newStatus} successfully.`, true);
    await ctx.sendActivity({ attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] });
  } catch (err) {
    context.error(`Failed to update violation ${violationId}:`, err);
    const card = buildConfirmationCard("Failed to update violation. Please try again.", false);
    await ctx.sendActivity({ attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] });
  }
}

async function handleTextCommand(
  ctx: TurnContext,
  text: string,
  activity: BotActivityPayload,
  context: InvocationContext
): Promise<void> {
  const parts = text.split(/\s+/);
  const command = parts[0];
  const projectId = parts[1] ?? process.env.DEFAULT_PROJECT_ID ?? "";

  if (command === "/status" || command === "status") {
    await handleStatusCommand(ctx, projectId, context);
    return;
  }

  if (command === "/violations" || command === "violations") {
    await handleViolationsCommand(ctx, projectId, context);
    return;
  }

  await ctx.sendActivity(
    "Available commands:\n" +
    "• `/status <projectId>` — Show current scope risk status\n" +
    "• `/violations <projectId>` — List pending violations"
  );
}

async function handleStatusCommand(ctx: TurnContext, projectId: string, context: InvocationContext): Promise<void> {
  if (!projectId) {
    await ctx.sendActivity("Please provide a project ID. Example: `/status my-project`");
    return;
  }
  try {
    const summary = await getScopeSummary(projectId);
    if (!summary) {
      await ctx.sendActivity(`No scope summary found for project \`${projectId}\`.`);
      return;
    }
    const card = buildScopeStatusCard(
      summary.projectName,
      summary.riskScore,
      summary.violationCount,
      summary.lastAnalyzedAt
    );
    await ctx.sendActivity({ attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }] });
  } catch (err) {
    context.error("Error fetching status:", err);
    await ctx.sendActivity("Failed to retrieve scope status. Please try again.");
  }
}

async function handleViolationsCommand(ctx: TurnContext, projectId: string, context: InvocationContext): Promise<void> {
  if (!projectId) {
    await ctx.sendActivity("Please provide a project ID. Example: `/violations my-project`");
    return;
  }
  try {
    const violations = await listViolations(projectId, "pending");
    if (violations.length === 0) {
      await ctx.sendActivity(`✅ No pending violations for project \`${projectId}\`.`);
      return;
    }
    const list = violations
      .slice(0, 10)
      .map((v, i) => `${i + 1}. **[${v.severity.toUpperCase()}]** ${v.taskTitle}`)
      .join("\n");
    await ctx.sendActivity(`Found **${violations.length}** pending violation(s) for \`${projectId}\`:\n\n${list}`);
  } catch (err) {
    context.error("Error fetching violations:", err);
    await ctx.sendActivity("Failed to retrieve violations. Please try again.");
  }
}

app.http("HandleTeamsFeedback", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "bot/messages",
  handler: handleTeamsFeedbackHandler,
});
