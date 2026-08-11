import { app, output, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphWebhookNotification, AnalysisQueueMessage } from "../models/TaskEvent";
import { getSubscriptionBySubscriptionId } from "../services/cosmosDbService";

const WEBHOOK_CLIENT_STATE = process.env.WEBHOOK_CLIENT_STATE ?? "scopeguardian-secret";

export const scopeAnalysisQueue = output.storageQueue({
  queueName: "scope-analysis",
  connection: "AzureWebJobsStorage",
});

/**
 * HTTP trigger — receives Microsoft Graph change notifications for SharePoint
 * lists. Handles the validation handshake, resolves the notification to a
 * project via the stored subscription record, and enqueues an analysis job.
 * All heavy work happens in ProcessScopeAnalysis — Graph requires a response
 * within seconds or it retries and eventually suspends the subscription.
 */
async function eventTriggerHandler(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  // Graph API validation handshake
  const validationToken = req.query.get("validationToken");
  if (validationToken) {
    context.log("Responding to Graph subscription validation handshake");
    return {
      status: 200,
      headers: { "Content-Type": "text/plain" },
      body: validationToken,
    };
  }

  let notification: GraphWebhookNotification;
  try {
    notification = (await req.json()) as GraphWebhookNotification;
  } catch {
    return { status: 400, body: "Invalid JSON payload" };
  }

  const messages: AnalysisQueueMessage[] = [];
  const seenSubscriptions = new Set<string>();

  for (const change of notification.value ?? []) {
    // Validate client state to prevent spoofed notifications
    if (change.clientState !== WEBHOOK_CLIENT_STATE) {
      context.warn(`Rejected notification with invalid clientState: ${change.clientState}`);
      continue;
    }

    // Graph batches multiple changes per subscription — one delta run covers them all
    if (seenSubscriptions.has(change.subscriptionId)) continue;
    seenSubscriptions.add(change.subscriptionId);

    try {
      const record = await getSubscriptionBySubscriptionId(change.subscriptionId);
      if (!record) {
        context.warn(`No registered project for subscription ${change.subscriptionId} — skipping`);
        continue;
      }
      messages.push({
        projectId: record.projectId,
        siteId: record.siteId,
        listId: record.listId,
        subscriptionId: change.subscriptionId,
        notifiedAt: new Date().toISOString(),
        resource: record.resourceType ?? "list",
        driveId: record.driveId,
      });
      context.log(`Queued ${record.resourceType ?? "list"} analysis for project ${record.projectId}`);
    } catch (err) {
      context.error(`Failed to resolve subscription ${change.subscriptionId}:`, err);
    }
  }

  if (messages.length > 0) {
    context.extraOutputs.set(scopeAnalysisQueue, messages);
  }

  // Always return 202 so Graph doesn't retry
  return { status: 202 };
}

app.http("EventTrigger", {
  methods: ["GET", "POST"],
  authLevel: "function",
  route: "webhook/graph",
  extraOutputs: [scopeAnalysisQueue],
  handler: eventTriggerHandler,
});
