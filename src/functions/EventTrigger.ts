import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { GraphWebhookNotification } from "../models/TaskEvent";
import { mapNotificationToTaskEvent } from "../services/graphService";
import { insertTaskEvent } from "../services/cosmosDbService";
import { downloadMasterScope } from "../services/blobStorageService";
import { vectorSearch } from "../services/aiSearchService";
import { embedText } from "../services/openaiService";
import { detectScopeViolations } from "../services/openaiService";
import { upsertViolation, upsertScopeSummary, getScopeSummary } from "../services/cosmosDbService";
import { buildViolationAlertCard } from "../utils/adaptiveCardBuilder";
import { sendTeamsAlert } from "./shared/teamsNotifier";
import { selectItemsWithinBudget } from "../utils/tokenOptimizer";
import { ScopeItem } from "../models/ProjectScope";

const PROJECT_ID = process.env.DEFAULT_PROJECT_ID ?? "";
const TEAMS_CHANNEL_ID = process.env.TEAMS_CHANNEL_ID ?? "";
const TEAMS_TEAM_ID = process.env.TEAMS_TEAM_ID ?? "";
const WEBHOOK_CLIENT_STATE = process.env.WEBHOOK_CLIENT_STATE ?? "scopeguardian-secret";
const MAX_SCOPE_ITEMS_TOKENS = 4000;

/**
 * HTTP trigger — receives Microsoft Graph change notifications for SharePoint list items.
 * Handles both the initial validation handshake and live event payloads.
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

  for (const change of notification.value ?? []) {
    // Validate client state to prevent spoofed notifications
    if (change.clientState !== WEBHOOK_CLIENT_STATE) {
      context.warn(`Rejected notification with invalid clientState: ${change.clientState}`);
      continue;
    }

    const projectId = PROJECT_ID || change.tenantId;
    context.log(`Processing change notification for project ${projectId}`);

    try {
      const taskEvent = await mapNotificationToTaskEvent(change, projectId);
      if (!taskEvent) {
        context.warn("Could not map notification to task event — skipping");
        continue;
      }

      await insertTaskEvent(taskEvent);

      // Retrieve master scope
      const scope = await downloadMasterScope(projectId);
      if (!scope) {
        context.warn(`No master scope found for project ${projectId}`);
        continue;
      }

      // Vector search: find the most relevant scope items
      const queryText = `${taskEvent.task.title} ${taskEvent.task.description}`;
      const [queryVector] = await embedText([queryText]);
      const searchHits = await vectorSearch(projectId, queryVector, 10);
      const relevantIds = new Set(searchHits.map((h) => h.item.id?.replace(`${projectId}-`, "")));

      // Filter and budget scope items
      let relevantItems = scope.scopeItems.filter((i) => relevantIds.has(i.id));
      if (relevantItems.length === 0) relevantItems = scope.scopeItems.slice(0, 5);
      const selectedItems = selectItemsWithinBudget<ScopeItem>(
        relevantItems,
        (i) => `${i.title} ${i.description} ${i.deliverables.join(" ")}`,
        MAX_SCOPE_ITEMS_TOKENS
      );

      // Run map-reduce violation detection
      const violations = await detectScopeViolations(taskEvent, selectedItems, scope.outOfScope);

      for (const violation of violations) {
        await upsertViolation(violation);
        context.log(`Violation detected: ${violation.violationId} (${violation.severity})`);

        // Post to Teams
        const card = buildViolationAlertCard(violation, scope.projectName);
        await sendTeamsAlert(TEAMS_TEAM_ID, TEAMS_CHANNEL_ID, card, context);
      }

      // Update scope summary risk score
      if (violations.length > 0) {
        const existing = await getScopeSummary(projectId);
        const updatedViolationCount = (existing?.violationCount ?? 0) + violations.length;
        await upsertScopeSummary({
          projectId,
          projectName: scope.projectName,
          totalItems: scope.scopeItems.length,
          lastAnalyzedAt: new Date().toISOString(),
          violationCount: updatedViolationCount,
          riskScore: Math.min(100, updatedViolationCount * 10),
        });
      }
    } catch (err) {
      context.error(`Error processing notification for project ${projectId}:`, err);
    }
  }

  // Always return 202 so Graph doesn't retry
  return { status: 202 };
}

app.http("EventTrigger", {
  methods: ["GET", "POST"],
  authLevel: "function",
  route: "webhook/graph",
  handler: eventTriggerHandler,
});
