import { app, InvocationContext } from "@azure/functions";
import { AnalysisQueueMessage } from "../models/TaskEvent";
import {
  getSubscriptionRecord,
  upsertSubscriptionRecord,
  tryMarkEventProcessed,
  unmarkEventProcessed,
} from "../services/cosmosDbService";
import {
  fetchListItemDelta,
  primeListItemDelta,
  mapListItemToTaskEvent,
} from "../services/graphService";
import { analyzeTaskEvent } from "../services/scopeAnalyzer";

/**
 * Queue-triggered worker — runs the delta query for the notified list,
 * dedupes item versions, and pushes each changed item through the
 * AI analysis pipeline. host.json caps queue concurrency at 1 so delta
 * links are consumed serially per project.
 */
async function processScopeAnalysisHandler(
  queueItem: unknown,
  context: InvocationContext
): Promise<void> {
  const msg = queueItem as AnalysisQueueMessage;
  if (!msg?.projectId || !msg.siteId || !msg.listId) {
    context.error("Invalid queue message — dropping:", JSON.stringify(queueItem));
    return;
  }

  const record = await getSubscriptionRecord(msg.projectId);
  if (!record) {
    context.warn(`No subscription record for project ${msg.projectId} — dropping message`);
    return;
  }

  // No delta link yet (e.g. record created by an older version): prime it so we
  // only analyse future changes instead of replaying the entire list history.
  if (!record.deltaLink) {
    const primed = await primeListItemDelta(msg.siteId, msg.listId);
    if (primed) await upsertSubscriptionRecord({ ...record, deltaLink: primed });
    context.warn(`Delta link primed for project ${msg.projectId} — this notification is skipped`);
    return;
  }

  const { items, deltaLink } = await fetchListItemDelta(msg.siteId, msg.listId, record.deltaLink);
  context.log(`Delta returned ${items.length} changed item(s) for project ${msg.projectId}`);

  let analyzed = 0;
  let violations = 0;

  for (const item of items) {
    if (item.deleted) continue;

    const version = item.eTag ?? item.lastModifiedDateTime ?? "unknown";
    const isNew = await tryMarkEventProcessed(msg.projectId, item.id, version);
    if (!isNew) {
      context.log(`Item ${item.id} (${version}) already processed — skipping`);
      continue;
    }

    const taskEvent = mapListItemToTaskEvent(item, msg.projectId);
    if (!taskEvent) continue;

    try {
      violations += await analyzeTaskEvent(taskEvent, context);
      analyzed++;
    } catch (err) {
      // Roll back the dedup marker so the queue retry can re-analyse this item
      await unmarkEventProcessed(msg.projectId, item.id, version);
      context.error(`Analysis failed for item ${item.id} in project ${msg.projectId}:`, err);
      throw err;
    }
  }

  // Persist the new delta link only after successful processing; combined with
  // per-item dedup this gives at-least-once delivery without double analysis.
  if (deltaLink) {
    await upsertSubscriptionRecord({ ...record, deltaLink });
  }

  context.log(
    `Analysis complete for project ${msg.projectId}: ${analyzed} item(s) analysed, ${violations} violation(s)`
  );
}

app.storageQueue("ProcessScopeAnalysis", {
  queueName: "scope-analysis",
  connection: "AzureWebJobsStorage",
  handler: processScopeAnalysisHandler,
});
