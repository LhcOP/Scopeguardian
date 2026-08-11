import { app, InvocationContext } from "@azure/functions";
import { v4 as uuidv4 } from "uuid";
import { AnalysisQueueMessage, TaskEvent } from "../models/TaskEvent";
import {
  getSubscriptionRecord,
  getSubscriptionRecordById,
  docsRecordId,
  upsertSubscriptionRecord,
  tryMarkEventProcessed,
  unmarkEventProcessed,
  SubscriptionRecord,
} from "../services/cosmosDbService";
import {
  fetchListItemDelta,
  primeListItemDelta,
  mapListItemToTaskEvent,
  fetchDriveDelta,
  primeDriveDelta,
  downloadDriveItemContent,
} from "../services/graphService";
import { analyzeTaskEvent, analyzeDocumentEvent } from "../services/scopeAnalyzer";
import { bootstrapMasterScope } from "../services/scopeBootstrapper";
import { downloadMasterScope } from "../services/blobStorageService";
import { extractTextFromFile, isSupportedDocument } from "../utils/documentText";

const MAX_DOC_FILE_BYTES = 15 * 1024 * 1024;
const MAX_DOC_TEXT_CHARS = 40_000;

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

  if (msg.resource === "drive") {
    await processDriveDelta(msg, context);
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

  // Auto-generate the master scope from project material on first activity
  if (items.length > 0 && !(await downloadMasterScope(msg.projectId))) {
    await bootstrapMasterScope(msg.projectId, msg.siteId, context);
  }

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

/**
 * Drive branch: delta over the project's document library — new/changed
 * documents (pdf, Word, emails, spreadsheets) are text-extracted and
 * analysed against the master scope.
 */
async function processDriveDelta(msg: AnalysisQueueMessage, context: InvocationContext): Promise<void> {
  const record = await getSubscriptionRecordById(docsRecordId(msg.projectId), msg.projectId);
  if (!record?.driveId) {
    context.warn(`No document subscription record for project ${msg.projectId} — dropping message`);
    return;
  }

  if (!record.deltaLink) {
    const primed = await primeDriveDelta(record.driveId);
    if (primed) await upsertSubscriptionRecord({ ...record, deltaLink: primed });
    context.warn(`Drive delta primed for project ${msg.projectId} — this notification is skipped`);
    return;
  }

  const { items, deltaLink } = await fetchDriveDelta(record.driveId, record.deltaLink);
  const files = items.filter((i) => i.file && !i.deleted && i.name);
  context.log(`Drive delta returned ${files.length} changed file(s) for project ${msg.projectId}`);

  let analyzed = 0;
  let violations = 0;

  for (const item of files) {
    const name = item.name ?? "";
    if (!isSupportedDocument(name)) {
      context.log(`Skipping unsupported document type: ${name}`);
      continue;
    }
    if ((item.size ?? 0) > MAX_DOC_FILE_BYTES) {
      context.warn(`Skipping oversized document: ${name} (${item.size} bytes)`);
      continue;
    }

    const version = item.eTag ?? item.lastModifiedDateTime ?? "unknown";
    const isNew = await tryMarkEventProcessed(msg.projectId, `doc-${item.id}`, version);
    if (!isNew) continue;

    try {
      const buffer = await downloadDriveItemContent(record.driveId, item.id);
      const text = await extractTextFromFile(name, buffer);
      if (!text?.trim()) {
        context.log(`No text extracted from ${name} — skipping`);
        continue;
      }

      const taskEvent: TaskEvent = {
        eventId: uuidv4(),
        projectId: msg.projectId,
        eventType: "document_added",
        occurredAt: new Date().toISOString(),
        source: "sharepoint",
        task: {
          id: `doc-${item.id}`,
          title: `Dokument: ${name}`,
          description: text.slice(0, 500),
          tags: ["document"],
        },
      };

      violations += await analyzeDocumentEvent(taskEvent, text.slice(0, MAX_DOC_TEXT_CHARS), context);
      analyzed++;
    } catch (err) {
      await unmarkEventProcessed(msg.projectId, `doc-${item.id}`, version);
      context.error(`Document analysis failed for ${name} in project ${msg.projectId}:`, err);
      throw err;
    }
  }

  if (deltaLink) {
    await upsertSubscriptionRecord({ ...record, deltaLink } as SubscriptionRecord);
  }

  context.log(
    `Document analysis complete for project ${msg.projectId}: ${analyzed} document(s) analysed, ${violations} violation(s)`
  );
}

app.storageQueue("ProcessScopeAnalysis", {
  queueName: "scope-analysis",
  connection: "AzureWebJobsStorage",
  handler: processScopeAnalysisHandler,
});
