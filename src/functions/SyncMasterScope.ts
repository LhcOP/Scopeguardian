import { app, Timer, InvocationContext } from "@azure/functions";
import { v4 as uuidv4 } from "uuid";
import { downloadMasterScope, uploadScopeMarkdown } from "../services/blobStorageService";
import { embedScopeItems, generateScopeMarkdown } from "../services/openaiService";
import { ensureIndexExists, upsertScopeItems, deleteScopeItemsByProject } from "../services/aiSearchService";
import { upsertScopeSummary, listViolations, tryMarkEventProcessed } from "../services/cosmosDbService";
import { parseProjectConfigs, ProjectConfig } from "../utils/projectConfig";
import { computeRiskScore } from "../utils/riskScore";
import { bootstrapMasterScope } from "../services/scopeBootstrapper";
import { getRegisteredHours } from "../services/timeTrackingService";
import { getListItemFields } from "../services/graphService";
import { analyzeTaskEvent } from "../services/scopeAnalyzer";
import { TaskEvent } from "../models/TaskEvent";

/**
 * Timer-triggered function — runs every 6 hours.
 * Re-indexes the master scope into Azure AI Search and regenerates the Markdown summary.
 */
async function syncMasterScopeHandler(_timer: Timer, context: InvocationContext): Promise<void> {
  context.log(`SyncMasterScope triggered at ${new Date().toISOString()}`);
  await ensureIndexExists();

  const configs = parseProjectConfigs();
  if (configs.length === 0) {
    context.warn("PROJECT_CONFIGS is empty — nothing to sync");
    return;
  }

  for (const config of configs) {
    const projectId = config.projectId;
    context.log(`Syncing master scope for project: ${projectId}`);
    try {
      let scope = await downloadMasterScope(projectId);
      if (!scope) {
        // Auto-generate from project material — no consultant involvement
        scope = await bootstrapMasterScope(projectId, config.siteId, context);
      }
      if (!scope) {
        context.warn(`No scope and no material for project ${projectId} — skipping`);
        continue;
      }

      // Re-index scope items (delete stale docs first to handle removed items)
      await deleteScopeItemsByProject(projectId);
      const embeddings = await embedScopeItems(scope.scopeItems);
      await upsertScopeItems(projectId, scope.scopeItems, embeddings);
      context.log(`Indexed ${scope.scopeItems.length} scope items for ${projectId}`);

      // Generate human-readable Markdown summary
      const markdown = await generateScopeMarkdown(
        scope.projectName,
        scope.scopeItems,
        scope.outOfScope,
        scope.assumptions,
        scope.constraints
      );
      await uploadScopeMarkdown(projectId, markdown);
      context.log(`Scope Markdown summary uploaded for ${projectId}`);

      // Refresh summary with the shared deterministic risk score
      const pending = await listViolations(projectId, "pending");
      const riskScore = computeRiskScore(pending);

      await upsertScopeSummary({
        projectId,
        projectName: scope.projectName,
        totalItems: scope.scopeItems.length,
        lastAnalyzedAt: new Date().toISOString(),
        violationCount: pending.length,
        riskScore,
      });

      context.log(`Scope sync complete for ${projectId}. Risk score: ${riskScore}`);

      // Sweep registered hours: catch overruns even when nobody edits the task
      await sweepRegisteredHours(config, context);
    } catch (err) {
      context.error(`Failed to sync scope for project ${projectId}:`, err);
    }
  }
}

/**
 * Compares registered hours (time-tracking hub) against task estimates and
 * pushes overruns through the analysis pipeline. Dedup markers keyed on the
 * registered total ensure each overrun level only alerts once.
 */
async function sweepRegisteredHours(config: ProjectConfig, context: InvocationContext): Promise<void> {
  if (!process.env.TIMETRACK_SITE_ID) return;
  const { projectId } = config;

  const registered = await getRegisteredHours(projectId, undefined, context);
  if (!registered || registered.byTask.size === 0) return;

  let fieldMap: Record<string, string> = {};
  try {
    fieldMap = JSON.parse(process.env.TASK_FIELD_MAP ?? "{}") as Record<string, string>;
  } catch {
    // fall through to defaults
  }
  const titleField = fieldMap.title ?? "Title";
  const hoursField = fieldMap.estimatedHours ?? "EstimatedHours";
  const idField = "id";

  const taskRows = await getListItemFields(config.siteId, config.listId);
  let overruns = 0;

  for (const row of taskRows) {
    const rawTitle = row[titleField];
    const title = typeof rawTitle === "string" ? rawTitle : null;
    const estimated = Number(row[hoursField]);
    if (!title || !isFinite(estimated) || estimated <= 0) continue;

    const logged = registered.byTask.get(title.trim().toLowerCase()) ?? 0;
    if (logged <= estimated) continue;

    const taskId = String(row[idField] ?? title);
    // One alert per overrun level — re-alerts only when more hours are logged
    const isNew = await tryMarkEventProcessed(projectId, `timelog-${taskId}`, String(logged));
    if (!isNew) continue;

    overruns++;
    const taskEvent: TaskEvent = {
      eventId: uuidv4(),
      projectId,
      eventType: "time_logged",
      occurredAt: new Date().toISOString(),
      source: "sharepoint",
      task: {
        id: taskId,
        title,
        description: `Registered hours (${logged}) exceed the estimate (${estimated}).`,
        estimatedHours: estimated,
        loggedHours: logged,
        tags: [],
      },
    };
    try {
      await analyzeTaskEvent(taskEvent, context);
    } catch (err) {
      context.error(`Overrun analysis failed for task "${title}" in ${projectId}:`, err);
    }
  }

  if (overruns > 0) {
    context.log(`Registered-hours sweep for ${projectId}: ${overruns} new overrun(s) analysed`);
  }
}

app.timer("SyncMasterScope", {
  schedule: "0 0 */6 * * *", // Every 6 hours
  handler: syncMasterScopeHandler,
});
