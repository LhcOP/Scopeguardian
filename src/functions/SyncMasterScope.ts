import { app, Timer, InvocationContext } from "@azure/functions";
import { downloadMasterScope, uploadScopeMarkdown } from "../services/blobStorageService";
import { embedScopeItems, generateScopeMarkdown } from "../services/openaiService";
import { ensureIndexExists, upsertScopeItems, deleteScopeItemsByProject } from "../services/aiSearchService";
import { upsertScopeSummary, listViolations } from "../services/cosmosDbService";
import { parseProjectConfigs } from "../utils/projectConfig";
import { computeRiskScore } from "../utils/riskScore";
import { bootstrapMasterScope } from "../services/scopeBootstrapper";

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
    } catch (err) {
      context.error(`Failed to sync scope for project ${projectId}:`, err);
    }
  }
}

app.timer("SyncMasterScope", {
  schedule: "0 0 */6 * * *", // Every 6 hours
  handler: syncMasterScopeHandler,
});
