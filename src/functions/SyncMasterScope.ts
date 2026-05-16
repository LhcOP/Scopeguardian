import { app, Timer, InvocationContext } from "@azure/functions";
import { downloadMasterScope, uploadScopeMarkdown } from "../services/blobStorageService";
import { embedScopeItems } from "../services/openaiService";
import { generateScopeMarkdown, computeProjectRiskScore } from "../services/openaiService";
import { ensureIndexExists, upsertScopeItems, deleteScopeItemsByProject } from "../services/aiSearchService";
import { upsertScopeSummary, listViolations } from "../services/cosmosDbService";

const PROJECT_IDS = (process.env.PROJECT_IDS ?? "").split(",").map((id) => id.trim()).filter(Boolean);

/**
 * Timer-triggered function — runs every 6 hours.
 * Re-indexes the master scope into Azure AI Search and regenerates the Markdown summary.
 */
async function syncMasterScopeHandler(_timer: Timer, context: InvocationContext): Promise<void> {
  context.log(`SyncMasterScope triggered at ${new Date().toISOString()}`);
  await ensureIndexExists();

  for (const projectId of PROJECT_IDS) {
    context.log(`Syncing master scope for project: ${projectId}`);
    try {
      const scope = await downloadMasterScope(projectId);
      if (!scope) {
        context.warn(`No scope found for project ${projectId} — skipping`);
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

      // Compute risk score from recent violations
      const violations = await listViolations(projectId, "pending");
      const violationsSummary = violations
        .slice(0, 20)
        .map((v) => `[${v.severity}] ${v.taskTitle}: ${v.reasoning}`)
        .join("\n");
      const riskScore = violations.length > 0
        ? await computeProjectRiskScore(violationsSummary, scope.projectName)
        : 0;

      await upsertScopeSummary({
        projectId,
        projectName: scope.projectName,
        totalItems: scope.scopeItems.length,
        lastAnalyzedAt: new Date().toISOString(),
        violationCount: violations.length,
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
