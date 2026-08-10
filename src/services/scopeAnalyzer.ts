import { InvocationContext } from "@azure/functions";
import { TaskEvent } from "../models/TaskEvent";
import { ScopeItem } from "../models/ProjectScope";
import { downloadMasterScope } from "./blobStorageService";
import { vectorSearch } from "./aiSearchService";
import { embedText, detectScopeViolations } from "./openaiService";
import {
  insertTaskEvent,
  upsertViolation,
  upsertScopeSummary,
  listViolations,
} from "./cosmosDbService";
import { buildViolationAlertCard } from "../utils/adaptiveCardBuilder";
import { sendTeamsAlert } from "./teamsNotifier";
import { selectItemsWithinBudget } from "../utils/tokenOptimizer";
import { computeRiskScore } from "../utils/riskScore";

const MAX_SCOPE_ITEMS_TOKENS = 4000;

/**
 * Core analysis pipeline: persists the event, finds the relevant scope items
 * via vector search, runs violation detection, alerts Teams, and refreshes
 * the project summary. Returns the number of violations detected.
 */
export async function analyzeTaskEvent(
  taskEvent: TaskEvent,
  context: InvocationContext
): Promise<number> {
  const { projectId } = taskEvent;

  await insertTaskEvent(taskEvent);

  const scope = await downloadMasterScope(projectId);
  if (!scope) {
    context.warn(`No master scope found for project ${projectId} — skipping analysis`);
    return 0;
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

  const violations = await detectScopeViolations(taskEvent, selectedItems, scope.outOfScope);

  const teamsChannelId = process.env.TEAMS_CHANNEL_ID ?? "";
  for (const violation of violations) {
    await upsertViolation(violation);
    context.log(`Violation detected: ${violation.violationId} (${violation.severity})`);
    const card = buildViolationAlertCard(violation, scope.projectName);
    await sendTeamsAlert(teamsChannelId, card, context);
  }

  if (violations.length > 0) {
    const pending = await listViolations(projectId, "pending");
    await upsertScopeSummary({
      projectId,
      projectName: scope.projectName,
      totalItems: scope.scopeItems.length,
      lastAnalyzedAt: new Date().toISOString(),
      violationCount: pending.length,
      riskScore: computeRiskScore(pending),
    });
  }

  return violations.length;
}
