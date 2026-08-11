import { InvocationContext } from "@azure/functions";
import { TaskEvent } from "../models/TaskEvent";
import { ScopeItem } from "../models/ProjectScope";
import { downloadMasterScope } from "./blobStorageService";
import { vectorSearch } from "./aiSearchService";
import { embedText, detectScopeViolations, detectScopeViolationsInDocument } from "./openaiService";
import {
  insertTaskEvent,
  getLastTaskEventForTask,
  upsertViolation,
  upsertScopeSummary,
  listViolations,
} from "./cosmosDbService";
import { buildViolationAlertCard } from "../utils/adaptiveCardBuilder";
import { sendTeamsAlert } from "./teamsNotifier";
import { sendViolationEmail } from "./emailNotifier";
import { selectItemsWithinBudget } from "../utils/tokenOptimizer";
import { computeRiskScore } from "../utils/riskScore";
import { getRegisteredHours } from "./timeTrackingService";

const MAX_SCOPE_ITEMS_TOKENS = 4000;

const DIFF_FIELDS = ["title", "estimatedHours", "loggedHours", "deadline", "status", "description"] as const;

function diffTasks(
  prev: TaskEvent["task"],
  current: TaskEvent["task"]
): NonNullable<TaskEvent["changeDetails"]> {
  const changes: NonNullable<TaskEvent["changeDetails"]> = [];
  for (const field of DIFF_FIELDS) {
    const oldValue = prev[field] != null ? String(prev[field]) : null;
    const newValue = current[field] != null ? String(current[field]) : null;
    if (oldValue !== newValue) {
      changes.push({ field, oldValue, newValue });
    }
  }
  return changes;
}

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

  // Enrich with actual registered hours from the central time-tracking hub —
  // the task list itself carries no logged hours in ProjektPoint setups.
  if (taskEvent.task.loggedHours == null) {
    const registered = await getRegisteredHours(projectId, taskEvent.task.title, context).catch(
      () => null
    );
    if (registered && registered.taskHours > 0) {
      taskEvent.task.loggedHours = registered.taskHours;
      context.log(
        `Registered hours for "${taskEvent.task.title}": ${registered.taskHours} (project total: ${registered.projectHours})`
      );
    }
  }

  // Diff against the previous version of this task so the analysis sees what
  // actually changed — hour growth, deadline slips and status transitions are
  // the primary scope-creep signals.
  const previous = await getLastTaskEventForTask(projectId, taskEvent.task.id);
  if (previous) {
    taskEvent.changeDetails = diffTasks(previous.task, taskEvent.task);
  }

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
    if (teamsChannelId) {
      const card = buildViolationAlertCard(violation, scope.projectName);
      await sendTeamsAlert(teamsChannelId, card, context);
    }
    await sendViolationEmail(violation, scope.projectName, context);
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

/**
 * Analyses a document added to the project (email, minutes, requirement doc)
 * against the master scope. Persists the event, alerts on violations, and
 * refreshes the summary — mirrors analyzeTaskEvent for document content.
 */
export async function analyzeDocumentEvent(
  taskEvent: TaskEvent,
  documentText: string,
  context: InvocationContext
): Promise<number> {
  const { projectId } = taskEvent;

  await insertTaskEvent(taskEvent);

  const scope = await downloadMasterScope(projectId);
  if (!scope) {
    context.warn(`No master scope found for project ${projectId} — skipping document analysis`);
    return 0;
  }

  // Vector search on the document excerpt to find the most relevant scope items
  const queryText = `${taskEvent.task.title} ${documentText.slice(0, 2000)}`;
  const [queryVector] = await embedText([queryText]);
  const searchHits = await vectorSearch(projectId, queryVector, 10);
  const relevantIds = new Set(searchHits.map((h) => h.item.id?.replace(`${projectId}-`, "")));

  let relevantItems = scope.scopeItems.filter((i) => relevantIds.has(i.id));
  if (relevantItems.length === 0) relevantItems = scope.scopeItems.slice(0, 5);
  const selectedItems = selectItemsWithinBudget<ScopeItem>(
    relevantItems,
    (i) => `${i.title} ${i.description} ${i.deliverables.join(" ")}`,
    MAX_SCOPE_ITEMS_TOKENS
  );

  const violations = await detectScopeViolationsInDocument(
    taskEvent,
    documentText,
    selectedItems,
    scope.outOfScope
  );

  const teamsChannelId = process.env.TEAMS_CHANNEL_ID ?? "";
  for (const violation of violations) {
    await upsertViolation(violation);
    context.log(`Document violation detected: ${violation.violationId} (${violation.severity})`);
    if (teamsChannelId) {
      const card = buildViolationAlertCard(violation, scope.projectName);
      await sendTeamsAlert(teamsChannelId, card, context);
    }
    await sendViolationEmail(violation, scope.projectName, context);
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
