import { AzureOpenAI } from "openai";
import { ScopeItem, ScopeViolation } from "../models/ProjectScope";
import { TaskEvent } from "../models/TaskEvent";
import { countTokens } from "../utils/tokenOptimizer";
import { v4 as uuidv4 } from "uuid";

const MINI_MODEL = "gpt-4o-mini";
const FULL_MODEL = "gpt-4o";
const EMBEDDING_MODEL = "text-embedding-ada-002";
const MAX_CHUNK_TOKENS = 6000;

function getClient(): AzureOpenAI {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const key = process.env.AZURE_OPENAI_KEY;
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION ?? "2024-08-01-preview";
  if (!endpoint || !key) throw new Error("AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_KEY is not set");
  return new AzureOpenAI({ endpoint, apiKey: key, apiVersion });
}

// ── Embeddings ────────────────────────────────────────────────────────────────

export async function embedText(texts: string[]): Promise<number[][]> {
  const client = getClient();
  const response = await client.embeddings.create({ model: EMBEDDING_MODEL, input: texts });
  return response.data.map((d) => d.embedding);
}

export async function embedScopeItems(items: ScopeItem[]): Promise<number[][]> {
  const texts = items.map(
    (item) => `${item.title}\n${item.description}\nDeliverables: ${item.deliverables.join(", ")}`
  );
  return embedText(texts);
}

// ── Map-Reduce Scope Violation Detection ──────────────────────────────────────

export async function detectScopeViolations(
  taskEvent: TaskEvent,
  relevantScopeItems: ScopeItem[],
  projectOutOfScope: string[]
): Promise<ScopeViolation[]> {
  const scopeContext = buildScopeContext(relevantScopeItems, projectOutOfScope);
  const taskContext = buildTaskContext(taskEvent);

  // MAP: chunk scope context if large, analyse each chunk with mini model
  const chunks = chunkText(scopeContext, MAX_CHUNK_TOKENS);
  const mapResults = await Promise.all(
    chunks.map((chunk) => mapAnalysis(chunk, taskContext))
  );

  const combinedAnalysis = mapResults.join("\n\n---\n\n");

  // REDUCE: synthesise findings with GPT-4o
  return reduceToViolations(combinedAnalysis, taskEvent, relevantScopeItems.map((i) => i.id));
}

async function mapAnalysis(scopeChunk: string, taskContext: string): Promise<string> {
  const client = getClient();
  const response = await client.chat.completions.create({
    model: MINI_MODEL,
    temperature: 0,
    max_tokens: 1024,
    messages: [
      {
        role: "system",
        content:
          "You are a project scope analyst. Given a portion of the project scope and a task event, " +
          "identify any potential scope creep indicators. Be concise. Output JSON with fields: " +
          '{"potentialCreep": boolean, "indicators": string[], "outOfScopeMatch": boolean, "notes": string}',
      },
      {
        role: "user",
        content: `SCOPE PORTION:\n${scopeChunk}\n\nTASK EVENT:\n${taskContext}`,
      },
    ],
  });
  return response.choices[0].message.content ?? "";
}

async function reduceToViolations(
  mapOutput: string,
  taskEvent: TaskEvent,
  scopeItemIds: string[]
): Promise<ScopeViolation[]> {
  const client = getClient();

  const systemPrompt = `You are a senior project manager AI specialising in scope management.
Given aggregated analysis of a task event against project scope, determine:
1. Whether this constitutes a genuine scope violation.
2. The severity (low/medium/high/critical).
3. Which scope item IDs are affected.
4. A clear, actionable recommendation.

Respond with a JSON array of violations (can be empty). Each violation:
{
  "severity": "low"|"medium"|"high"|"critical",
  "reasoning": "string",
  "affectedScopeItems": ["id1","id2"],
  "suggestedAction": "string"
}
Only include real violations, not hypothetical concerns.`;

  const userPrompt = `AGGREGATED MAP ANALYSIS:\n${mapOutput}\n\nTASK ID: ${taskEvent.task.id}\nTASK TITLE: ${taskEvent.task.title}\nAVAILABLE SCOPE ITEM IDS: ${scopeItemIds.join(", ")}`;

  const response = await client.chat.completions.create({
    model: FULL_MODEL,
    temperature: 0,
    max_tokens: 2048,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
  });

  type ViolationResult = {
    violations?: {
      severity: ScopeViolation["severity"];
      reasoning: string;
      affectedScopeItems: string[];
      suggestedAction: string;
    }[];
  };
  const parsed = JSON.parse(response.choices[0].message.content ?? "{}") as ViolationResult;
  const raw = parsed.violations ?? [];

  return raw.map((v) => ({
    violationId: uuidv4(),
    projectId: taskEvent.projectId,
    detectedAt: new Date().toISOString(),
    severity: v.severity,
    taskEventId: taskEvent.eventId,
    taskTitle: taskEvent.task.title,
    taskDescription: taskEvent.task.description,
    reasoning: v.reasoning,
    affectedScopeItems: v.affectedScopeItems,
    suggestedAction: v.suggestedAction,
    status: "pending",
  }));
}

// ── Scope Summary Generation ───────────────────────────────────────────────────

export async function generateScopeMarkdown(
  projectName: string,
  items: ScopeItem[],
  outOfScope: string[],
  assumptions: string[],
  constraints: string[]
): Promise<string> {
  const client = getClient();
  const scopeJson = JSON.stringify({ items, outOfScope, assumptions, constraints }, null, 2);

  const response = await client.chat.completions.create({
    model: MINI_MODEL,
    temperature: 0.2,
    max_tokens: 3000,
    messages: [
      {
        role: "system",
        content:
          "You are a technical writer. Convert the provided project scope JSON into a well-structured " +
          "Markdown document suitable for non-technical stakeholders. Include sections for Overview, " +
          "In-Scope Items, Out-of-Scope, Assumptions, and Constraints.",
      },
      { role: "user", content: `Project: ${projectName}\n\nScope Data:\n${scopeJson}` },
    ],
  });

  return response.choices[0].message.content ?? "";
}

// ── Risk Scoring ──────────────────────────────────────────────────────────────

export async function computeProjectRiskScore(
  recentViolationsSummary: string,
  projectName: string
): Promise<number> {
  const client = getClient();
  const response = await client.chat.completions.create({
    model: MINI_MODEL,
    temperature: 0,
    max_tokens: 64,
    messages: [
      {
        role: "system",
        content:
          "You are a risk analyst. Given a summary of recent scope violations, return a single integer " +
          "risk score from 0 (no risk) to 100 (critical risk). Respond with only the integer.",
      },
      { role: "user", content: `Project: ${projectName}\n\nRecent Violations:\n${recentViolationsSummary}` },
    ],
  });
  const score = parseInt(response.choices[0].message.content?.trim() ?? "0", 10);
  return isNaN(score) ? 0 : Math.min(100, Math.max(0, score));
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function buildScopeContext(items: ScopeItem[], outOfScope: string[]): string {
  const itemsText = items
    .map(
      (i) =>
        `[${i.id}] ${i.title}\n  Description: ${i.description}\n  Deliverables: ${i.deliverables.join(", ")}\n  Criteria: ${i.acceptanceCriteria.join(", ")}`
    )
    .join("\n\n");
  const oosText = outOfScope.map((o) => `- ${o}`).join("\n");
  return `IN SCOPE:\n${itemsText}\n\nOUT OF SCOPE:\n${oosText}`;
}

function buildTaskContext(event: TaskEvent): string {
  return [
    `Event Type: ${event.eventType}`,
    `Task: ${event.task.title}`,
    `Description: ${event.task.description}`,
    event.task.tags?.length ? `Tags: ${event.task.tags.join(", ")}` : "",
    event.task.estimatedHours != null ? `Estimated Hours: ${event.task.estimatedHours}` : "",
    event.changeDetails?.map((c) => `Changed ${c.field}: ${c.oldValue} → ${c.newValue}`).join("\n") ?? "",
  ]
    .filter(Boolean)
    .join("\n");
}

function chunkText(text: string, maxTokens: number): string[] {
  const words = text.split(" ");
  const chunks: string[] = [];
  let current: string[] = [];
  let tokenCount = 0;

  for (const word of words) {
    const wordTokens = countTokens(word);
    if (tokenCount + wordTokens > maxTokens && current.length > 0) {
      chunks.push(current.join(" "));
      current = [];
      tokenCount = 0;
    }
    current.push(word);
    tokenCount += wordTokens;
  }
  if (current.length > 0) chunks.push(current.join(" "));
  return chunks;
}
