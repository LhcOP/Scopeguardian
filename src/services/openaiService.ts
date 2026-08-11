import { AzureOpenAI } from "openai";
import { ProjectScope, ScopeItem, ScopeViolation } from "../models/ProjectScope";
import { TaskEvent } from "../models/TaskEvent";
import { countTokens } from "../utils/tokenOptimizer";
import { v4 as uuidv4 } from "uuid";

// Azure OpenAI uses deployment names — configurable per environment
const MINI_MODEL = process.env.AZURE_OPENAI_MINI_DEPLOYMENT ?? "gpt-4o-mini";
const FULL_MODEL = process.env.AZURE_OPENAI_CHAT_DEPLOYMENT ?? "gpt-4o";
const EMBEDDING_MODEL = process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT ?? "text-embedding-3-small";
const MAX_CHUNK_TOKENS = 6000;

let cachedClient: AzureOpenAI | null = null;

function getClient(): AzureOpenAI {
  if (!cachedClient) {
    const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
    const key = process.env.AZURE_OPENAI_KEY;
    const apiVersion = process.env.AZURE_OPENAI_API_VERSION ?? "2024-08-01-preview";
    if (!endpoint || !key) throw new Error("AZURE_OPENAI_ENDPOINT or AZURE_OPENAI_KEY is not set");
    cachedClient = new AzureOpenAI({ endpoint, apiKey: key, apiVersion });
  }
  return cachedClient;
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

  // REDUCE: synthesise findings with the full chat model
  return reduceToViolations(combinedAnalysis, taskContext, taskEvent, relevantScopeItems.map((i) => i.id));
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
          "You are a project scope and delivery analyst. Given a portion of the project scope and a task event, " +
          "identify scope creep and delivery-risk indicators. Prioritise these signals: " +
          "(1) work not covered by the scope items or matching the OUT OF SCOPE list, " +
          "(2) estimated hours increased compared to the previous value (see CHANGES), " +
          "(3) logged/registered hours exceeding the estimated hours, " +
          "(4) deadlines moved later than before, " +
          "(5) status/progress anomalies such as work started on tasks outside scope. " +
          "Monetary amounts are context only — hours, deadlines and progress are the primary signals. Be concise. " +
          'Output JSON: {"potentialCreep": boolean, "indicators": string[], "outOfScopeMatch": boolean, "hourOverrun": boolean, "deadlineSlip": boolean, "notes": string}',
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
  taskContext: string,
  taskEvent: TaskEvent,
  scopeItemIds: string[]
): Promise<ScopeViolation[]> {
  const client = getClient();

  const systemPrompt = `You are a senior project manager AI specialising in scope and delivery management.
Given aggregated analysis of a task event against project scope, determine:
1. Whether this constitutes a genuine scope violation or delivery risk.
2. The severity (low/medium/high/critical).
3. Which scope item IDs are affected.
4. A clear, actionable recommendation addressed to the consultant/project manager:
   state whether the change warrants a customer dialogue and a formal change
   order, and what exactly to clarify with the customer (hours impact, cause,
   who absorbs the extra effort). The goal is raising the flag EARLY — before
   the hours are spent — so the consultant can renegotiate instead of
   discovering the overrun at invoicing.

Judge severity primarily on hours, deadlines and progress. These rules are BINDING — when a threshold is met, report the violation even if it could also be read as a planning correction:
- Estimated hours grown ≥50% or by ≥8 hours versus the previous value → violation (medium; high when growth is ≥100% AND ≥16 hours).
- Logged/registered hours exceeding estimated hours → violation (medium; high when overrun ≥50%).
- Deadline moved later → violation (medium; high on required tasks or slips over a week).
- Work matching the OUT OF SCOPE list or not covered by any scope item → violation (high or critical).
- Monetary amounts alone are NOT the signal — translate everything to hours/deadline/progress impact.
Use the TASK EVENT numbers (including CHANGES previous → current) as the factual basis, not just the aggregated analysis prose.

Always respond with a JSON object of exactly this shape (violations may be empty):
{
  "violations": [
    {
      "severity": "low"|"medium"|"high"|"critical",
      "reasoning": "string",
      "affectedScopeItems": ["id1","id2"],
      "suggestedAction": "string"
    }
  ]
}
Only include real violations, not hypothetical concerns.`;

  const userPrompt = `TASK EVENT:\n${taskContext}\n\nAGGREGATED MAP ANALYSIS:\n${mapOutput}\n\nTASK ID: ${taskEvent.task.id}\nAVAILABLE SCOPE ITEM IDS: ${scopeItemIds.join(", ")}`;

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

  const raw = parseViolationPayload(response.choices[0].message.content ?? "{}");

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

export interface RawViolation {
  severity: ScopeViolation["severity"];
  reasoning: string;
  affectedScopeItems: string[];
  suggestedAction: string;
}

const VALID_SEVERITIES = new Set(["low", "medium", "high", "critical"]);

/**
 * Tolerant parser for the reduce model's response. With response_format
 * json_object the model cannot return a bare array, so it variously returns
 * {"violations":[...]}, a single violation object, or an array — accept all.
 */
export function parseViolationPayload(content: string): RawViolation[] {
  let parsed: unknown;
  try {
    parsed = JSON.parse(content);
  } catch {
    return [];
  }

  let candidates: unknown[];
  if (Array.isArray(parsed)) {
    candidates = parsed;
  } else if (parsed && typeof parsed === "object") {
    const obj = parsed as Record<string, unknown>;
    if (Array.isArray(obj.violations)) {
      candidates = obj.violations;
    } else {
      candidates = [obj]; // single violation object without wrapper
    }
  } else {
    return [];
  }

  return candidates.filter(
    (v): v is RawViolation =>
      !!v &&
      typeof v === "object" &&
      VALID_SEVERITIES.has(String((v as Record<string, unknown>).severity)) &&
      typeof (v as Record<string, unknown>).reasoning === "string"
  );
}

// ── Automatic Scope Generation ────────────────────────────────────────────────

/**
 * Generates a master scope from the project's own material (project
 * description, quote lines, estimation-basis documents). Returns null when
 * the material is insufficient to derive a meaningful scope.
 */
export async function generateScopeFromMaterial(
  projectId: string,
  projectName: string,
  material: string
): Promise<ProjectScope | null> {
  const client = getClient();

  const systemPrompt = `You are a project scoping analyst. From the provided project material
(project description, quote lines, estimation-basis documents) produce a master scope used for
automated scope-creep detection. Rules:
- Derive scope items ONLY from what the material actually covers — do not invent deliverables.
- Each scope item gets a short stable id (S1, S2, ...), title, description, deliverables and
  acceptance criteria; include estimatedHours when the material states hours per item.
- Put explicit exclusions from the material into outOfScope. Additionally add outOfScope entries
  for adjacent work the material clearly does NOT include (e.g. if the material describes
  configuration only, custom development is out of scope) — be conservative.
- Keep the content language of the source material (e.g. Danish stays Danish).
- If the material is too thin to derive at least one concrete scope item, respond {"insufficient": true}.

Respond with JSON:
{"projectName": "string", "scopeItems": [{"id","title","description","deliverables":[],"acceptanceCriteria":[],"estimatedHours":number|null,"tags":[]}], "outOfScope":[], "assumptions":[], "constraints":[]}`;

  const response = await client.chat.completions.create({
    model: FULL_MODEL,
    temperature: 0,
    max_tokens: 4096,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: `PROJECT: ${projectName} (${projectId})\n\nMATERIAL:\n${material}` },
    ],
  });

  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(response.choices[0].message.content ?? "{}") as Record<string, unknown>;
  } catch {
    return null;
  }
  if (parsed.insufficient === true) return null;

  const items = Array.isArray(parsed.scopeItems) ? (parsed.scopeItems as Partial<ScopeItem>[]) : [];
  const validItems = items.filter((i) => i?.id && i?.title && i?.description);
  if (validItems.length === 0) return null;

  const now = new Date().toISOString();
  return {
    projectId,
    projectName: typeof parsed.projectName === "string" && parsed.projectName ? parsed.projectName : projectName,
    version: "1",
    createdAt: now,
    updatedAt: now,
    scopeItems: validItems.map((i) => ({
      id: String(i.id),
      title: String(i.title),
      description: String(i.description),
      deliverables: Array.isArray(i.deliverables) ? i.deliverables.map(String) : [],
      acceptanceCriteria: Array.isArray(i.acceptanceCriteria) ? i.acceptanceCriteria.map(String) : [],
      estimatedHours: typeof i.estimatedHours === "number" ? i.estimatedHours : undefined,
      tags: Array.isArray(i.tags) ? i.tags.map(String) : [],
    })),
    outOfScope: Array.isArray(parsed.outOfScope) ? parsed.outOfScope.map(String) : [],
    assumptions: Array.isArray(parsed.assumptions) ? parsed.assumptions.map(String) : [],
    constraints: Array.isArray(parsed.constraints) ? parsed.constraints.map(String) : [],
  };
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
  const changes = event.changeDetails?.length
    ? "CHANGES (previous → current):\n" +
      event.changeDetails.map((c) => `  ${c.field}: ${c.oldValue ?? "(empty)"} → ${c.newValue ?? "(empty)"}`).join("\n")
    : "";
  return [
    `Event Type: ${event.eventType}`,
    `Task: ${event.task.title}`,
    `Description: ${event.task.description}`,
    event.task.status ? `Status: ${event.task.status}` : "",
    event.task.estimatedHours != null ? `Estimated Hours: ${event.task.estimatedHours}` : "",
    event.task.loggedHours != null ? `Logged/Registered Hours: ${event.task.loggedHours}` : "",
    event.task.deadline ? `Deadline: ${event.task.deadline}` : "",
    event.task.tags?.length ? `Tags: ${event.task.tags.join(", ")}` : "",
    changes,
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
