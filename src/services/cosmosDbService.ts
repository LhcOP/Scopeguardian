import { CosmosClient, Container } from "@azure/cosmos";
import { ScopeViolation, ScopeSummary } from "../models/ProjectScope";
import { FeedbackLog } from "../models/FeedbackLog";
import { TaskEvent } from "../models/TaskEvent";

const DB_NAME = "scopeguardian";

function getClient(): CosmosClient {
  const endpoint = process.env.COSMOS_ENDPOINT;
  const key = process.env.COSMOS_KEY;
  if (!endpoint || !key) throw new Error("COSMOS_ENDPOINT or COSMOS_KEY is not set");
  return new CosmosClient({ endpoint, key });
}

async function getContainer(containerId: string): Promise<Container> {
  const client = getClient();
  const { database } = await client.databases.createIfNotExists({ id: DB_NAME });
  const { container } = await database.containers.createIfNotExists({
    id: containerId,
    partitionKey: { paths: ["/projectId"] },
  });
  return container;
}

// ── Violations ───────────────────────────────────────────────────────────────

export async function upsertViolation(violation: ScopeViolation): Promise<void> {
  const container = await getContainer("violations");
  await container.items.upsert({ ...violation, id: violation.violationId });
}

export async function getViolation(violationId: string, projectId: string): Promise<ScopeViolation | null> {
  const container = await getContainer("violations");
  try {
    const { resource } = await container.item(violationId, projectId).read<ScopeViolation>();
    return resource ?? null;
  } catch {
    return null;
  }
}

export async function listViolations(
  projectId: string,
  status?: ScopeViolation["status"]
): Promise<ScopeViolation[]> {
  const container = await getContainer("violations");
  const query = status
    ? `SELECT * FROM c WHERE c.projectId = @pid AND c.status = @status ORDER BY c.detectedAt DESC`
    : `SELECT * FROM c WHERE c.projectId = @pid ORDER BY c.detectedAt DESC`;
  const params = status
    ? [{ name: "@pid", value: projectId }, { name: "@status", value: status }]
    : [{ name: "@pid", value: projectId }];
  const { resources } = await container.items.query<ScopeViolation>({ query, parameters: params }).fetchAll();
  return resources;
}

export async function updateViolationStatus(
  violationId: string,
  projectId: string,
  patch: Partial<Pick<ScopeViolation, "status" | "acknowledgedBy" | "acknowledgedAt">>
): Promise<void> {
  const existing = await getViolation(violationId, projectId);
  if (!existing) throw new Error(`Violation ${violationId} not found`);
  const container = await getContainer("violations");
  await container.items.upsert({ ...existing, ...patch, id: violationId });
}

// ── Task Events ───────────────────────────────────────────────────────────────

export async function insertTaskEvent(event: TaskEvent): Promise<void> {
  const container = await getContainer("taskEvents");
  await container.items.create({ ...event, id: event.eventId });
}

export async function getRecentTaskEvents(projectId: string, hours = 24): Promise<TaskEvent[]> {
  const container = await getContainer("taskEvents");
  const since = new Date(Date.now() - hours * 3_600_000).toISOString();
  const { resources } = await container.items
    .query<TaskEvent>({
      query: `SELECT * FROM c WHERE c.projectId = @pid AND c.occurredAt >= @since ORDER BY c.occurredAt DESC`,
      parameters: [
        { name: "@pid", value: projectId },
        { name: "@since", value: since },
      ],
    })
    .fetchAll();
  return resources;
}

// ── Feedback Logs ─────────────────────────────────────────────────────────────

export async function insertFeedbackLog(log: FeedbackLog): Promise<void> {
  const container = await getContainer("feedbackLogs");
  await container.items.create({ ...log, id: log.feedbackId });
}

// ── Scope Summaries ───────────────────────────────────────────────────────────

export async function upsertScopeSummary(summary: ScopeSummary): Promise<void> {
  const container = await getContainer("scopeSummaries");
  await container.items.upsert({ ...summary, id: summary.projectId });
}

export async function getScopeSummary(projectId: string): Promise<ScopeSummary | null> {
  const container = await getContainer("scopeSummaries");
  try {
    const { resource } = await container.item(projectId, projectId).read<ScopeSummary>();
    return resource ?? null;
  } catch {
    return null;
  }
}
