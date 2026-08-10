import { CosmosClient, Container, Database } from "@azure/cosmos";
import { ScopeViolation, ScopeSummary } from "../models/ProjectScope";
import { FeedbackLog } from "../models/FeedbackLog";
import { TaskEvent } from "../models/TaskEvent";

const DB_NAME = "scopeguardian";
const PROCESSED_EVENT_TTL_SECONDS = 7 * 24 * 3600;

interface ContainerSpec {
  id: string;
  partitionKeyPath: string;
  defaultTtl?: number;
}

const CONTAINERS = {
  violations: { id: "violations", partitionKeyPath: "/projectId" },
  taskEvents: { id: "taskEvents", partitionKeyPath: "/projectId" },
  feedbackLogs: { id: "feedbackLogs", partitionKeyPath: "/projectId" },
  scopeSummaries: { id: "scopeSummaries", partitionKeyPath: "/projectId" },
  graphSubscriptions: { id: "graphSubscriptions", partitionKeyPath: "/projectId" },
  processedEvents: {
    id: "processedEvents",
    partitionKeyPath: "/projectId",
    defaultTtl: PROCESSED_EVENT_TTL_SECONDS,
  },
  conversationRefs: { id: "conversationRefs", partitionKeyPath: "/id" },
} satisfies Record<string, ContainerSpec>;

let client: CosmosClient | null = null;
let database: Database | null = null;
const containerCache = new Map<string, Container>();

function getClient(): CosmosClient {
  if (!client) {
    const endpoint = process.env.COSMOS_ENDPOINT;
    const key = process.env.COSMOS_KEY;
    if (!endpoint || !key) throw new Error("COSMOS_ENDPOINT or COSMOS_KEY is not set");
    client = new CosmosClient({ endpoint, key });
  }
  return client;
}

async function getContainer(spec: ContainerSpec): Promise<Container> {
  const cached = containerCache.get(spec.id);
  if (cached) return cached;

  if (!database) {
    ({ database } = await getClient().databases.createIfNotExists({ id: DB_NAME }));
  }
  const { container } = await database.containers.createIfNotExists({
    id: spec.id,
    partitionKey: { paths: [spec.partitionKeyPath] },
    ...(spec.defaultTtl ? { defaultTtl: spec.defaultTtl } : {}),
  });
  containerCache.set(spec.id, container);
  return container;
}

/** Cosmos document ids must not contain '/', '\', '?' or '#'. */
function sanitizeId(raw: string): string {
  return raw.replace(/[/\\?#\s]/g, "-");
}

// ── Violations ───────────────────────────────────────────────────────────────

export async function upsertViolation(violation: ScopeViolation): Promise<void> {
  const container = await getContainer(CONTAINERS.violations);
  await container.items.upsert({ ...violation, id: violation.violationId });
}

export async function getViolation(violationId: string, projectId: string): Promise<ScopeViolation | null> {
  const container = await getContainer(CONTAINERS.violations);
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
  const container = await getContainer(CONTAINERS.violations);
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
  const container = await getContainer(CONTAINERS.violations);
  await container.items.upsert({ ...existing, ...patch, id: violationId });
}

// ── Task Events ───────────────────────────────────────────────────────────────

export async function insertTaskEvent(event: TaskEvent): Promise<void> {
  const container = await getContainer(CONTAINERS.taskEvents);
  await container.items.create({ ...event, id: event.eventId });
}

export async function getRecentTaskEvents(projectId: string, hours = 24): Promise<TaskEvent[]> {
  const container = await getContainer(CONTAINERS.taskEvents);
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
  const container = await getContainer(CONTAINERS.feedbackLogs);
  await container.items.create({ ...log, id: log.feedbackId });
}

// ── Scope Summaries ───────────────────────────────────────────────────────────

export async function upsertScopeSummary(summary: ScopeSummary): Promise<void> {
  const container = await getContainer(CONTAINERS.scopeSummaries);
  await container.items.upsert({ ...summary, id: summary.projectId });
}

export async function getScopeSummary(projectId: string): Promise<ScopeSummary | null> {
  const container = await getContainer(CONTAINERS.scopeSummaries);
  try {
    const { resource } = await container.item(projectId, projectId).read<ScopeSummary>();
    return resource ?? null;
  } catch {
    return null;
  }
}

// ── Graph Subscription Records ───────────────────────────────────────────────

export interface SubscriptionRecord {
  id: string; // == projectId
  projectId: string;
  subscriptionId: string;
  siteId: string;
  listId: string;
  expiresAt: string;
  notificationUrl: string;
  /** Graph delta link — consumed and refreshed by the analysis worker. */
  deltaLink?: string;
}

export async function upsertSubscriptionRecord(record: SubscriptionRecord): Promise<void> {
  const container = await getContainer(CONTAINERS.graphSubscriptions);
  await container.items.upsert({ ...record });
}

export async function listSubscriptionRecords(): Promise<SubscriptionRecord[]> {
  const container = await getContainer(CONTAINERS.graphSubscriptions);
  const { resources } = await container.items
    .query<SubscriptionRecord>("SELECT * FROM c")
    .fetchAll();
  return resources;
}

export async function getSubscriptionRecord(projectId: string): Promise<SubscriptionRecord | null> {
  const container = await getContainer(CONTAINERS.graphSubscriptions);
  try {
    const { resource } = await container.item(projectId, projectId).read<SubscriptionRecord>();
    return resource ?? null;
  } catch {
    return null;
  }
}

export async function getSubscriptionBySubscriptionId(
  subscriptionId: string
): Promise<SubscriptionRecord | null> {
  const container = await getContainer(CONTAINERS.graphSubscriptions);
  const { resources } = await container.items
    .query<SubscriptionRecord>({
      query: "SELECT * FROM c WHERE c.subscriptionId = @sid",
      parameters: [{ name: "@sid", value: subscriptionId }],
    })
    .fetchAll();
  return resources[0] ?? null;
}

// ── Processed-Event Dedup ────────────────────────────────────────────────────

/**
 * Idempotency guard: atomically marks an item version (id + eTag) as processed.
 * Returns false when this exact version was already handled (Graph redeliveries,
 * overlapping delta pages). Records expire via container TTL.
 */
export async function tryMarkEventProcessed(
  projectId: string,
  itemId: string,
  version: string
): Promise<boolean> {
  const container = await getContainer(CONTAINERS.processedEvents);
  const id = sanitizeId(`${itemId}_${version}`);
  try {
    await container.items.create({ id, projectId });
    return true;
  } catch (err) {
    if ((err as { code?: number }).code === 409) return false;
    throw err;
  }
}

/** Rolls back a dedup marker so a failed analysis can be retried. */
export async function unmarkEventProcessed(
  projectId: string,
  itemId: string,
  version: string
): Promise<void> {
  const container = await getContainer(CONTAINERS.processedEvents);
  const id = sanitizeId(`${itemId}_${version}`);
  try {
    await container.item(id, projectId).delete();
  } catch {
    // Best effort — TTL cleans up eventually
  }
}

// ── Teams Conversation References ────────────────────────────────────────────

/**
 * Stores the conversation reference captured from incoming bot activity, so
 * proactive alerts use the real regional serviceUrl instead of a guessed one.
 */
export async function saveConversationRef(
  channelId: string,
  reference: Record<string, unknown>
): Promise<void> {
  const container = await getContainer(CONTAINERS.conversationRefs);
  await container.items.upsert({ id: sanitizeId(channelId), reference });
}

export async function getConversationRef(channelId: string): Promise<Record<string, unknown> | null> {
  const container = await getContainer(CONTAINERS.conversationRefs);
  const id = sanitizeId(channelId);
  try {
    const { resource } = await container.item(id, id).read<{ reference: Record<string, unknown> }>();
    return resource?.reference ?? null;
  } catch {
    return null;
  }
}
