import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential, DefaultAzureCredential, TokenCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { TaskEvent } from "../models/TaskEvent";
import { v4 as uuidv4 } from "uuid";

const GRAPH_SCOPES = ["https://graph.microsoft.com/.default"];

// If created/modified are within this window, treat the change as a creation.
const CREATED_EVENT_WINDOW_MS = 2 * 60 * 1000;

let graphClient: Client | null = null;

function getCredential(): TokenCredential {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  if (tenantId && clientId && clientSecret) {
    return new ClientSecretCredential(tenantId, clientId, clientSecret);
  }
  // Falls back to Managed Identity / az login when no client secret is configured
  return new DefaultAzureCredential();
}

function getGraphClient(): Client {
  if (!graphClient) {
    const authProvider = new TokenCredentialAuthenticationProvider(getCredential(), {
      scopes: GRAPH_SCOPES,
    });
    graphClient = Client.initWithMiddleware({ authProvider });
  }
  return graphClient;
}

// ── Subscription Management ───────────────────────────────────────────────────

/**
 * Creates a Graph change subscription on a SharePoint list.
 * Note: SharePoint list subscriptions only support changeType "updated" and
 * must target the list resource itself (not /items). Notifications carry no
 * resourceData — changed items are discovered via delta queries.
 */
export async function createSharePointSubscription(
  siteId: string,
  listId: string,
  notificationUrl: string
): Promise<string> {
  const client = getGraphClient();
  const expiresAt = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();
  const subscription = (await client.api("/subscriptions").post({
    changeType: "updated",
    notificationUrl,
    resource: `/sites/${siteId}/lists/${listId}`,
    expirationDateTime: expiresAt,
    clientState: process.env.WEBHOOK_CLIENT_STATE ?? "scopeguardian-secret",
  })) as { id: string };
  return subscription.id;
}

export async function renewSubscription(subscriptionId: string): Promise<void> {
  const client = getGraphClient();
  const expiresAt = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();
  await client.api(`/subscriptions/${subscriptionId}`).patch({ expirationDateTime: expiresAt });
}

export async function deleteSubscription(subscriptionId: string): Promise<void> {
  const client = getGraphClient();
  await client.api(`/subscriptions/${subscriptionId}`).delete();
}

// ── Delta Queries ─────────────────────────────────────────────────────────────

export interface SharePointListItem {
  id: string;
  eTag?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  deleted?: { state: string };
  fields?: Record<string, unknown>;
}

interface DeltaPage {
  value?: SharePointListItem[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

export interface DeltaResult {
  items: SharePointListItem[];
  deltaLink: string | null;
}

/**
 * Fetches all changed list items since the given deltaLink (or the full list
 * when no link is provided), following nextLinks until exhausted.
 */
export async function fetchListItemDelta(
  siteId: string,
  listId: string,
  deltaLink?: string
): Promise<DeltaResult> {
  const client = getGraphClient();
  let url = deltaLink ?? `/sites/${siteId}/lists/${listId}/items/delta?$expand=fields`;
  const items: SharePointListItem[] = [];

  for (;;) {
    const page = (await client.api(url).get()) as DeltaPage;
    items.push(...(page.value ?? []));
    if (page["@odata.nextLink"]) {
      url = page["@odata.nextLink"];
      continue;
    }
    return { items, deltaLink: page["@odata.deltaLink"] ?? null };
  }
}

/**
 * Gets an up-to-date deltaLink without enumerating existing items
 * (token=latest), so a fresh subscription only reacts to future changes.
 */
export async function primeListItemDelta(siteId: string, listId: string): Promise<string | null> {
  const client = getGraphClient();
  const page = (await client
    .api(`/sites/${siteId}/lists/${listId}/items/delta?token=latest`)
    .get()) as DeltaPage;
  return page["@odata.deltaLink"] ?? null;
}

// ── List Item → TaskEvent Mapping ────────────────────────────────────────────

export interface TaskFieldMap {
  title: string;
  description: string;
  assignedTo: string;
  estimatedHours: string;
  loggedHours: string;
  tags: string;
}

const DEFAULT_FIELD_MAP: TaskFieldMap = {
  title: "Title",
  description: "Description",
  assignedTo: "AssignedTo",
  estimatedHours: "EstimatedHours",
  loggedHours: "LoggedHours",
  tags: "Tags",
};

/**
 * Lists differ per customer (e.g. ProjektPoint uses Remarks/Employee/Hours).
 * TASK_FIELD_MAP overrides individual internal field names as JSON, e.g.
 *   {"description":"Remarks","assignedTo":"Employee","estimatedHours":"Hours"}
 */
function getFieldMap(): TaskFieldMap {
  const raw = process.env.TASK_FIELD_MAP;
  if (!raw) return DEFAULT_FIELD_MAP;
  try {
    return { ...DEFAULT_FIELD_MAP, ...(JSON.parse(raw) as Partial<TaskFieldMap>) };
  } catch {
    return DEFAULT_FIELD_MAP;
  }
}

/** Person/lookup columns arrive as objects or as <Field>LookupId — normalise to a string. */
function fieldToString(value: unknown): string | undefined {
  if (value == null) return undefined;
  if (typeof value === "object") {
    const o = value as Record<string, unknown>;
    const name = o.displayName ?? o.LookupValue ?? o.Title ?? o.email;
    return name != null ? String(name) : undefined;
  }
  return String(value);
}

/**
 * Maps a SharePoint list item from a delta result to a TaskEvent.
 * Returns null for deleted items — deletions are not scope-creep candidates.
 */
export function mapListItemToTaskEvent(item: SharePointListItem, projectId: string): TaskEvent | null {
  if (item.deleted) return null;

  const map = getFieldMap();
  const fields = item.fields ?? {};
  const created = item.createdDateTime ? Date.parse(item.createdDateTime) : 0;
  const modified = item.lastModifiedDateTime ? Date.parse(item.lastModifiedDateTime) : created;
  const eventType = created && modified - created < CREATED_EVENT_WINDOW_MS ? "task_created" : "task_updated";

  const assignedTo =
    fieldToString(fields[map.assignedTo]) ?? fieldToString(fields[`${map.assignedTo}LookupId`]);
  const estimatedHours = fields[map.estimatedHours];
  const loggedHours = fields[map.loggedHours];
  const tags = fieldToString(fields[map.tags]);

  return {
    eventId: uuidv4(),
    projectId,
    eventType,
    occurredAt: new Date().toISOString(),
    source: "sharepoint",
    task: {
      id: item.id,
      title: fieldToString(fields[map.title]) ?? "Untitled",
      description: fieldToString(fields[map.description]) ?? "",
      assignedTo,
      estimatedHours: estimatedHours != null ? Number(estimatedHours) : undefined,
      loggedHours: loggedHours != null ? Number(loggedHours) : undefined,
      tags: tags ? tags.split(";").map((t) => t.trim()).filter(Boolean) : [],
    },
    rawPayload: item as unknown as Record<string, unknown>,
  };
}
