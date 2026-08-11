import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
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

export function getGraphClient(): Client {
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

/**
 * Creates a Graph change subscription on a document library (drive).
 * Same constraints as lists: changeType "updated" only, no resourceData —
 * changed files are discovered via drive delta queries.
 */
export async function createDriveSubscription(
  driveId: string,
  notificationUrl: string
): Promise<string> {
  const client = getGraphClient();
  const expiresAt = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString();
  const subscription = (await client.api("/subscriptions").post({
    changeType: "updated",
    notificationUrl,
    resource: `/drives/${driveId}/root`,
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

// ── Drive Delta (document library monitoring) ────────────────────────────────

export interface DriveDeltaItem {
  id: string;
  name?: string;
  eTag?: string;
  size?: number;
  lastModifiedDateTime?: string;
  deleted?: { state: string };
  file?: { mimeType?: string };
  folder?: unknown;
  parentReference?: { driveId?: string };
}

interface DriveDeltaPage {
  value?: DriveDeltaItem[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

/** Returns the id of the site's default document library drive. */
export async function getSiteDefaultDriveId(siteId: string): Promise<string> {
  const client = getGraphClient();
  const drive = (await client.api(`/sites/${siteId}/drive`).select("id").get()) as { id: string };
  return drive.id;
}

/** Fetches changed drive items since the delta link (files and folders). */
export async function fetchDriveDelta(
  driveId: string,
  deltaLink?: string
): Promise<{ items: DriveDeltaItem[]; deltaLink: string | null }> {
  const client = getGraphClient();
  let url = deltaLink ?? `/drives/${driveId}/root/delta`;
  const items: DriveDeltaItem[] = [];

  for (;;) {
    const page = (await client.api(url).get()) as DriveDeltaPage;
    items.push(...(page.value ?? []));
    if (page["@odata.nextLink"]) {
      url = page["@odata.nextLink"];
      continue;
    }
    return { items, deltaLink: page["@odata.deltaLink"] ?? null };
  }
}

export async function primeDriveDelta(driveId: string): Promise<string | null> {
  const client = getGraphClient();
  const page = (await client.api(`/drives/${driveId}/root/delta?token=latest`).get()) as DriveDeltaPage;
  return page["@odata.deltaLink"] ?? null;
}

// ── Generic List/Drive Access (scope bootstrapping) ──────────────────────────

const MAX_LIST_PAGES = 10;

/** Fetches items (fields only) from a list addressed by display name or id, following paging. */
export async function getListItemFields(
  siteId: string,
  listNameOrId: string,
  filter?: string,
  allowNonIndexedFilter = false
): Promise<Record<string, unknown>[]> {
  const client = getGraphClient();
  let request = client
    .api(`/sites/${siteId}/lists/${encodeURIComponent(listNameOrId)}/items`)
    .expand("fields")
    .top(200);
  if (filter) request = request.filter(filter);
  if (allowNonIndexedFilter) {
    request = request.header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
  }

  type ItemsPage = { value?: { fields?: Record<string, unknown> }[]; "@odata.nextLink"?: string };
  const results: Record<string, unknown>[] = [];
  let page = (await request.get()) as ItemsPage;
  for (let i = 0; i < MAX_LIST_PAGES; i++) {
    results.push(...(page.value ?? []).map((item) => item.fields ?? {}));
    const next = page["@odata.nextLink"];
    if (!next) break;
    page = (await client.api(next).get()) as ItemsPage;
  }
  return results;
}

/** Lists all lists on a site (id + display name). */
export async function getSiteLists(siteId: string): Promise<{ id: string; displayName: string }[]> {
  const client = getGraphClient();
  const page = (await client.api(`/sites/${siteId}/lists`).select("id,displayName").top(200).get()) as {
    value?: { id: string; displayName: string }[];
  };
  return page.value ?? [];
}

export interface DriveFileHit {
  driveId: string;
  itemId: string;
  name: string;
  size: number;
}

/** Searches all document libraries on a site for files matching the query. */
export async function searchProjectFiles(siteId: string, query: string): Promise<DriveFileHit[]> {
  const client = getGraphClient();
  const drives = (await client.api(`/sites/${siteId}/drives`).get()) as {
    value?: { id: string }[];
  };

  const hits: DriveFileHit[] = [];
  for (const drive of drives.value ?? []) {
    try {
      const results = (await client
        .api(`/drives/${drive.id}/root/search(q='${query.replace(/'/g, "''")}')`)
        .top(10)
        .get()) as { value?: { id: string; name: string; size?: number; file?: unknown }[] };
      for (const item of results.value ?? []) {
        if (item.file) {
          hits.push({ driveId: drive.id, itemId: item.id, name: item.name, size: item.size ?? 0 });
        }
      }
    } catch {
      // Drive not searchable — skip
    }
  }
  return hits;
}

export async function downloadDriveItemContent(driveId: string, itemId: string): Promise<Buffer> {
  const client = getGraphClient();
  const data = (await client
    .api(`/drives/${driveId}/items/${itemId}/content`)
    .responseType(ResponseType.ARRAYBUFFER)
    .get()) as ArrayBuffer;
  return Buffer.from(data);
}

// ── List Item → TaskEvent Mapping ────────────────────────────────────────────

export interface TaskFieldMap {
  title: string;
  description: string;
  assignedTo: string;
  estimatedHours: string;
  loggedHours: string;
  deadline: string;
  status: string;
  tags: string;
}

const DEFAULT_FIELD_MAP: TaskFieldMap = {
  title: "Title",
  description: "Description",
  assignedTo: "AssignedTo",
  estimatedHours: "EstimatedHours",
  loggedHours: "LoggedHours",
  deadline: "Deadline",
  status: "Status",
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
      deadline: fieldToString(fields[map.deadline]),
      status: fieldToString(fields[map.status]),
      tags: tags ? tags.split(";").map((t) => t.trim()).filter(Boolean) : [],
    },
    rawPayload: item as unknown as Record<string, unknown>,
  };
}
