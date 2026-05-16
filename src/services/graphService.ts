import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { GraphChangeNotification, TaskEvent } from "../models/TaskEvent";
import { v4 as uuidv4 } from "uuid";

const GRAPH_SCOPES = ["https://graph.microsoft.com/.default"];

function getGraphClient(): Client {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("AZURE_TENANT_ID, AZURE_CLIENT_ID, or AZURE_CLIENT_SECRET is not set");
  }
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: GRAPH_SCOPES });
  return Client.initWithMiddleware({ authProvider });
}

// ── Subscription Management ───────────────────────────────────────────────────

export async function createSharePointSubscription(
  siteId: string,
  listId: string,
  notificationUrl: string
): Promise<string> {
  const client = getGraphClient();
  const expiresAt = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toISOString(); // 3 days max for SP lists
  const subscription = await client.api("/subscriptions").post({
    changeType: "created,updated,deleted",
    notificationUrl,
    resource: `/sites/${siteId}/lists/${listId}/items`,
    expirationDateTime: expiresAt,
    clientState: process.env.WEBHOOK_CLIENT_STATE ?? "scopeguardian-secret",
  });
  return subscription.id as string;
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

// ── SharePoint List Item Fetching ─────────────────────────────────────────────

export async function fetchSharePointListItem(
  siteId: string,
  listId: string,
  itemId: string
): Promise<Record<string, unknown>> {
  const client = getGraphClient();
  return client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`).get() as Promise<
    Record<string, unknown>
  >;
}

// ── Notification → TaskEvent Mapping ─────────────────────────────────────────

export async function mapNotificationToTaskEvent(
  notification: GraphChangeNotification,
  projectId: string
): Promise<TaskEvent | null> {
  if (!notification.resourceData) return null;

  // Extract siteId and listId from the resource path
  // resource format: "sites/{siteId}/lists/{listId}/items/{itemId}"
  const parts = notification.resource.split("/");
  const siteId = parts[1];
  const listId = parts[3];
  const itemId = notification.resourceData.id;

  try {
    const rawItem = await fetchSharePointListItem(siteId, listId, itemId);
    const fields = rawItem.fields as Record<string, unknown>;

    return {
      eventId: uuidv4(),
      projectId,
      eventType: mapChangeType(notification.changeType),
      occurredAt: new Date().toISOString(),
      source: "sharepoint",
      task: {
        id: itemId,
        title: String(fields["Title"] ?? "Untitled"),
        description: String(fields["Description"] ?? ""),
        assignedTo: fields["AssignedTo"] ? String(fields["AssignedTo"]) : undefined,
        estimatedHours: fields["EstimatedHours"] ? Number(fields["EstimatedHours"]) : undefined,
        loggedHours: fields["LoggedHours"] ? Number(fields["LoggedHours"]) : undefined,
        tags: fields["Tags"] ? String(fields["Tags"]).split(";").map((t) => t.trim()) : [],
      },
      rawPayload: rawItem,
    };
  } catch (err) {
    console.error(`Failed to fetch SharePoint item ${itemId}:`, err);
    return null;
  }
}

function mapChangeType(changeType: string): TaskEvent["eventType"] {
  switch (changeType) {
    case "created":
      return "task_created";
    case "deleted":
      return "task_completed";
    default:
      return "task_updated";
  }
}

// ── Teams Message Sending ─────────────────────────────────────────────────────

export async function sendTeamsMessage(
  channelId: string,
  teamId: string,
  content: string
): Promise<void> {
  const client = getGraphClient();
  await client.api(`/teams/${teamId}/channels/${channelId}/messages`).post({
    body: { contentType: "html", content },
  });
}
