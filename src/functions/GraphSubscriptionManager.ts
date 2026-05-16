import { app, Timer, InvocationContext } from "@azure/functions";
import { CosmosClient } from "@azure/cosmos";
import { createSharePointSubscription, renewSubscription } from "../services/graphService";

const DB_NAME = "scopeguardian";
const CONTAINER_ID = "graphSubscriptions";

interface SubscriptionRecord {
  id: string;
  projectId: string;
  subscriptionId: string;
  siteId: string;
  listId: string;
  expiresAt: string;
  notificationUrl: string;
}

function getCosmosContainer() {
  const endpoint = process.env.COSMOS_ENDPOINT;
  const key = process.env.COSMOS_KEY;
  if (!endpoint || !key) throw new Error("COSMOS_ENDPOINT or COSMOS_KEY is not set");
  const client = new CosmosClient({ endpoint, key });
  return client.database(DB_NAME).container(CONTAINER_ID);
}

async function listAllSubscriptions(): Promise<SubscriptionRecord[]> {
  const container = getCosmosContainer();
  const { resources } = await container.items
    .query<SubscriptionRecord>("SELECT * FROM c")
    .fetchAll();
  return resources;
}

async function upsertSubscriptionRecord(record: SubscriptionRecord): Promise<void> {
  const container = getCosmosContainer();
  await container.items.upsert(record);
}

/**
 * Timer-triggered function — runs every 12 hours.
 * Renews any Graph subscriptions expiring within the next 24 hours,
 * and bootstraps missing subscriptions for registered projects.
 */
async function graphSubscriptionManagerHandler(_timer: Timer, context: InvocationContext): Promise<void> {
  const notificationUrl = process.env.WEBHOOK_NOTIFICATION_URL;
  if (!notificationUrl) {
    context.warn("WEBHOOK_NOTIFICATION_URL not set — skipping subscription management");
    return;
  }

  const existing = await listAllSubscriptions();
  const renewalCutoff = new Date(Date.now() + 24 * 3_600_000).toISOString();

  // Renew subscriptions expiring within 24 hours
  for (const record of existing) {
    if (record.expiresAt < renewalCutoff) {
      try {
        await renewSubscription(record.subscriptionId);
        const newExpiry = new Date(Date.now() + 3 * 24 * 3_600_000).toISOString();
        await upsertSubscriptionRecord({ ...record, expiresAt: newExpiry });
        context.log(`Renewed subscription ${record.subscriptionId} for project ${record.projectId}`);
      } catch (err) {
        context.error(`Failed to renew subscription ${record.subscriptionId}:`, err);

        // Subscription may have been deleted externally — re-create it
        try {
          const newId = await createSharePointSubscription(record.siteId, record.listId, notificationUrl);
          const newExpiry = new Date(Date.now() + 3 * 24 * 3_600_000).toISOString();
          await upsertSubscriptionRecord({ ...record, subscriptionId: newId, expiresAt: newExpiry });
          context.log(`Re-created subscription for project ${record.projectId} as ${newId}`);
        } catch (createErr) {
          context.error(`Failed to re-create subscription for project ${record.projectId}:`, createErr);
        }
      }
    }
  }

  // Bootstrap subscriptions for projects not yet registered
  const projectConfigs = parseProjectConfigs();
  const registeredProjects = new Set(existing.map((r) => r.projectId));

  for (const config of projectConfigs) {
    if (registeredProjects.has(config.projectId)) continue;
    try {
      const subscriptionId = await createSharePointSubscription(
        config.siteId,
        config.listId,
        notificationUrl
      );
      const record: SubscriptionRecord = {
        id: config.projectId,
        projectId: config.projectId,
        subscriptionId,
        siteId: config.siteId,
        listId: config.listId,
        expiresAt: new Date(Date.now() + 3 * 24 * 3_600_000).toISOString(),
        notificationUrl,
      };
      await upsertSubscriptionRecord(record);
      context.log(`Created new subscription for project ${config.projectId}: ${subscriptionId}`);
    } catch (err) {
      context.error(`Failed to create subscription for project ${config.projectId}:`, err);
    }
  }
}

interface ProjectConfig {
  projectId: string;
  siteId: string;
  listId: string;
}

/**
 * Parses PROJECT_CONFIGS env var.
 * Format: "projectId:siteId:listId,projectId2:siteId2:listId2"
 */
function parseProjectConfigs(): ProjectConfig[] {
  const raw = process.env.PROJECT_CONFIGS ?? "";
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .map((entry) => {
      const [projectId, siteId, listId] = entry.split(":");
      return { projectId, siteId, listId };
    })
    .filter((c) => c.projectId && c.siteId && c.listId);
}

app.timer("GraphSubscriptionManager", {
  schedule: "0 0 */12 * * *", // Every 12 hours
  runOnStartup: true,
  handler: graphSubscriptionManagerHandler,
});
