import { app, Timer, InvocationContext } from "@azure/functions";
import {
  listSubscriptionRecords,
  upsertSubscriptionRecord,
  docsRecordId,
  SubscriptionRecord,
} from "../services/cosmosDbService";
import {
  createSharePointSubscription,
  createDriveSubscription,
  renewSubscription,
  primeListItemDelta,
  primeDriveDelta,
  getSiteDefaultDriveId,
} from "../services/graphService";
import { parseProjectConfigs } from "../utils/projectConfig";
import { bootstrapMasterScope } from "../services/scopeBootstrapper";
import { downloadMasterScope } from "../services/blobStorageService";

const SUBSCRIPTION_LIFETIME_MS = 3 * 24 * 3_600_000;
const RENEWAL_WINDOW_MS = 24 * 3_600_000;

/**
 * Timer-triggered function — runs every 12 hours.
 * Renews Graph subscriptions expiring within the next 24 hours and bootstraps
 * missing subscriptions for projects registered in PROJECT_CONFIGS. New
 * subscriptions get their delta link primed so only future changes are analysed.
 */
async function graphSubscriptionManagerHandler(_timer: Timer, context: InvocationContext): Promise<void> {
  const notificationUrl = process.env.WEBHOOK_NOTIFICATION_URL;
  if (!notificationUrl) {
    context.warn("WEBHOOK_NOTIFICATION_URL not set — skipping subscription management");
    return;
  }

  const existing = await listSubscriptionRecords();
  const renewalCutoff = new Date(Date.now() + RENEWAL_WINDOW_MS).toISOString();

  // Renew subscriptions expiring within 24 hours
  for (const record of existing) {
    if (record.expiresAt >= renewalCutoff) continue;
    try {
      await renewSubscription(record.subscriptionId);
      const newExpiry = new Date(Date.now() + SUBSCRIPTION_LIFETIME_MS).toISOString();
      await upsertSubscriptionRecord({ ...record, expiresAt: newExpiry });
      context.log(`Renewed subscription ${record.subscriptionId} for project ${record.projectId}`);
    } catch (err) {
      context.error(`Failed to renew subscription ${record.subscriptionId}:`, err);

      // Subscription may have been deleted externally — re-create it
      try {
        const newId =
          record.resourceType === "drive" && record.driveId
            ? await createDriveSubscription(record.driveId, notificationUrl)
            : await createSharePointSubscription(record.siteId, record.listId, notificationUrl);
        const newExpiry = new Date(Date.now() + SUBSCRIPTION_LIFETIME_MS).toISOString();
        await upsertSubscriptionRecord({ ...record, subscriptionId: newId, expiresAt: newExpiry });
        context.log(`Re-created subscription for project ${record.projectId} as ${newId}`);
      } catch (createErr) {
        context.error(`Failed to re-create subscription for project ${record.projectId}:`, createErr);
      }
    }
  }

  // Bootstrap subscriptions for projects not yet registered
  const projectConfigs = parseProjectConfigs();
  const registeredIds = new Set(existing.map((r) => r.id));

  for (const config of projectConfigs) {
    // Document-library subscription (monitors new/changed documents)
    if (!registeredIds.has(docsRecordId(config.projectId))) {
      try {
        const driveId = await getSiteDefaultDriveId(config.siteId);
        const subscriptionId = await createDriveSubscription(driveId, notificationUrl);
        const deltaLink = (await primeDriveDelta(driveId)) ?? undefined;
        await upsertSubscriptionRecord({
          id: docsRecordId(config.projectId),
          projectId: config.projectId,
          subscriptionId,
          siteId: config.siteId,
          listId: config.listId,
          expiresAt: new Date(Date.now() + SUBSCRIPTION_LIFETIME_MS).toISOString(),
          notificationUrl,
          deltaLink,
          resourceType: "drive",
          driveId,
        });
        context.log(`Created document subscription for project ${config.projectId}: ${subscriptionId}`);
      } catch (err) {
        context.error(`Failed to create document subscription for project ${config.projectId}:`, err);
      }
    }

    if (registeredIds.has(config.projectId)) continue;
    try {
      const subscriptionId = await createSharePointSubscription(
        config.siteId,
        config.listId,
        notificationUrl
      );
      // Prime delta so the first notification only sees changes made after now
      const deltaLink = (await primeListItemDelta(config.siteId, config.listId)) ?? undefined;
      const record: SubscriptionRecord = {
        id: config.projectId,
        projectId: config.projectId,
        subscriptionId,
        siteId: config.siteId,
        listId: config.listId,
        expiresAt: new Date(Date.now() + SUBSCRIPTION_LIFETIME_MS).toISOString(),
        notificationUrl,
        deltaLink,
        resourceType: "list",
      };
      await upsertSubscriptionRecord(record);
      context.log(`Created new subscription for project ${config.projectId}: ${subscriptionId}`);

      // Fully automatic onboarding: generate master scope from project material
      if (!(await downloadMasterScope(config.projectId))) {
        try {
          await bootstrapMasterScope(config.projectId, config.siteId, context);
        } catch (scopeErr) {
          context.error(`Scope bootstrap failed for project ${config.projectId}:`, scopeErr);
        }
      }
    } catch (err) {
      context.error(`Failed to create subscription for project ${config.projectId}:`, err);
    }
  }
}

app.timer("GraphSubscriptionManager", {
  schedule: "0 0 */12 * * *", // Every 12 hours
  runOnStartup: true,
  handler: graphSubscriptionManagerHandler,
});
