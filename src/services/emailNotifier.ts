import { InvocationContext } from "@azure/functions";
import { Client } from "@microsoft/microsoft-graph-client";
import { ScopeViolation } from "../models/ProjectScope";
import { getGraphClient } from "./graphService";

const SEVERITY_COLORS: Record<ScopeViolation["severity"], string> = {
  low: "#2563eb",
  medium: "#f59e0b",
  high: "#ea580c",
  critical: "#dc2626",
};

/** Builds the HTML body for a violation alert email. Exported for tests. */
export function buildViolationEmailHtml(violation: ScopeViolation, projectName: string): string {
  const color = SEVERITY_COLORS[violation.severity];
  const esc = (s: string) =>
    s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  return `<!DOCTYPE html>
<html><body style="font-family: Segoe UI, Arial, sans-serif; color: #1f2937; max-width: 640px;">
  <div style="border-left: 6px solid ${color}; padding: 12px 16px; background: #f9fafb;">
    <h2 style="margin: 0 0 4px 0;">Scope creep detected — <span style="color: ${color};">${violation.severity.toUpperCase()}</span></h2>
    <p style="margin: 0; color: #6b7280;">Project: <strong>${esc(projectName)}</strong></p>
  </div>
  <table style="border-collapse: collapse; margin-top: 16px; width: 100%;">
    <tr><td style="padding: 6px 12px 6px 0; color: #6b7280; vertical-align: top;">Task</td>
        <td style="padding: 6px 0;"><strong>${esc(violation.taskTitle)}</strong></td></tr>
    <tr><td style="padding: 6px 12px 6px 0; color: #6b7280; vertical-align: top;">Detected</td>
        <td style="padding: 6px 0;">${new Date(violation.detectedAt).toLocaleString("da-DK", { timeZone: "Europe/Copenhagen" })}</td></tr>
    <tr><td style="padding: 6px 12px 6px 0; color: #6b7280; vertical-align: top;">Violation ID</td>
        <td style="padding: 6px 0; font-family: monospace; font-size: 12px;">${violation.violationId}</td></tr>
  </table>
  <h3 style="margin: 20px 0 6px 0;">Analysis</h3>
  <p style="margin: 0;">${esc(violation.reasoning)}</p>
  <h3 style="margin: 20px 0 6px 0;">Suggested action</h3>
  <p style="margin: 0;">${esc(violation.suggestedAction)}</p>
  <p style="margin-top: 24px; color: #9ca3af; font-size: 12px;">
    Sent by ScopeGuardian. Manage this violation via the Teams bot (<code>/violations ${esc(violation.projectId)}</code>) when enabled.
  </p>
</body></html>`;
}

/**
 * Sends a violation alert email via Graph sendMail (application permission
 * Mail.Send). Sender mailbox is ALERT_EMAIL_FROM; recipients ALERT_EMAIL_TO
 * (comma-separated). Both must be set for email alerts to be active.
 */
export async function sendViolationEmail(
  violation: ScopeViolation,
  projectName: string,
  context: InvocationContext
): Promise<void> {
  const from = process.env.ALERT_EMAIL_FROM ?? "";
  const recipients = (process.env.ALERT_EMAIL_TO ?? "")
    .split(",")
    .map((r) => r.trim())
    .filter(Boolean);

  if (!from || recipients.length === 0) {
    context.log("ALERT_EMAIL_FROM/ALERT_EMAIL_TO not set — skipping email alert");
    return;
  }

  const client: Client = getGraphClient();
  const message = {
    message: {
      subject: `[ScopeGuardian] ${violation.severity.toUpperCase()}: ${violation.taskTitle} (${projectName})`,
      body: { contentType: "HTML", content: buildViolationEmailHtml(violation, projectName) },
      toRecipients: recipients.map((address) => ({ emailAddress: { address } })),
    },
    saveToSentItems: false,
  };

  try {
    await client.api(`/users/${encodeURIComponent(from)}/sendMail`).post(message);
    context.log(`Violation email sent to ${recipients.length} recipient(s)`);
  } catch (err) {
    context.error("Failed to send violation email:", err);
    // Non-fatal — violation is still persisted in Cosmos DB
  }
}
