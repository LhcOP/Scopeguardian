import { ScopeViolation } from "../models/ProjectScope";

interface AdaptiveCard {
  type: "AdaptiveCard";
  version: string;
  body: unknown[];
  actions?: unknown[];
}

const SEVERITY_COLORS: Record<ScopeViolation["severity"], string> = {
  low: "Good",
  medium: "Warning",
  high: "Attention",
  critical: "Attention",
};

const SEVERITY_ICONS: Record<ScopeViolation["severity"], string> = {
  low: "🔵",
  medium: "🟡",
  high: "🟠",
  critical: "🔴",
};

export function buildViolationAlertCard(violation: ScopeViolation, projectName: string): AdaptiveCard {
  const icon = SEVERITY_ICONS[violation.severity];
  const color = SEVERITY_COLORS[violation.severity];

  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "Container",
        style: "emphasis",
        items: [
          {
            type: "TextBlock",
            text: `${icon} Scope Creep Detected — ${violation.severity.toUpperCase()}`,
            weight: "Bolder",
            size: "Large",
            color,
          },
          {
            type: "TextBlock",
            text: `Project: **${projectName}**`,
            spacing: "None",
            wrap: true,
          },
        ],
      },
      {
        type: "FactSet",
        facts: [
          { title: "Task", value: violation.taskTitle },
          { title: "Detected At", value: new Date(violation.detectedAt).toLocaleString() },
          { title: "Violation ID", value: violation.violationId },
        ],
      },
      {
        type: "Container",
        items: [
          {
            type: "TextBlock",
            text: "**Analysis**",
            weight: "Bolder",
          },
          {
            type: "TextBlock",
            text: violation.reasoning,
            wrap: true,
          },
          {
            type: "TextBlock",
            text: "**Suggested Action**",
            weight: "Bolder",
            spacing: "Small",
          },
          {
            type: "TextBlock",
            text: violation.suggestedAction,
            wrap: true,
          },
        ],
      },
      {
        type: "Input.Text",
        id: "comment",
        placeholder: "Optional comment (e.g., justification for dismissal)...",
        isMultiline: true,
        maxLength: 500,
      },
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "✅ Acknowledge",
        style: "positive",
        data: {
          action: "scope_violation_acknowledged",
          violationId: violation.violationId,
          projectId: violation.projectId,
        },
      },
      {
        type: "Action.Submit",
        title: "❌ Dismiss (False Positive)",
        data: {
          action: "false_positive_reported",
          violationId: violation.violationId,
          projectId: violation.projectId,
        },
      },
      {
        type: "Action.Submit",
        title: "⬆️ Escalate",
        style: "destructive",
        data: {
          action: "scope_violation_escalated",
          violationId: violation.violationId,
          projectId: violation.projectId,
        },
      },
    ],
  };
}

export function buildConfirmationCard(message: string, success: boolean): AdaptiveCard {
  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: success ? `✅ ${message}` : `❌ ${message}`,
        weight: "Bolder",
        color: success ? "Good" : "Attention",
        wrap: true,
      },
    ],
  };
}

export function buildScopeStatusCard(
  projectName: string,
  riskScore: number,
  pendingViolations: number,
  lastAnalyzedAt: string
): AdaptiveCard {
  const riskColor = riskScore >= 70 ? "Attention" : riskScore >= 40 ? "Warning" : "Good";
  const riskLabel = riskScore >= 70 ? "HIGH" : riskScore >= 40 ? "MEDIUM" : "LOW";

  return {
    type: "AdaptiveCard",
    version: "1.4",
    body: [
      {
        type: "TextBlock",
        text: `📊 Scope Status: ${projectName}`,
        weight: "Bolder",
        size: "Large",
      },
      {
        type: "ColumnSet",
        columns: [
          {
            type: "Column",
            width: "stretch",
            items: [
              { type: "TextBlock", text: "Risk Score", weight: "Bolder" },
              { type: "TextBlock", text: `${riskScore}/100 (${riskLabel})`, color: riskColor, size: "ExtraLarge" },
            ],
          },
          {
            type: "Column",
            width: "stretch",
            items: [
              { type: "TextBlock", text: "Pending Violations", weight: "Bolder" },
              {
                type: "TextBlock",
                text: String(pendingViolations),
                color: pendingViolations > 0 ? "Attention" : "Good",
                size: "ExtraLarge",
              },
            ],
          },
        ],
      },
      {
        type: "TextBlock",
        text: `Last analysed: ${new Date(lastAnalyzedAt).toLocaleString()}`,
        isSubtle: true,
        size: "Small",
      },
    ],
  };
}

export function cardToJson(card: AdaptiveCard): string {
  return JSON.stringify(card);
}
