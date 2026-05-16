export type FeedbackAction =
  | "scope_violation_acknowledged"
  | "scope_violation_dismissed"
  | "scope_violation_escalated"
  | "scope_updated"
  | "false_positive_reported";

export interface FeedbackLog {
  feedbackId: string;
  projectId: string;
  violationId: string;
  action: FeedbackAction;
  actorId: string;
  actorName: string;
  actorEmail: string;
  comment?: string;
  occurredAt: string;
  teamsConversationId?: string;
  teamsActivityId?: string;
}

export interface TeamsCardAction {
  type: "Action.Submit";
  title: string;
  data: {
    action: FeedbackAction;
    violationId: string;
    projectId: string;
    comment?: string;
  };
}

export interface BotActivityPayload {
  type: string;
  id: string;
  timestamp: string;
  channelId: string;
  from: {
    id: string;
    name: string;
    aadObjectId?: string;
  };
  conversation: {
    id: string;
    tenantId?: string;
  };
  value?: Record<string, unknown>;
  text?: string;
}
