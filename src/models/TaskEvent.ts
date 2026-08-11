export type TaskEventType =
  | "task_created"
  | "task_updated"
  | "task_assigned"
  | "task_completed"
  | "time_logged"
  | "comment_added"
  | "document_added";

export interface TaskEvent {
  eventId: string;
  projectId: string;
  eventType: TaskEventType;
  occurredAt: string;
  source: "sharepoint" | "planner" | "devops" | "jira";
  task: {
    id: string;
    title: string;
    description: string;
    assignedTo?: string;
    estimatedHours?: number;
    loggedHours?: number;
    deadline?: string;
    status?: string;
    tags?: string[];
    parentTaskId?: string;
    comments?: string[];
  };
  changeDetails?: {
    field: string;
    oldValue: string | null;
    newValue: string | null;
  }[];
  rawPayload?: Record<string, unknown>;
}

export interface GraphWebhookNotification {
  value: GraphChangeNotification[];
}

/** Message enqueued by EventTrigger and consumed by ProcessScopeAnalysis. */
export interface AnalysisQueueMessage {
  projectId: string;
  siteId: string;
  listId: string;
  subscriptionId: string;
  notifiedAt: string;
  /** "list" (task list, default) or "drive" (document library). */
  resource?: "list" | "drive";
  driveId?: string;
}

export interface GraphChangeNotification {
  subscriptionId: string;
  subscriptionExpirationDateTime: string;
  changeType: string;
  resource: string;
  resourceData?: {
    "@odata.type": string;
    "@odata.id": string;
    "@odata.etag"?: string;
    id: string;
  };
  clientState?: string;
  tenantId: string;
}
