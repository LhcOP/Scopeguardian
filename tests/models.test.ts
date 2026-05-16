import { ProjectScope, ScopeViolation } from "../src/models/ProjectScope";
import { TaskEvent } from "../src/models/TaskEvent";
import { FeedbackLog } from "../src/models/FeedbackLog";

describe("Model shape validation", () => {
  it("ProjectScope has required fields", () => {
    const scope: ProjectScope = {
      projectId: "p1",
      projectName: "Test Project",
      version: "1.0",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      scopeItems: [],
      outOfScope: [],
      assumptions: [],
      constraints: [],
    };
    expect(scope.projectId).toBe("p1");
    expect(scope.scopeItems).toHaveLength(0);
  });

  it("ScopeViolation defaults to pending status", () => {
    const violation: ScopeViolation = {
      violationId: "v1",
      projectId: "p1",
      detectedAt: new Date().toISOString(),
      severity: "medium",
      taskEventId: "e1",
      taskTitle: "New task",
      taskDescription: "Description",
      reasoning: "Reason",
      affectedScopeItems: [],
      suggestedAction: "Action",
      status: "pending",
    };
    expect(violation.status).toBe("pending");
  });

  it("TaskEvent source is one of known sources", () => {
    const event: TaskEvent = {
      eventId: "e1",
      projectId: "p1",
      eventType: "task_created",
      occurredAt: new Date().toISOString(),
      source: "sharepoint",
      task: {
        id: "t1",
        title: "Task 1",
        description: "Do stuff",
        tags: [],
      },
    };
    expect(["sharepoint", "planner", "devops", "jira"]).toContain(event.source);
  });

  it("FeedbackLog captures actor info", () => {
    const log: FeedbackLog = {
      feedbackId: "f1",
      projectId: "p1",
      violationId: "v1",
      action: "scope_violation_acknowledged",
      actorId: "u1",
      actorName: "Alice",
      actorEmail: "alice@example.com",
      occurredAt: new Date().toISOString(),
    };
    expect(log.actorName).toBe("Alice");
    expect(log.action).toBe("scope_violation_acknowledged");
  });
});
