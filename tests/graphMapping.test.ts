import { mapListItemToTaskEvent, SharePointListItem } from "../src/services/graphService";

describe("mapListItemToTaskEvent", () => {
  const baseItem: SharePointListItem = {
    id: "42",
    eTag: '"abc,1"',
    createdDateTime: "2026-08-01T10:00:00Z",
    lastModifiedDateTime: "2026-08-05T10:00:00Z",
    fields: {
      Title: "Build export module",
      Description: "CSV export for reports",
      AssignedTo: "Jane Doe",
      EstimatedHours: 12,
      Tags: "backend; export",
    },
  };

  it("maps fields to a TaskEvent", () => {
    const event = mapListItemToTaskEvent(baseItem, "proj-1");
    expect(event).not.toBeNull();
    expect(event!.projectId).toBe("proj-1");
    expect(event!.source).toBe("sharepoint");
    expect(event!.task.id).toBe("42");
    expect(event!.task.title).toBe("Build export module");
    expect(event!.task.description).toBe("CSV export for reports");
    expect(event!.task.estimatedHours).toBe(12);
    expect(event!.task.tags).toEqual(["backend", "export"]);
  });

  it("classifies as task_updated when modified long after creation", () => {
    const event = mapListItemToTaskEvent(baseItem, "proj-1");
    expect(event!.eventType).toBe("task_updated");
  });

  it("classifies as task_created when modified right after creation", () => {
    const item = {
      ...baseItem,
      createdDateTime: "2026-08-05T10:00:00Z",
      lastModifiedDateTime: "2026-08-05T10:00:30Z",
    };
    const event = mapListItemToTaskEvent(item, "proj-1");
    expect(event!.eventType).toBe("task_created");
  });

  it("returns null for deleted items", () => {
    const item = { ...baseItem, deleted: { state: "deleted" } };
    expect(mapListItemToTaskEvent(item, "proj-1")).toBeNull();
  });

  it("defaults missing fields gracefully", () => {
    const item: SharePointListItem = { id: "7", fields: {} };
    const event = mapListItemToTaskEvent(item, "proj-1");
    expect(event!.task.title).toBe("Untitled");
    expect(event!.task.description).toBe("");
    expect(event!.task.tags).toEqual([]);
    expect(event!.task.estimatedHours).toBeUndefined();
  });

  it("handles items without fields object", () => {
    const item: SharePointListItem = { id: "8" };
    const event = mapListItemToTaskEvent(item, "proj-1");
    expect(event).not.toBeNull();
    expect(event!.task.title).toBe("Untitled");
  });

  it("respects TASK_FIELD_MAP overrides (ProjektPoint-style lists)", () => {
    process.env.TASK_FIELD_MAP = JSON.stringify({
      description: "Remarks",
      assignedTo: "Employee",
      estimatedHours: "Hours",
    });
    try {
      const item: SharePointListItem = {
        id: "9",
        fields: {
          Title: "Opsæt integration",
          Remarks: "Integration til TimeLog",
          Employee: { displayName: "Jane Doe", email: "jd@x.dk" },
          Hours: 8,
        },
      };
      const event = mapListItemToTaskEvent(item, "proj-1");
      expect(event!.task.description).toBe("Integration til TimeLog");
      expect(event!.task.assignedTo).toBe("Jane Doe");
      expect(event!.task.estimatedHours).toBe(8);
    } finally {
      delete process.env.TASK_FIELD_MAP;
    }
  });
});
