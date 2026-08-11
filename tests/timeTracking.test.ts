import { aggregateRegisteredHours } from "../src/services/timeTrackingService";

describe("aggregateRegisteredHours", () => {
  const rows = [
    { RegProjectNumber: "040560", RegTaskName: "1.02 Create item list", RegHours: 4 },
    { RegProjectNumber: "040560", RegTaskName: "1.02 Create item list", RegHours: 3.5 },
    { RegProjectNumber: "040560", RegTaskName: "1.10 Check machines", RegHours: 2 },
    { RegProjectNumber: "040560", RegTaskName: "  1.02 CREATE ITEM LIST ", RegHours: 1 },
  ];

  it("sums project hours across all rows", () => {
    const result = aggregateRegisteredHours(rows);
    expect(result.projectHours).toBe(10.5);
  });

  it("matches task hours case-insensitively with trimming", () => {
    const result = aggregateRegisteredHours(rows, "1.02 Create item list");
    expect(result.taskHours).toBe(8.5);
  });

  it("returns zero task hours for unknown task", () => {
    const result = aggregateRegisteredHours(rows, "9.99 Unknown");
    expect(result.taskHours).toBe(0);
  });

  it("groups totals by normalised task name", () => {
    const result = aggregateRegisteredHours(rows);
    expect(result.byTask.get("1.02 create item list")).toBe(8.5);
    expect(result.byTask.get("1.10 check machines")).toBe(2);
  });

  it("ignores invalid and non-positive hours", () => {
    const dirty = [
      { RegTaskName: "A", RegHours: "abc" },
      { RegTaskName: "A", RegHours: -5 },
      { RegTaskName: "A", RegHours: 0 },
      { RegTaskName: "A", RegHours: 2 },
    ];
    const result = aggregateRegisteredHours(dirty, "A");
    expect(result.taskHours).toBe(2);
    expect(result.projectHours).toBe(2);
  });

  it("handles rows without task name", () => {
    const result = aggregateRegisteredHours([{ RegHours: 5 }]);
    expect(result.projectHours).toBe(5);
    expect(result.byTask.size).toBe(0);
  });
});
