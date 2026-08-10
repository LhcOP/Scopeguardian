import { buildViolationEmailHtml } from "../src/services/emailNotifier";
import { ScopeViolation } from "../src/models/ProjectScope";

describe("buildViolationEmailHtml", () => {
  const violation: ScopeViolation = {
    violationId: "v-123",
    projectId: "040560",
    detectedAt: "2026-08-10T12:00:00Z",
    severity: "high",
    taskEventId: "e-1",
    taskTitle: "Purchase 3 new sewing machines",
    taskDescription: "Buy machines",
    reasoning: "Machinery purchase is explicitly out of scope",
    affectedScopeItems: ["S2"],
    suggestedAction: "Halt procurement & escalate",
    status: "pending",
  };

  it("includes severity, task, reasoning and action", () => {
    const html = buildViolationEmailHtml(violation, "Test Project");
    expect(html).toContain("HIGH");
    expect(html).toContain("Purchase 3 new sewing machines");
    expect(html).toContain("Machinery purchase is explicitly out of scope");
    expect(html).toContain("Halt procurement &amp; escalate");
    expect(html).toContain("Test Project");
    expect(html).toContain("v-123");
  });

  it("escapes HTML in user-controlled fields", () => {
    const nasty = { ...violation, taskTitle: '<script>alert("x")</script>' };
    const html = buildViolationEmailHtml(nasty, "P");
    expect(html).not.toContain("<script>");
    expect(html).toContain("&lt;script&gt;");
  });

  it("uses severity-specific color", () => {
    expect(buildViolationEmailHtml(violation, "P")).toContain("#ea580c");
    expect(buildViolationEmailHtml({ ...violation, severity: "critical" }, "P")).toContain("#dc2626");
  });
});
