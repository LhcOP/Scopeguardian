import {
  buildViolationAlertCard,
  buildConfirmationCard,
  buildScopeStatusCard,
  cardToJson,
} from "../src/utils/adaptiveCardBuilder";
import { ScopeViolation } from "../src/models/ProjectScope";

const sampleViolation: ScopeViolation = {
  violationId: "v-001",
  projectId: "proj-001",
  detectedAt: new Date().toISOString(),
  severity: "high",
  taskEventId: "evt-001",
  taskTitle: "Add payment gateway integration",
  taskDescription: "Integrate Stripe for subscription billing",
  reasoning: "Payment processing is explicitly out of scope for Phase 1.",
  affectedScopeItems: ["scope-item-1"],
  suggestedAction: "Defer to Phase 2 backlog or raise a change request.",
  status: "pending",
};

describe("buildViolationAlertCard", () => {
  it("returns a valid AdaptiveCard structure", () => {
    const card = buildViolationAlertCard(sampleViolation, "Project Alpha");
    expect(card.type).toBe("AdaptiveCard");
    expect(card.version).toBe("1.4");
    expect(Array.isArray(card.body)).toBe(true);
    expect(Array.isArray(card.actions)).toBe(true);
  });

  it("includes 3 action buttons", () => {
    const card = buildViolationAlertCard(sampleViolation, "Project Alpha");
    expect(card.actions).toHaveLength(3);
  });

  it("embeds violationId and projectId in action data", () => {
    const card = buildViolationAlertCard(sampleViolation, "Project Alpha");
    const acknowledgeAction = card.actions![0] as { data: { violationId: string; projectId: string } };
    expect(acknowledgeAction.data.violationId).toBe("v-001");
    expect(acknowledgeAction.data.projectId).toBe("proj-001");
  });
});

describe("buildConfirmationCard", () => {
  it("shows success indicator on success", () => {
    const card = buildConfirmationCard("Done!", true);
    const textBlock = (card.body[0] as { text: string }).text;
    expect(textBlock).toContain("✅");
  });

  it("shows failure indicator on failure", () => {
    const card = buildConfirmationCard("Error!", false);
    const textBlock = (card.body[0] as { text: string }).text;
    expect(textBlock).toContain("❌");
  });
});

describe("buildScopeStatusCard", () => {
  it("uses Attention color for high risk score", () => {
    const card = buildScopeStatusCard("Alpha", 85, 5, new Date().toISOString());
    const json = cardToJson(card);
    expect(json).toContain("Attention");
  });

  it("uses Good color for low risk score", () => {
    const card = buildScopeStatusCard("Alpha", 10, 0, new Date().toISOString());
    const json = cardToJson(card);
    expect(json).toContain("Good");
  });
});

describe("cardToJson", () => {
  it("produces valid JSON string", () => {
    const card = buildConfirmationCard("Test", true);
    const json = cardToJson(card);
    expect(() => JSON.parse(json)).not.toThrow();
  });
});
