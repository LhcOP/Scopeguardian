import { parseViolationPayload } from "../src/services/openaiService";

const violation = {
  severity: "high",
  reasoning: "Estimate doubled",
  affectedScopeItems: ["S1"],
  suggestedAction: "Review",
};

describe("parseViolationPayload", () => {
  it("parses the canonical wrapper shape", () => {
    const result = parseViolationPayload(JSON.stringify({ violations: [violation] }));
    expect(result).toHaveLength(1);
    expect(result[0].severity).toBe("high");
  });

  it("parses a bare single violation object (no wrapper)", () => {
    const result = parseViolationPayload(JSON.stringify(violation));
    expect(result).toHaveLength(1);
    expect(result[0].reasoning).toBe("Estimate doubled");
  });

  it("parses a bare array", () => {
    const result = parseViolationPayload(JSON.stringify([violation, { ...violation, severity: "low" }]));
    expect(result).toHaveLength(2);
  });

  it("returns empty for empty wrapper", () => {
    expect(parseViolationPayload(JSON.stringify({ violations: [] }))).toHaveLength(0);
  });

  it("returns empty for non-violation objects", () => {
    expect(parseViolationPayload(JSON.stringify({ note: "all good" }))).toHaveLength(0);
  });

  it("filters entries with invalid severity", () => {
    const result = parseViolationPayload(
      JSON.stringify({ violations: [violation, { ...violation, severity: "extreme" }] })
    );
    expect(result).toHaveLength(1);
  });

  it("returns empty on malformed JSON", () => {
    expect(parseViolationPayload("not json{")).toHaveLength(0);
  });
});
