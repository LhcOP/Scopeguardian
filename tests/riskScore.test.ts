import { computeRiskScore } from "../src/utils/riskScore";

describe("computeRiskScore", () => {
  it("returns 0 for no violations", () => {
    expect(computeRiskScore([])).toBe(0);
  });

  it("weights severities differently", () => {
    expect(computeRiskScore([{ severity: "low" }])).toBe(5);
    expect(computeRiskScore([{ severity: "medium" }])).toBe(10);
    expect(computeRiskScore([{ severity: "high" }])).toBe(20);
    expect(computeRiskScore([{ severity: "critical" }])).toBe(30);
  });

  it("sums multiple violations", () => {
    expect(
      computeRiskScore([{ severity: "low" }, { severity: "high" }, { severity: "critical" }])
    ).toBe(55);
  });

  it("caps at 100", () => {
    const many = Array.from({ length: 10 }, () => ({ severity: "critical" as const }));
    expect(computeRiskScore(many)).toBe(100);
  });

  it("treats unknown severity as medium", () => {
    expect(computeRiskScore([{ severity: "weird" as never }])).toBe(10);
  });
});
