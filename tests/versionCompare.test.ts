import { compareVersions } from "../src/utils/versionCompare";

describe("compareVersions", () => {
  it("compares simple versions", () => {
    expect(compareVersions("2", "1")).toBeGreaterThan(0);
    expect(compareVersions("1", "2")).toBeLessThan(0);
    expect(compareVersions("3", "3")).toBe(0);
  });

  it("compares numerically per segment (1.10 > 1.9)", () => {
    expect(compareVersions("1.10", "1.9")).toBeGreaterThan(0);
    expect(compareVersions("1.9", "1.10")).toBeLessThan(0);
  });

  it("handles differing segment counts", () => {
    expect(compareVersions("1.0.1", "1")).toBeGreaterThan(0);
    expect(compareVersions("1.0.0", "1")).toBe(0);
    expect(compareVersions("2", "1.9.9")).toBeGreaterThan(0);
  });

  it("handles semver-style versions", () => {
    expect(compareVersions("2.1.3", "2.1.2")).toBeGreaterThan(0);
    expect(compareVersions("10.0.0", "9.9.9")).toBeGreaterThan(0);
  });
});
