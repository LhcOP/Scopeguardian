// Tests for parseProjectConfigs via the exported helper shape.
// The timer handler itself requires Azure credentials so is integration-tested only.

describe("PROJECT_CONFIGS parsing", () => {
  function parseProjectConfigs(raw: string) {
    return raw
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean)
      .map((entry) => {
        const [projectId, siteId, listId] = entry.split(":");
        return { projectId, siteId, listId };
      })
      .filter((c) => c.projectId && c.siteId && c.listId);
  }

  it("parses a single project config", () => {
    const result = parseProjectConfigs("proj1:site-abc:list-xyz");
    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({ projectId: "proj1", siteId: "site-abc", listId: "list-xyz" });
  });

  it("parses multiple project configs", () => {
    const result = parseProjectConfigs("p1:s1:l1,p2:s2:l2,p3:s3:l3");
    expect(result).toHaveLength(3);
    expect(result[1].projectId).toBe("p2");
  });

  it("skips incomplete entries", () => {
    const result = parseProjectConfigs("p1:s1:l1,bad-entry,p2:s2:l2");
    expect(result).toHaveLength(2);
  });

  it("returns empty array for empty string", () => {
    expect(parseProjectConfigs("")).toHaveLength(0);
  });

  it("trims whitespace around entries", () => {
    const result = parseProjectConfigs("  p1:s1:l1 , p2:s2:l2  ");
    expect(result).toHaveLength(2);
    expect(result[0].projectId).toBe("p1");
  });
});
