import { parseProjectConfigs } from "../src/utils/projectConfig";

describe("parseProjectConfigs", () => {
  it("parses the canonical pipe format (site ids contain commas)", () => {
    const result = parseProjectConfigs(
      "040560|itinventech.sharepoint.com,f8a65783-e7dc,3ac0f626-ccc7|1c7792dc-61d2"
    );
    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({
      projectId: "040560",
      siteId: "itinventech.sharepoint.com,f8a65783-e7dc,3ac0f626-ccc7",
      listId: "1c7792dc-61d2",
    });
  });

  it("parses multiple pipe-format entries separated by semicolon", () => {
    const result = parseProjectConfigs("p1|host,g1,g2|l1;p2|host,g3,g4|l2");
    expect(result).toHaveLength(2);
    expect(result[1].projectId).toBe("p2");
    expect(result[1].siteId).toBe("host,g3,g4");
  });

  it("parses a single legacy-format config", () => {
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

  it("reads PROJECT_CONFIGS env var by default", () => {
    process.env.PROJECT_CONFIGS = "env-proj:env-site:env-list";
    try {
      const result = parseProjectConfigs();
      expect(result).toEqual([{ projectId: "env-proj", siteId: "env-site", listId: "env-list" }]);
    } finally {
      delete process.env.PROJECT_CONFIGS;
    }
  });
});
