export interface ProjectConfig {
  projectId: string;
  siteId: string;
  listId: string;
}

/**
 * Parses the PROJECT_CONFIGS env var — the single source of truth for
 * which SharePoint lists are monitored for which projects.
 *
 * Canonical format (Graph site ids contain commas, so entries are separated
 * by ";" and fields by "|"):
 *   "projectId|siteId|listId;projectId2|siteId2|listId2"
 *
 * Legacy format ("projectId:siteId:listId,...") is still accepted for
 * entries without "|".
 */
export function parseProjectConfigs(raw: string = process.env.PROJECT_CONFIGS ?? ""): ProjectConfig[] {
  const entries = raw.includes("|") ? raw.split(";") : raw.split(",");
  return entries
    .map((s) => s.trim())
    .filter(Boolean)
    .map((entry) => {
      const [projectId, siteId, listId] = entry.includes("|") ? entry.split("|") : entry.split(":");
      return { projectId: projectId?.trim(), siteId: siteId?.trim(), listId: listId?.trim() };
    })
    .filter((c) => c.projectId && c.siteId && c.listId);
}
