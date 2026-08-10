export interface ProjectConfig {
  projectId: string;
  siteId: string;
  listId: string;
}

/**
 * Parses the PROJECT_CONFIGS env var — the single source of truth for
 * which SharePoint lists are monitored for which projects.
 * Format: "projectId:siteId:listId,projectId2:siteId2:listId2"
 */
export function parseProjectConfigs(raw: string = process.env.PROJECT_CONFIGS ?? ""): ProjectConfig[] {
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
