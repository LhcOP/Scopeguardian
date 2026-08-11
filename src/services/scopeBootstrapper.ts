import { InvocationContext } from "@azure/functions";
import { ProjectScope } from "../models/ProjectScope";
import {
  getListItemFields,
  searchProjectFiles,
  downloadDriveItemContent,
} from "./graphService";
import { generateScopeFromMaterial, embedScopeItems } from "./openaiService";
import { uploadMasterScope } from "./blobStorageService";
import { ensureIndexExists, upsertScopeItems, deleteScopeItemsByProject } from "./aiSearchService";
import { upsertScopeSummary } from "./cosmosDbService";
import { extractTextFromFile, isSupportedDocument } from "../utils/documentText";

const MAX_DOCS = 3;
const MAX_DOC_CHARS = 30_000;
const MAX_FILE_BYTES = 15 * 1024 * 1024;

/** Stringifies primitive field values; objects (lookups etc.) are ignored. */
function fieldText(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return null;
}

/**
 * Fully automatic master-scope generation — no consultant involvement.
 * Gathers the project's own material and lets the LLM derive the scope:
 *   1. Project description from the hub Projects register (PROJECTS_HUB_SITE_ID,
 *      matched on ProjectNumber == projectId)
 *   2. Quote lines from the project site's QuoteLines list (when present)
 *   3. Estimation-basis documents on the project site (drive search on
 *      SCOPE_DOC_SEARCH, default "estimat"; docx/xlsx/pdf/txt/md)
 * Publishes and indexes the generated scope. Returns null when the material
 * is insufficient — the project is then skipped until material appears.
 */
export async function bootstrapMasterScope(
  projectId: string,
  siteId: string,
  context: InvocationContext
): Promise<ProjectScope | null> {
  const sections: string[] = [];
  let projectName = projectId;

  // 1. Hub project register
  const hubSiteId = process.env.PROJECTS_HUB_SITE_ID;
  if (hubSiteId) {
    try {
      const hubList = process.env.PROJECTS_HUB_LIST ?? "Projects";
      const records = await getListItemFields(
        hubSiteId,
        hubList,
        `fields/ProjectNumber eq '${projectId.replace(/'/g, "''")}'`
      );
      const record = records[0];
      if (record) {
        projectName = fieldText(record.Title) ?? projectId;
        const description = fieldText(record.Description);
        const projectType = fieldText(record.ProjectType);
        const totalEstimate = fieldText(record.TotalEstimate);
        const parts = [
          `Project: ${projectName}`,
          projectType ? `Type: ${projectType}` : "",
          description ? `Description:\n${description}` : "",
          totalEstimate ? `Total estimated hours: ${totalEstimate}` : "",
        ].filter(Boolean);
        if (parts.length > 1) {
          sections.push(`=== PROJECT REGISTER ===\n${parts.join("\n")}`);
        }
        context.log(`Scope bootstrap: found hub record for ${projectId} ("${projectName}")`);
      } else {
        context.log(`Scope bootstrap: no hub record for ${projectId}`);
      }
    } catch (err) {
      context.warn(`Scope bootstrap: hub lookup failed for ${projectId}:`, err);
    }
  }

  // 2. Quote lines on the project site
  try {
    const quoteLines = await getListItemFields(siteId, "QuoteLines");
    if (quoteLines.length > 0) {
      const lines = quoteLines
        .map((q) => {
          const fields = ["Title", "Description", "Hours", "Quantity", "LineTotal"]
            .map((f) => {
              const value = fieldText(q[f]);
              return value ? `${f}: ${value}` : "";
            })
            .filter(Boolean)
            .join(", ");
          return `- ${fields}`;
        })
        .join("\n");
      sections.push(`=== QUOTE LINES ===\n${lines}`);
      context.log(`Scope bootstrap: ${quoteLines.length} quote line(s) for ${projectId}`);
    }
  } catch {
    // No QuoteLines list on this site — fine
  }

  // 3. Estimation-basis documents
  const searchTerm = process.env.SCOPE_DOC_SEARCH ?? "estimat";
  try {
    const files = (await searchProjectFiles(siteId, searchTerm))
      .filter((f) => isSupportedDocument(f.name) && f.size <= MAX_FILE_BYTES)
      .slice(0, MAX_DOCS);
    for (const file of files) {
      const buffer = await downloadDriveItemContent(file.driveId, file.itemId);
      const text = await extractTextFromFile(file.name, buffer);
      if (text?.trim()) {
        sections.push(`=== DOCUMENT: ${file.name} ===\n${text.slice(0, MAX_DOC_CHARS)}`);
        context.log(`Scope bootstrap: extracted ${text.length} chars from ${file.name}`);
      }
    }
  } catch (err) {
    context.warn(`Scope bootstrap: document search failed for ${projectId}:`, err);
  }

  if (sections.length === 0) {
    context.warn(`Scope bootstrap: no material found for ${projectId} — cannot generate scope`);
    return null;
  }

  const scope = await generateScopeFromMaterial(projectId, projectName, sections.join("\n\n"));
  if (!scope) {
    context.warn(`Scope bootstrap: material for ${projectId} was insufficient for scope generation`);
    return null;
  }

  await uploadMasterScope(scope);
  await ensureIndexExists();
  await deleteScopeItemsByProject(projectId);
  const embeddings = await embedScopeItems(scope.scopeItems);
  await upsertScopeItems(projectId, scope.scopeItems, embeddings);
  await upsertScopeSummary({
    projectId,
    projectName: scope.projectName,
    totalItems: scope.scopeItems.length,
    lastAnalyzedAt: new Date().toISOString(),
    violationCount: 0,
    riskScore: 0,
  });

  context.log(
    `Scope bootstrap: generated and indexed master scope for ${projectId} (${scope.scopeItems.length} items, ${scope.outOfScope.length} exclusions)`
  );
  return scope;
}
