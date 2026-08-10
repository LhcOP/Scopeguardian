import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { ProjectScope, ScopeItem } from "../models/ProjectScope";
import {
  uploadMasterScope,
  downloadMasterScope,
  downloadScopeMarkdown,
} from "../services/blobStorageService";
import { embedScopeItems } from "../services/openaiService";
import { ensureIndexExists, upsertScopeItems, deleteScopeItemsByProject } from "../services/aiSearchService";
import { upsertScopeSummary, listViolations } from "../services/cosmosDbService";
import { computeRiskScore } from "../utils/riskScore";
import { compareVersions } from "../utils/versionCompare";

/**
 * HTTP trigger — scope onboarding and retrieval.
 *   PUT/POST /api/projects/{projectId}/scope  — upload master scope + index immediately
 *   GET      /api/projects/{projectId}/scope  — fetch latest master scope JSON
 */
async function scopeManagementHandler(
  req: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const projectId = req.params.projectId;
  if (!projectId) return { status: 400, jsonBody: { error: "projectId route parameter is required" } };

  if (req.method === "GET") {
    const scope = await downloadMasterScope(projectId);
    if (!scope) return { status: 404, jsonBody: { error: `No scope found for project ${projectId}` } };
    return { status: 200, jsonBody: scope };
  }

  // PUT/POST — upload new scope version
  let body: Partial<ProjectScope>;
  try {
    body = (await req.json()) as Partial<ProjectScope>;
  } catch {
    return { status: 400, jsonBody: { error: "Invalid JSON payload" } };
  }

  const validationError = validateScopePayload(body);
  if (validationError) return { status: 400, jsonBody: { error: validationError } };

  const existing = await downloadMasterScope(projectId);
  const version = body.version?.trim() || nextVersion(existing?.version);
  if (existing && compareVersions(version, existing.version) <= 0) {
    return {
      status: 409,
      jsonBody: { error: `Version ${version} must be greater than existing version ${existing.version}` },
    };
  }

  const now = new Date().toISOString();
  const scope: ProjectScope = {
    projectId,
    projectName: body.projectName!,
    version,
    createdAt: existing?.createdAt ?? now,
    updatedAt: now,
    scopeItems: body.scopeItems!.map(normalizeScopeItem),
    outOfScope: body.outOfScope ?? [],
    assumptions: body.assumptions ?? [],
    constraints: body.constraints ?? [],
  };

  await uploadMasterScope(scope);
  context.log(`Master scope v${version} uploaded for project ${projectId}`);

  // Index immediately so violation detection works without waiting for the sync timer
  await ensureIndexExists();
  await deleteScopeItemsByProject(projectId);
  const embeddings = await embedScopeItems(scope.scopeItems);
  await upsertScopeItems(projectId, scope.scopeItems, embeddings);
  context.log(`Indexed ${scope.scopeItems.length} scope items for ${projectId}`);

  const pending = await listViolations(projectId, "pending");
  await upsertScopeSummary({
    projectId,
    projectName: scope.projectName,
    totalItems: scope.scopeItems.length,
    lastAnalyzedAt: now,
    violationCount: pending.length,
    riskScore: computeRiskScore(pending),
  });

  return {
    status: 200,
    jsonBody: { projectId, version, itemsIndexed: scope.scopeItems.length },
  };
}

/** GET /api/projects/{projectId}/scope/summary — Markdown summary for stakeholders. */
async function scopeSummaryHandler(
  req: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  const projectId = req.params.projectId;
  if (!projectId) return { status: 400, jsonBody: { error: "projectId route parameter is required" } };

  const markdown = await downloadScopeMarkdown(projectId);
  if (!markdown) return { status: 404, jsonBody: { error: `No scope summary found for project ${projectId}` } };
  return { status: 200, headers: { "Content-Type": "text/markdown" }, body: markdown };
}

function validateScopePayload(body: Partial<ProjectScope>): string | null {
  if (!body.projectName?.trim()) return "projectName is required";
  if (!Array.isArray(body.scopeItems) || body.scopeItems.length === 0) {
    return "scopeItems must be a non-empty array";
  }
  for (const [i, item] of body.scopeItems.entries()) {
    if (!item.id?.trim()) return `scopeItems[${i}].id is required`;
    if (!item.title?.trim()) return `scopeItems[${i}].title is required`;
    if (!item.description?.trim()) return `scopeItems[${i}].description is required`;
  }
  const ids = new Set(body.scopeItems.map((i) => i.id));
  if (ids.size !== body.scopeItems.length) return "scopeItems ids must be unique";
  return null;
}

function normalizeScopeItem(item: ScopeItem): ScopeItem {
  return {
    id: item.id,
    title: item.title,
    description: item.description,
    deliverables: item.deliverables ?? [],
    acceptanceCriteria: item.acceptanceCriteria ?? [],
    estimatedHours: item.estimatedHours,
    tags: item.tags ?? [],
  };
}

function nextVersion(existing?: string): string {
  if (!existing) return "1";
  const major = parseInt(existing.split(".")[0], 10) || 0;
  return String(major + 1);
}

app.http("ScopeManagement", {
  methods: ["GET", "PUT", "POST"],
  authLevel: "function",
  route: "projects/{projectId}/scope",
  handler: scopeManagementHandler,
});

app.http("ScopeSummary", {
  methods: ["GET"],
  authLevel: "function",
  route: "projects/{projectId}/scope/summary",
  handler: scopeSummaryHandler,
});
