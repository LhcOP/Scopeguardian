export interface ScopeItem {
  id: string;
  title: string;
  description: string;
  deliverables: string[];
  acceptanceCriteria: string[];
  estimatedHours?: number;
  tags: string[];
}

export interface ProjectScope {
  projectId: string;
  projectName: string;
  version: string;
  createdAt: string;
  updatedAt: string;
  scopeItems: ScopeItem[];
  outOfScope: string[];
  assumptions: string[];
  constraints: string[];
  embeddingVector?: number[];
}

export interface ScopeViolation {
  violationId: string;
  projectId: string;
  detectedAt: string;
  severity: "low" | "medium" | "high" | "critical";
  taskEventId: string;
  taskTitle: string;
  taskDescription: string;
  reasoning: string;
  affectedScopeItems: string[];
  suggestedAction: string;
  status: "pending" | "acknowledged" | "escalated" | "dismissed";
  acknowledgedBy?: string;
  acknowledgedAt?: string;
}

export interface ScopeSummary {
  projectId: string;
  projectName: string;
  totalItems: number;
  lastAnalyzedAt: string;
  violationCount: number;
  riskScore: number;
}
