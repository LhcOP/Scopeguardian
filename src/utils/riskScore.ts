import { ScopeViolation } from "../models/ProjectScope";

const SEVERITY_WEIGHTS: Record<ScopeViolation["severity"], number> = {
  low: 5,
  medium: 10,
  high: 20,
  critical: 30,
};

/**
 * Deterministic project risk score (0-100) computed from pending violations.
 * Single source of truth — used by both the event pipeline and the sync timer
 * so the two never disagree.
 */
export function computeRiskScore(pendingViolations: Pick<ScopeViolation, "severity">[]): number {
  const total = pendingViolations.reduce(
    (sum, v) => sum + (SEVERITY_WEIGHTS[v.severity] ?? SEVERITY_WEIGHTS.medium),
    0
  );
  return Math.min(100, total);
}
