import { InvocationContext } from "@azure/functions";
import { getSiteLists, getListItemFields } from "./graphService";

export interface RegisteredHours {
  /** Hours registered on the specific task (matched on task name). */
  taskHours: number;
  /** Total hours registered on the project across all tasks. */
  projectHours: number;
  /** Per-task totals keyed by normalised task name. */
  byTask: Map<string, number>;
}

function normalizeTaskName(name: string): string {
  return name.trim().toLowerCase();
}

/** Pure aggregation over time-registration rows — exported for tests. */
export function aggregateRegisteredHours(
  rows: Record<string, unknown>[],
  taskName?: string
): RegisteredHours {
  const byTask = new Map<string, number>();
  let projectHours = 0;
  for (const row of rows) {
    const hours = Number(row.RegHours);
    if (!isFinite(hours) || hours <= 0) continue;
    projectHours += hours;
    const rowTask = typeof row.RegTaskName === "string" ? normalizeTaskName(row.RegTaskName) : "";
    if (rowTask) byTask.set(rowTask, (byTask.get(rowTask) ?? 0) + hours);
  }
  const taskHours = taskName ? byTask.get(normalizeTaskName(taskName)) ?? 0 : 0;
  return { taskHours, projectHours, byTask };
}

/**
 * Sums registered hours for a project from the central time-tracking hub
 * (TIMETRACK_SITE_ID) — one "Tidsreg - <employee>" list per employee,
 * rows matched on RegProjectNumber. Returns null when the hub is not
 * configured; empty totals when no registrations exist.
 */
export async function getRegisteredHours(
  projectNumber: string,
  taskName: string | undefined,
  context: InvocationContext
): Promise<RegisteredHours | null> {
  const siteId = process.env.TIMETRACK_SITE_ID;
  if (!siteId) return null;
  const prefix = (process.env.TIMETRACK_LIST_PREFIX ?? "Tidsreg").toLowerCase();

  const lists = (await getSiteLists(siteId)).filter((l) =>
    l.displayName.toLowerCase().startsWith(prefix)
  );

  const allRows: Record<string, unknown>[] = [];
  const escaped = projectNumber.replace(/'/g, "''");
  for (const list of lists) {
    try {
      const rows = await getListItemFields(
        siteId,
        list.id,
        `fields/RegProjectNumber eq '${escaped}'`,
        true
      );
      allRows.push(...rows);
    } catch {
      // Non-indexed filter can fail on large lists — fall back to client-side filtering
      try {
        const rows = await getListItemFields(siteId, list.id);
        allRows.push(...rows.filter((r) => String(r.RegProjectNumber ?? "") === projectNumber));
      } catch (err) {
        context.warn(`Time tracking: could not read list "${list.displayName}":`, err);
      }
    }
  }

  return aggregateRegisteredHours(allRows, taskName);
}
