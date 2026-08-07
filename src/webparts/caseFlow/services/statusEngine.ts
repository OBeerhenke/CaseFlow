/**
 * Pure status-evaluation engine for CaseFlow.
 *
 * Zero dependencies (no SharePoint, no React). Everything is a
 * deterministic function of its arguments — ideal for unit testing and
 * safe to use from anywhere in the app.
 *
 * The engine consumes *configurable* status keys (passed in via
 * IStatusEngineConfig) so it does NOT hardcode German status strings.
 * Production callers typically pass `StatusConfigService`-derived values;
 * tests pass whatever they like.
 */

import { CONFIG_KEYS } from '../models/config';

/**
 * Configuration the engine needs to evaluate a single case's status.
 * All values are strings to match the ConfigService key-value storage;
 * the engine parses what it needs.
 */
export interface IStatusEngineConfig {
  /** Number of workdays before the planned date when the status switches to "review". */
  reviewDays: number;
  /** The status key to emit when no planned date has been set. */
  statusNoPlan: string;
  /** The status key for cases whose planned date is past today. */
  statusOverdue: string;
  /** The status key for cases approaching the deadline (within reviewDays workdays). */
  statusReview: string;
  /** The status key for cases on track (planned date well in the future). */
  statusOnTrack: string;
  /** The status key for completed cases — the engine always skips these. */
  statusCompleted: string;
}

/** Default config matching the original hardcoded Barlog behavior. */
export const DEFAULT_STATUS_ENGINE_CONFIG: IStatusEngineConfig = {
  reviewDays: 3,
  statusNoPlan: 'Termin planen',
  statusOverdue: 'überfällig',
  statusReview: 'prüfen',
  statusOnTrack: 'läuft planmäßig',
  statusCompleted: 'abgeschlossen'
};

/**
 * Build an IStatusEngineConfig from a ConfigService-like key-value map.
 * This is the wire-in point that keeps SharePoint out of the engine.
 */
export function configFromMap(map: Record<string, string>): IStatusEngineConfig {
  return {
    reviewDays: map[CONFIG_KEYS.REVIEW_DAYS] ? parseInt(map[CONFIG_KEYS.REVIEW_DAYS], 10) || 3 : 3,
    statusNoPlan: 'Termin planen',
    statusOverdue: 'überfällig',
    statusReview: 'prüfen',
    statusOnTrack: 'läuft planmäßig',
    statusCompleted: 'abgeschlossen'
  };
}

/**
 * Parse `plannedDate` (ISO "YYYY-MM-DD" or German "DD.MM.YYYY") into a
 * Date at midnight local time. Returns `undefined` for falsy / unparseable
 * input.
 */
export function parsePlannedDate(plannedDate: string | undefined | null): Date | undefined {
  if (!plannedDate) return undefined;
  if (plannedDate.includes('-')) {
    const d = new Date(plannedDate);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
    return undefined;
  }
  const parts = plannedDate.split('.');
  if (parts.length === 3) {
    const d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }
  return undefined;
}

/**
 * Add N workdays (Mon–Fri) to a given date. Replicates Excel's WORKDAY()
 * function. Pure and deterministic.
 */
export function addWorkdays(start: Date, days: number): Date {
  const result = new Date(start.getTime());
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const dow = result.getDay();
    if (dow !== 0 && dow !== 6) {
      added++;
    }
  }
  return result;
}

export interface IEvaluationResult {
  /** The computed status key. */
  status: string;
  /** Whether the status differs from the caller-supplied current status. */
  changed: boolean;
}

/**
 * Evaluate the status for a single case.
 *
 * @param plannedDate  The value of the planned-date field (field_6), or undefined/null.
 * @param currentStatus  The case's current status key, or undefined.
 * @param config  Engine configuration (thresholds, status key labels).
 * @param today  The date to compare against — injectable for deterministic tests.
 * @returns  The computed status + a `changed` flag.
 */
export function evaluateCaseStatus(
  plannedDate: string | undefined | null,
  currentStatus: string | undefined,
  config: IStatusEngineConfig = DEFAULT_STATUS_ENGINE_CONFIG,
  today: Date = new Date()
): IEvaluationResult {
  // Completed cases are never re-evaluated.
  if (currentStatus === config.statusCompleted) {
    return { status: config.statusCompleted, changed: false };
  }

  const parsed = parsePlannedDate(plannedDate);

  if (!parsed) {
    // No planned date → statusNoPlan (unless it already is that).
    return {
      status: config.statusNoPlan,
      changed: currentStatus !== config.statusNoPlan
    };
  }

  // Clone to avoid mutating the caller's Date object.
  const todayNorm = new Date(today);
  todayNorm.setHours(0, 0, 0, 0);

  if (parsed <= todayNorm) {
    return {
      status: config.statusOverdue,
      changed: currentStatus !== config.statusOverdue
    };
  }

  const reviewDeadline = addWorkdays(todayNorm, config.reviewDays);

  if (parsed <= reviewDeadline) {
    return {
      status: config.statusReview,
      changed: currentStatus !== config.statusReview
    };
  }

  return {
    status: config.statusOnTrack,
    changed: currentStatus !== config.statusOnTrack
  };
}
