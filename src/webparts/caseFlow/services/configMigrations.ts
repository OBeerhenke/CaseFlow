import { CONFIG_KEYS, CURRENT_CONFIG_SCHEMA_VERSION } from '../models/config';

export interface IConfigMigration {
  fromVersion: string;
  toVersion: string;
  /** Pure transform: old config shape -> new config shape. */
  migrate: (config: Record<string, string>) => Record<string, string>;
}

/**
 * Example migration: pre-CaseFlow (Barlog "TA-Management") config used the
 * key "ReviewDays" for the number of workdays before a case is flagged for
 * review. CaseFlow 1.1 renames it to the German "PruefenTage" to match the
 * rest of the config vocabulary (DelayThresholdDays, etc.).
 *
 * This illustrates the migration pattern: each entry moves a config from
 * exactly one version to the next; `migrateConfig` walks the chain.
 */
export const CONFIG_MIGRATIONS: IConfigMigration[] = [
  {
    fromVersion: '1.0',
    toVersion: '1.1',
    migrate: (config) => {
      const next = { ...config };
      if ('ReviewDays' in next && !(CONFIG_KEYS.REVIEW_DAYS in next)) {
        next[CONFIG_KEYS.REVIEW_DAYS] = next.ReviewDays;
      }
      delete next.ReviewDays;
      return next;
    }
  }
];

export interface IMigrationResult {
  config: Record<string, string>;
  version: string;
  migrated: boolean;
  appliedMigrations: string[];
}

/**
 * Walks `config` forward through `migrations` from its current
 * `ConfigSchemaVersion` (defaulting to "1.0" if absent — i.e. a config
 * written before schema versioning existed) up to `targetVersion`.
 *
 * Pure function — no SharePoint I/O. Safe to call repeatedly / on every
 * app load; a config already at `targetVersion` is returned unchanged.
 */
export function migrateConfig(
  config: Record<string, string>,
  migrations: IConfigMigration[] = CONFIG_MIGRATIONS,
  targetVersion: string = CURRENT_CONFIG_SCHEMA_VERSION
): IMigrationResult {
  let currentVersion = config[CONFIG_KEYS.SCHEMA_VERSION] || '1.0';
  let result = { ...config };
  const appliedMigrations: string[] = [];

  // Bound the loop by the number of available migrations to guard against
  // a misconfigured migration chain (e.g. a cycle) looping forever.
  let guard = 0;
  while (currentVersion !== targetVersion && guard <= migrations.length) {
    const migration = migrations.find(m => m.fromVersion === currentVersion);
    if (!migration) break; // no known path from here — stop, leave version as-is
    result = migration.migrate(result);
    currentVersion = migration.toVersion;
    appliedMigrations.push(`${migration.fromVersion}->${migration.toVersion}`);
    guard++;
  }

  result[CONFIG_KEYS.SCHEMA_VERSION] = currentVersion;

  return {
    config: result,
    version: currentVersion,
    migrated: appliedMigrations.length > 0,
    appliedMigrations
  };
}
