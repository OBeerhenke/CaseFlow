import { migrateConfig, IConfigMigration } from './configMigrations';
import { CONFIG_KEYS, CURRENT_CONFIG_SCHEMA_VERSION } from '../models/config';

describe('migrateConfig', () => {
  it('is a no-op for a config already at the target version', () => {
    const config = { [CONFIG_KEYS.SCHEMA_VERSION]: CURRENT_CONFIG_SCHEMA_VERSION, DelayThresholdDays: '2' };
    const result = migrateConfig(config);

    expect(result.migrated).toBe(false);
    expect(result.version).toBe(CURRENT_CONFIG_SCHEMA_VERSION);
    expect(result.appliedMigrations).toEqual([]);
    expect(result.config).toEqual(config);
  });

  it('defaults a config with no schema version marker to "1.0" and migrates forward', () => {
    const config = { ReviewDays: '3' };
    const result = migrateConfig(config);

    expect(result.migrated).toBe(true);
    expect(result.version).toBe('1.1');
    expect(result.appliedMigrations).toEqual(['1.0->1.1']);
  });

  it('renames the legacy "ReviewDays" key to "PruefenTage" (1.0 -> 1.1 example migration)', () => {
    const config = { [CONFIG_KEYS.SCHEMA_VERSION]: '1.0', ReviewDays: '5' };
    const result = migrateConfig(config);

    expect(result.config.PruefenTage).toBe('5');
    expect(result.config).not.toHaveProperty('ReviewDays');
    expect(result.config[CONFIG_KEYS.SCHEMA_VERSION]).toBe('1.1');
  });

  it('does not overwrite an already-present "PruefenTage" value with the legacy one', () => {
    const config = { [CONFIG_KEYS.SCHEMA_VERSION]: '1.0', ReviewDays: '5', PruefenTage: '9' };
    const result = migrateConfig(config);

    expect(result.config.PruefenTage).toBe('9');
    expect(result.config).not.toHaveProperty('ReviewDays');
  });

  it('preserves unrelated config keys untouched', () => {
    const config = { [CONFIG_KEYS.SCHEMA_VERSION]: '1.0', DelayThresholdDays: '2', EntityLabelSingular: 'Ticket' };
    const result = migrateConfig(config);

    expect(result.config.DelayThresholdDays).toBe('2');
    expect(result.config.EntityLabelSingular).toBe('Ticket');
  });

  it('walks multiple chained migrations in sequence', () => {
    const migrations: IConfigMigration[] = [
      { fromVersion: '1.0', toVersion: '1.1', migrate: (c) => ({ ...c, stepA: 'done' }) },
      { fromVersion: '1.1', toVersion: '1.2', migrate: (c) => ({ ...c, stepB: 'done' }) }
    ];
    const result = migrateConfig({ [CONFIG_KEYS.SCHEMA_VERSION]: '1.0' }, migrations, '1.2');

    expect(result.version).toBe('1.2');
    expect(result.appliedMigrations).toEqual(['1.0->1.1', '1.1->1.2']);
    expect(result.config.stepA).toBe('done');
    expect(result.config.stepB).toBe('done');
  });

  it('stops and leaves the version as-is when no migration path exists from the current version', () => {
    const migrations: IConfigMigration[] = [
      { fromVersion: '2.0', toVersion: '2.1', migrate: (c) => c }
    ];
    const result = migrateConfig({ [CONFIG_KEYS.SCHEMA_VERSION]: '1.0' }, migrations, '2.1');

    expect(result.version).toBe('1.0');
    expect(result.migrated).toBe(false);
    expect(result.appliedMigrations).toEqual([]);
  });

  it('does not loop forever when migrations form a cycle', () => {
    const migrations: IConfigMigration[] = [
      { fromVersion: '1.0', toVersion: '1.1', migrate: (c) => c },
      { fromVersion: '1.1', toVersion: '1.0', migrate: (c) => c }
    ];
    const result = migrateConfig({ [CONFIG_KEYS.SCHEMA_VERSION]: '1.0' }, migrations, '9.9');

    // Guarded loop must terminate; exact stopping version is not the point, termination is.
    expect(result.appliedMigrations.length).toBeLessThanOrEqual(migrations.length + 1);
  });
});
