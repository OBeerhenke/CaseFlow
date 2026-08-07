import { CONFIG_KEYS, DEFAULT_CONFIG_VALUES } from '../models/config';
import { migrateConfig } from './configMigrations';

/**
 * Minimal storage abstraction for the CaseFlow_Config key-value list.
 * Kept separate from ConfigService — and from any `@pnp/sp` import — so
 * tests can inject an in-memory fake instead of mocking PnPjs. The real
 * SharePoint-backed implementation lives in `SharePointConfigStore.ts`
 * and is only ever imported by production wiring code (CaseFlowWebPart).
 */
export interface IConfigStore {
  getAll(): Promise<Record<string, string>>;
  setValue(key: string, value: string): Promise<void>;
}

export interface ISchemaVersionResult {
  migrated: boolean;
  from: string;
  to: string;
}

/**
 * Reads/writes the CaseFlow_Config SharePoint list (key-value store) with
 * an in-memory cache, and manages schema-version migrations.
 *
 * Singleton, mirroring SharePointService's pattern — initialize once via
 * `ConfigService.init(listName)` during WebPart `onInit`, then use
 * `ConfigService.instance` everywhere else.
 */
export class ConfigService {
  private static _instance: ConfigService;
  private _cache: Record<string, string> | undefined;

  private constructor(private readonly store: IConfigStore) {}

  /**
   * Initialize with any IConfigStore implementation. Production code
   * (CaseFlowWebPart.onInit) passes a `SharePointConfigStore`; tests pass
   * an in-memory fake.
   */
  public static init(store: IConfigStore): void {
    ConfigService._instance = new ConfigService(store);
  }

  public static get instance(): ConfigService {
    return ConfigService._instance;
  }

  private async loadAll(forceReload: boolean = false): Promise<Record<string, string>> {
    if (!this._cache || forceReload) {
      this._cache = await this.store.getAll();
    }
    return this._cache;
  }

  public async getValue(key: string, defaultValue?: string): Promise<string> {
    const all = await this.loadAll();
    if (key in all) return all[key];
    if (defaultValue !== undefined) return defaultValue;
    return DEFAULT_CONFIG_VALUES[key] || '';
  }

  public async setValue(key: string, value: string): Promise<void> {
    await this.store.setValue(key, value);
    if (this._cache) this._cache[key] = value;
  }

  public async getAll(): Promise<Record<string, string>> {
    return { ...(await this.loadAll()) };
  }

  /**
   * Brings the stored config up to `CURRENT_CONFIG_SCHEMA_VERSION`,
   * persisting any keys changed by migration. Safe to call on every app
   * load — a no-op once the config is already current.
   */
  public async ensureSchemaVersion(): Promise<ISchemaVersionResult> {
    const all = await this.loadAll(true);
    const fromVersion = all[CONFIG_KEYS.SCHEMA_VERSION] || '1.0';
    const { config, version, migrated } = migrateConfig(all);

    if (migrated) {
      for (const key of Object.keys(config)) {
        if (all[key] !== config[key]) {
          await this.store.setValue(key, config[key]);
        }
      }
      // A migration may also remove keys (e.g. renaming). SharePoint list
      // rows for removed keys are left in place (harmless, unused) rather
      // than deleted, to keep this operation additive/safe.
      this._cache = config;
    }

    return { migrated, from: fromVersion, to: version };
  }
}
