import { ConfigService, IConfigStore } from './ConfigService';
import { CONFIG_KEYS, CURRENT_CONFIG_SCHEMA_VERSION } from '../models/config';

/** In-memory fake store — avoids mocking PnPjs entirely. */
class FakeConfigStore implements IConfigStore {
  public setValueCalls: { key: string; value: string }[] = [];

  public constructor(private data: Record<string, string> = {}) {}

  public async getAll(): Promise<Record<string, string>> {
    return { ...this.data };
  }

  public async setValue(key: string, value: string): Promise<void> {
    this.data[key] = value;
    this.setValueCalls.push({ key, value });
  }
}

describe('ConfigService', () => {
  describe('getValue', () => {
    it('returns the stored value when the key exists', async () => {
      const store = new FakeConfigStore({ DelayThresholdDays: '5' });
      ConfigService.init(store);

      expect(await ConfigService.instance.getValue('DelayThresholdDays', '2')).toBe('5');
    });

    it('returns the provided default when the key is missing', async () => {
      const store = new FakeConfigStore({});
      ConfigService.init(store);

      expect(await ConfigService.instance.getValue('DelayThresholdDays', '2')).toBe('2');
    });

    it('falls back to the well-known default value when no explicit default is given', async () => {
      const store = new FakeConfigStore({});
      ConfigService.init(store);

      expect(await ConfigService.instance.getValue(CONFIG_KEYS.REVIEW_DAYS)).toBe('3');
    });

    it('caches config after the first read (does not call the store again)', async () => {
      const store = new FakeConfigStore({ DelayThresholdDays: '2' });
      const getAllSpy = jest.spyOn(store, 'getAll');
      ConfigService.init(store);

      await ConfigService.instance.getValue('DelayThresholdDays');
      await ConfigService.instance.getValue('DelayThresholdDays');

      expect(getAllSpy).toHaveBeenCalledTimes(1);
    });
  });

  describe('setValue', () => {
    it('writes through to the store and updates the cache', async () => {
      const store = new FakeConfigStore({ DelayThresholdDays: '2' });
      ConfigService.init(store);

      await ConfigService.instance.setValue('DelayThresholdDays', '4');

      expect(store.setValueCalls).toEqual([{ key: 'DelayThresholdDays', value: '4' }]);
      expect(await ConfigService.instance.getValue('DelayThresholdDays')).toBe('4');
    });
  });

  describe('getAll', () => {
    it('returns a snapshot copy — mutating it does not affect the service', async () => {
      const store = new FakeConfigStore({ DelayThresholdDays: '2' });
      ConfigService.init(store);

      const all = await ConfigService.instance.getAll();
      all.DelayThresholdDays = 'mutated';

      expect(await ConfigService.instance.getValue('DelayThresholdDays')).toBe('2');
    });
  });

  describe('ensureSchemaVersion', () => {
    it('is a no-op when the config is already at the current schema version', async () => {
      const store = new FakeConfigStore({ [CONFIG_KEYS.SCHEMA_VERSION]: CURRENT_CONFIG_SCHEMA_VERSION, DelayThresholdDays: '2' });
      ConfigService.init(store);

      const result = await ConfigService.instance.ensureSchemaVersion();

      expect(result).toEqual({ migrated: false, from: CURRENT_CONFIG_SCHEMA_VERSION, to: CURRENT_CONFIG_SCHEMA_VERSION });
      expect(store.setValueCalls).toEqual([]);
    });

    it('migrates a legacy config (no schema version, old key name) and persists the changes', async () => {
      const store = new FakeConfigStore({ ReviewDays: '5' });
      ConfigService.init(store);

      const result = await ConfigService.instance.ensureSchemaVersion();

      expect(result.migrated).toBe(true);
      expect(result.from).toBe('1.0');
      expect(result.to).toBe('1.1');

      // The renamed key + the new schema version marker must both be persisted
      const persistedKeys = store.setValueCalls.map(c => c.key);
      expect(persistedKeys).toContain(CONFIG_KEYS.REVIEW_DAYS);
      expect(persistedKeys).toContain(CONFIG_KEYS.SCHEMA_VERSION);

      expect(await ConfigService.instance.getValue(CONFIG_KEYS.REVIEW_DAYS)).toBe('5');
    });

    it('updates the in-memory cache after migrating, so subsequent reads see the new shape', async () => {
      const store = new FakeConfigStore({ ReviewDays: '7' });
      ConfigService.init(store);

      await ConfigService.instance.ensureSchemaVersion();

      expect(await ConfigService.instance.getValue(CONFIG_KEYS.REVIEW_DAYS)).toBe('7');
      expect(await ConfigService.instance.getValue('ReviewDays', 'gone')).toBe('gone');
    });
  });
});
