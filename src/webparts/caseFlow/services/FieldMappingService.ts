import { DEFAULT_CASE_FIELD_MAPPING, IFieldMapping } from '../models/config';

/**
 * Translates between CaseFlow's normalized field keys (the property names
 * used throughout the app, e.g. `field_8` for "Kunde") and the actual
 * SharePoint internal field names configured for a given tenant.
 *
 * This is the seam that makes CaseFlow multi-tenant: the rest of the
 * codebase (components, SharePointService callers) only ever deals with
 * normalized keys. SharePointService uses this service right before
 * building `.select()` queries and update payloads, and right after
 * reading raw items back from SharePoint.
 *
 * Pure, dependency-free, and safe to unit test without mocking PnPjs.
 */
export class FieldMappingService {
  private static _mapping: IFieldMapping = { ...DEFAULT_CASE_FIELD_MAPPING };

  /** Configure the active mapping. Missing keys fall back to the default. */
  public static init(overrides?: IFieldMapping): void {
    FieldMappingService._mapping = { ...DEFAULT_CASE_FIELD_MAPPING, ...(overrides || {}) };
  }

  /** Reset to the built-in default mapping (mainly for tests). */
  public static reset(): void {
    FieldMappingService._mapping = { ...DEFAULT_CASE_FIELD_MAPPING };
  }

  public static get current(): IFieldMapping {
    return { ...FieldMappingService._mapping };
  }

  /** Normalized key -> actual SharePoint internal field name. */
  public static resolve(normalizedKey: string): string {
    return FieldMappingService._mapping[normalizedKey] || normalizedKey;
  }

  /** Resolve a list of normalized keys at once (e.g. for a `.select()` call). */
  public static resolveMany(normalizedKeys: string[]): string[] {
    return normalizedKeys.map(key => FieldMappingService.resolve(key));
  }

  /**
   * Translate a raw SharePoint item (actual field names) into CaseFlow's
   * normalized shape. Unmapped raw keys are dropped — callers should only
   * request fields that are part of the mapping.
   */
  public static denormalize<T = Record<string, unknown>>(rawItem: Record<string, unknown>): T {
    const result: Record<string, unknown> = {};
    for (const normalizedKey of Object.keys(FieldMappingService._mapping)) {
      const spKey = FieldMappingService.resolve(normalizedKey);
      if (Object.prototype.hasOwnProperty.call(rawItem, spKey)) {
        result[normalizedKey] = rawItem[spKey];
      }
    }
    return result as T;
  }

  /**
   * Translate a normalized update/create payload (keyed by CaseFlow's
   * internal field names) into actual SharePoint internal field names.
   * Keys not present in the mapping pass through unchanged (e.g.
   * `ErstellerId`, `VerantwortlicherId` which are not part of the case
   * field mapping).
   */
  public static normalize(payload: Record<string, unknown>): Record<string, unknown> {
    const result: Record<string, unknown> = {};
    for (const key of Object.keys(payload)) {
      result[FieldMappingService.resolve(key)] = payload[key];
    }
    return result;
  }
}
