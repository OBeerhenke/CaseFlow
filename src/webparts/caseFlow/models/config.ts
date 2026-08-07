/**
 * Multi-tenant configuration types for CaseFlow.
 *
 * These types describe everything a tenant admin can customize via the
 * WebPart Property Pane (list names, field mapping, status definitions)
 * plus the runtime configuration stored in the CaseFlow_Config SharePoint
 * list (thresholds, entity labels, schema version).
 *
 * Defaults below replicate the original hardcoded Barlog behaviour, so an
 * un-configured install behaves exactly like before.
 */

// ─── Field Mapping ──────────────────────────────────────────────

/**
 * Maps a normalized (logical) field key — the property name CaseFlow uses
 * internally on `ICaseItem` — to the actual SharePoint internal field name
 * for a given tenant's list.
 *
 * Example: a tenant that renamed "field_8" to "Customer" would configure
 * `{ field_8: 'Customer' }`; CaseFlow keeps using `caseItem.field_8`
 * everywhere internally, while SharePointService silently reads/writes
 * "Customer" on the wire.
 */
export interface IFieldMapping {
  [normalizedKey: string]: string;
}

/** Identity mapping — matches the original Barlog list schema. */
export const DEFAULT_CASE_FIELD_MAPPING: IFieldMapping = {
  ID: 'ID',
  Title: 'Title',
  field_1: 'field_1',
  field_2: 'field_2',
  field_4: 'field_4',
  field_5: 'field_5',
  field_6: 'field_6',
  field_8: 'field_8',
  field_9: 'field_9',
  field_10: 'field_10',
  field_11: 'field_11',
  field_12: 'field_12',
  field_13: 'field_13',
  field_14: 'field_14',
  field_15: 'field_15',
  field_16: 'field_16',
  field_17: 'field_17',
  field_18: 'field_18',
  field_19: 'field_19',
  field_20: 'field_20',
  field_21: 'field_21',
  field_22: 'field_22',
  Erledigungsdatum: 'Erledigungsdatum',
  Status: 'Status',
  Ersteller: 'Ersteller',
  Verantwortlicher: 'Verantwortlicher',
  Modified: 'Modified',
  Aufgabenstellung: 'Aufgabenstellung',
  Projektnummer: 'Projektnummer',
  SOP: 'SOP',
  SegCode: 'SegCode',
  AntwortIn: 'AntwortIn',
  Zielpreis: 'Zielpreis'
};

// ─── List Names ─────────────────────────────────────────────────

export interface IListNamesConfig {
  cases: string;
  projects: string;
  categories: string;
  customerApplications: string;
  config: string;
}

export const DEFAULT_LIST_NAMES: IListNamesConfig = {
  cases: 'CaseFlow_Cases',
  projects: 'CaseFlow_Projects',
  categories: 'CaseFlow_Categories',
  customerApplications: 'CaseFlow_CustomerApplications',
  config: 'CaseFlow_Config'
};

// ─── Status Definitions ─────────────────────────────────────────

export interface IStatusDefinition {
  /** Internal key used for comparisons (matches the SharePoint choice value). */
  key: string;
  /** Label shown in the UI. */
  label: string;
  /** CSS color token (hex or named) used to render the status pill/KPI tile. */
  color: string;
}

// Colors match the original hardcoded $clr-* SCSS variables / Dashboard KPI
// tile colors exactly, so an un-configured install looks identical to before.
export const DEFAULT_STATUS_DEFINITIONS: IStatusDefinition[] = [
  { key: 'Termin planen', label: 'Termin planen', color: '#3B82F6' },
  { key: 'läuft planmäßig', label: 'läuft planmäßig', color: '#10B981' },
  { key: 'prüfen', label: 'prüfen', color: '#F59E0B' },
  { key: 'überfällig', label: 'überfällig', color: '#EF4444' },
  { key: 'abgeschlossen', label: 'abgeschlossen', color: '#6B7280' }
];

// ─── Config Schema Versioning ───────────────────────────────────

/** The config schema version this build of CaseFlow expects. */
export const CURRENT_CONFIG_SCHEMA_VERSION = '1.1';

/** Well-known keys stored in the CaseFlow_Config SharePoint list. */
export const CONFIG_KEYS = {
  SCHEMA_VERSION: 'ConfigSchemaVersion',
  DELAY_THRESHOLD_DAYS: 'DelayThresholdDays',
  REVIEW_DAYS: 'PruefenTage',
  ENTITY_LABEL_SINGULAR: 'EntityLabelSingular',
  ENTITY_LABEL_PLURAL: 'EntityLabelPlural',
  /** Server-relative path to the project-data CSV file (e.g. "/sites/{site}/{library}/Projektliste.csv"). */
  PROJECTS_CSV_PATH: 'ProjectsCsvPath'
} as const;

/** Default values for well-known config keys, applied when a key is missing. */
export const DEFAULT_CONFIG_VALUES: Record<string, string> = {
  [CONFIG_KEYS.DELAY_THRESHOLD_DAYS]: '2',
  [CONFIG_KEYS.REVIEW_DAYS]: '3',
  [CONFIG_KEYS.ENTITY_LABEL_SINGULAR]: 'Anfrage',
  [CONFIG_KEYS.ENTITY_LABEL_PLURAL]: 'Anfragen',
  [CONFIG_KEYS.PROJECTS_CSV_PATH]: ''
};
