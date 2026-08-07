import { CONFIG_KEYS, DEFAULT_CONFIG_VALUES } from '../models/config';

/**
 * Resolves the customer-chosen term for a "case" (e.g. "Anfrage", "Case",
 * "Ticket", "Vorgang") from config, defaulting to the original Barlog
 * wording ("Anfrage"/"Anfragen").
 *
 * Deliberately narrow in scope: this only covers the *entity name* used in
 * shared UI chrome (nav tabs, headers, toasts) — the German business
 * vocabulary for individual fields (Kunde, Termin, Bemerkung, ...) is not
 * parameterized, since it is not Barlog-specific and fits the target
 * market (German Mittelstand) as-is.
 *
 * Pure function of a config snapshot — no SharePoint dependency, easy to
 * unit test. Callers pass in whatever `ConfigService.instance.getAll()`
 * (or a subset) returned.
 */
export class LabelService {
  public static getEntityLabelSingular(config: Record<string, string>): string {
    return config[CONFIG_KEYS.ENTITY_LABEL_SINGULAR] || DEFAULT_CONFIG_VALUES[CONFIG_KEYS.ENTITY_LABEL_SINGULAR];
  }

  public static getEntityLabelPlural(config: Record<string, string>): string {
    return config[CONFIG_KEYS.ENTITY_LABEL_PLURAL] || DEFAULT_CONFIG_VALUES[CONFIG_KEYS.ENTITY_LABEL_PLURAL];
  }

  /** "Neue {Anfrage} anlegen" */
  public static getCreateActionLabel(config: Record<string, string>): string {
    return `Neue ${LabelService.getEntityLabelSingular(config)} anlegen`;
  }

  /** "{Anfrage} erfolgreich angelegt!" — number/title is prepended by the caller. */
  public static getCreatedToastSuffix(config: Record<string, string>): string {
    return `${LabelService.getEntityLabelSingular(config)} erfolgreich angelegt!`;
  }

  /** "Alle {Anfragen}" — used as a nav tab / list view title. */
  public static getListViewTitle(config: Record<string, string>): string {
    return `Alle ${LabelService.getEntityLabelPlural(config)}`;
  }
}
