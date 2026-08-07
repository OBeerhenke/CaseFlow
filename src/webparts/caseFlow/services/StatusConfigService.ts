import { DEFAULT_STATUS_DEFINITIONS, IStatusDefinition } from '../models/config';

/**
 * Holds the active set of status definitions (key, label, color) for the
 * tenant. Configured via the WebPart Property Pane (JSON); defaults to the
 * original Barlog 5-status set.
 *
 * Pure/static, no SharePoint dependency — easy to unit test and to use from
 * any component without prop-drilling.
 */
export class StatusConfigService {
  private static _definitions: IStatusDefinition[] = DEFAULT_STATUS_DEFINITIONS;

  public static init(definitions?: IStatusDefinition[]): void {
    StatusConfigService._definitions = definitions && definitions.length > 0
      ? definitions
      : DEFAULT_STATUS_DEFINITIONS;
  }

  public static reset(): void {
    StatusConfigService._definitions = DEFAULT_STATUS_DEFINITIONS;
  }

  public static getAll(): IStatusDefinition[] {
    return [...StatusConfigService._definitions];
  }

  public static getByKey(key: string): IStatusDefinition | undefined {
    return StatusConfigService._definitions.find(d => d.key === key);
  }

  public static getLabel(key: string): string {
    return StatusConfigService.getByKey(key)?.label ?? key;
  }

  public static getColor(key: string): string {
    return StatusConfigService.getByKey(key)?.color ?? '#8a8f98';
  }
}
