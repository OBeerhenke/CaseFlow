import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import App from './components/App';
import { ICaseFlowProps } from './components/ICaseFlowProps';
import { SharePointService } from './services/SharePointService';
import { ConfigService } from './services/ConfigService';
import { SharePointConfigStore } from './services/SharePointConfigStore';
import { FieldMappingService } from './services/FieldMappingService';
import { StatusConfigService } from './services/StatusConfigService';
import { DEFAULT_LIST_NAMES, IFieldMapping, IListNamesConfig, IStatusDefinition } from './models/config';

export interface ICaseFlowWebPartProps {
  description: string;
  casesListName: string;
  categoriesListName: string;
  customerApplicationsListName: string;
  configListName: string;
  fieldMappingJson: string;
  statusDefinitionsJson: string;
}

/** Parse a JSON Property Pane field, falling back to `fallback` and logging a warning on invalid input. */
function parseJsonProperty<T>(raw: string | undefined, label: string, fallback: T): T {
  if (!raw || !raw.trim()) return fallback;
  try {
    return JSON.parse(raw) as T;
  } catch (e) {
    console.warn(`CaseFlow: could not parse "${label}" Property Pane JSON — using defaults.`, e);
    return fallback;
  }
}

export default class CaseFlowWebPart extends BaseClientSideWebPart<ICaseFlowWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected async onInit(): Promise<void> {
    const listNames: IListNamesConfig = {
      cases: this.properties.casesListName || DEFAULT_LIST_NAMES.cases,
      projects: DEFAULT_LIST_NAMES.projects,
      categories: this.properties.categoriesListName || DEFAULT_LIST_NAMES.categories,
      customerApplications: this.properties.customerApplicationsListName || DEFAULT_LIST_NAMES.customerApplications,
      config: this.properties.configListName || DEFAULT_LIST_NAMES.config
    };

    const fieldMappingOverrides = parseJsonProperty<IFieldMapping>(this.properties.fieldMappingJson, 'fieldMappingJson', {});
    const statusDefinitions = parseJsonProperty<IStatusDefinition[]>(this.properties.statusDefinitionsJson, 'statusDefinitionsJson', undefined as unknown as IStatusDefinition[]);

    FieldMappingService.init(fieldMappingOverrides);
    StatusConfigService.init(statusDefinitions);

    SharePointService.init(this.context, listNames);
    ConfigService.init(new SharePointConfigStore(listNames.config));

    try {
      await ConfigService.instance.ensureSchemaVersion();
    } catch (e) {
      console.warn('CaseFlow: could not verify/migrate config schema version.', e);
    }

    this._environmentMessage = await this._getEnvironmentMessage();
  }

  public render(): void {
    const element: React.ReactElement<ICaseFlowProps> = React.createElement(
      App,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userLoginName: this.context.pageContext.user.loginName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      const ctx = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch (ctx.app.host.name) {
        case 'Office': return 'Office';
        case 'Outlook': return 'Outlook';
        case 'Teams':
        case 'TeamsModern': return 'Teams';
        default: return 'Unknown';
      }
    }
    return 'SharePoint';
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'CaseFlow Einstellungen' },
          groups: [
            {
              groupName: 'Allgemein',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Beschreibung'
                })
              ]
            },
            {
              groupName: 'SharePoint-Listen',
              groupFields: [
                PropertyPaneTextField('casesListName', {
                  label: 'Cases-Liste',
                  description: `Standard: ${DEFAULT_LIST_NAMES.cases}`
                }),
                PropertyPaneTextField('categoriesListName', {
                  label: 'Kategorien-Liste',
                  description: `Standard: ${DEFAULT_LIST_NAMES.categories}`
                }),
                PropertyPaneTextField('customerApplicationsListName', {
                  label: 'Kunden-Anwendungen-Liste',
                  description: `Standard: ${DEFAULT_LIST_NAMES.customerApplications}`
                }),
                PropertyPaneTextField('configListName', {
                  label: 'Config-Liste',
                  description: `Standard: ${DEFAULT_LIST_NAMES.config}`
                })
              ]
            },
            {
              groupName: 'Feld-Mapping (fortgeschritten)',
              groupFields: [
                PropertyPaneTextField('fieldMappingJson', {
                  label: 'Feld-Mapping (JSON)',
                  description: 'Überschreibt einzelne Feldnamen, z.B. {"field_8": "Customer", "field_6": "PlannedDate"}. Leer lassen für Standard-Mapping.',
                  multiline: true,
                  rows: 6
                })
              ]
            },
            {
              groupName: 'Status-Definitionen (fortgeschritten)',
              groupFields: [
                PropertyPaneTextField('statusDefinitionsJson', {
                  label: 'Status-Definitionen (JSON)',
                  description: 'Array von {"key","label","color"}. Leer lassen für Standard-Stati (Termin planen / läuft planmäßig / prüfen / überfällig / abgeschlossen).',
                  multiline: true,
                  rows: 8
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
