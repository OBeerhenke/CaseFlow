import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import TaManagement from './components/TaManagement';
import { ITaManagementProps } from './components/ITaManagementProps';
import { SharePointService } from './services/SharePointService';

export interface ITaManagementWebPartProps {
  description: string;
}

export default class TaManagementWebPart extends BaseClientSideWebPart<ITaManagementWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected async onInit(): Promise<void> {
    // Initialize SharePoint service
    SharePointService.init(this.context);
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  public render(): void {
    const element: React.ReactElement<ITaManagementProps> = React.createElement(
      TaManagement,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
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
          header: { description: 'TA-Management Einstellungen' },
          groups: [
            {
              groupName: 'Allgemein',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Beschreibung'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
