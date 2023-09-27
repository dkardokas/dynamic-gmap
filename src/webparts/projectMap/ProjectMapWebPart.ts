import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProjectMapWebPartStrings';
import ProjectMap from './components/ProjectMap';
import { IProjectMapProps } from './components/IProjectMapProps';
import { SPFI } from "@pnp/sp";
import { getSP } from '../../services/pnpjsConfig';
import { IItems } from '@pnp/sp/items';

export interface IProjectMapWebPartProps {
  description: string;
  mapApiKey: string;
  mapDataListName: string;
  startLat: string;
  startLon: string;
}

export default class ProjectMapWebPart extends BaseClientSideWebPart<IProjectMapWebPartProps> {
  _sp: SPFI;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listItems: IItems;

  public render(): void {
    const element: React.ReactElement<IProjectMapProps> = React.createElement(
      ProjectMap,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        listItems: this._listItems,
        mapApiKey: this.properties.mapApiKey,
        mapDataListName: this.properties.mapDataListName,
        startLat: parseFloat(this.properties.startLat),
        startLon: parseFloat(this.properties.startLon),
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._sp = getSP(this.context);
    this._listItems = await this._getSharePointListItems();
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }


  private async _getSharePointListItems(): Promise<IItems> {
    const listName = this.properties.mapDataListName;
    const pnpList = await this._sp.web.lists.getByTitle(listName);
    return pnpList.items;
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

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
          header: {
            description: "Settings for Project Map web part."
          },
          groups: [
            {
              groupName: "Map Properties",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title"
                }),
                PropertyPaneTextField('mapApiKey', {
                  label: "API Key"
                }),
                PropertyPaneTextField('mapDataListName', {
                  label: "List Name",
                  description: "Title of a list in current site to read data from."
                }),
                PropertyPaneTextField('startLat', {
                  label: "Starting Latitude",
                  description: "Center of the map on load. Must be a valid Latitude value",
                  placeholder: "55.5555",
                  onGetErrorMessage(value) {
                    const floatVal: number = parseFloat(value);
                    if (isNaN(floatVal) || floatVal > 90 || floatVal < -90)
                      return ("Please select a valid Latitude value");
                    return "";
                  },
                }),
                PropertyPaneTextField('startLon', {
                  label: "Starting Longitude",
                  description: "Center of the map on load. Must be a valid Longitude value",
                  placeholder: "-1.1111",
                  onGetErrorMessage(value) {
                    const floatVal: number = parseFloat(value);
                    if (isNaN(floatVal) || floatVal > 180 || floatVal < -180)
                      return ("Please select a valid Longitude value");
                    return "";
                  },
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
