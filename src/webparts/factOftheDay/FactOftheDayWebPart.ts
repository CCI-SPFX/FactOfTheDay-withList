import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'; 

import * as strings from 'FactOftheDayWebPartStrings';
import FactOftheDay from './components/FactOftheDay';
import { IFactOftheDayProps } from './components/IFactOftheDayProps';

export interface IFactOftheDayWebPartProps {
  listName: string;
  // description: string;
  imgURL: string;
  siteUrl:string;
}

export default class FactOftheDayWebPart extends BaseClientSideWebPart<IFactOftheDayWebPartProps> {

  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFactOftheDayProps> = React.createElement(
      FactOftheDay,
      {
        list: this.properties.listName,
        imgURL: this.properties.imgURL,
        environmentMessage: this._environmentMessage,
        siteUrl: this.context.pageContext.web.absoluteUrl
        // description: this.properties.description,
         
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
            description: "Webpart Propreties"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: "List Name",
                  description: "Enter list name to fetch data from"
                }),
                PropertyPaneTextField('imgURL', {
                  label: "Image URL",
                  description: "Provide image URL"
                })
                /*PropertyPaneTextField('description', {
                  label: "Description",
                  description: "Enter description text",
                  multiline: true
                })
                */
              ]
            }
          ]
        }
      ]
    };
  }
}
