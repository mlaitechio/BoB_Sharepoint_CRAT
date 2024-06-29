import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BobCircularRepositoryWebPartStrings';
import BobCircularRepository from './components/BobCircularRepository';
import { IBobCircularRepositoryProps } from './components/IBobCircularRepositoryProps';
import { Services } from './services/Services';
import { IServices } from './services/IServices';
import { Constants } from './Constants/Constants';
import { IListInfo } from '@pnp/sp/presets/all';

export interface IBobCircularRepositoryWebPartProps {
  description: string;
}

export default class BobCircularRepositoryWebPart extends BaseClientSideWebPart<IBobCircularRepositoryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: any = '';
  private _services: IServices;
  private _circularRepoListID = ``;
  private isUserMaker = false;
  private isUserChecker = false;
  private isUserCompliance = false;

  public render(): void {
    const element: React.ReactElement<IBobCircularRepositoryProps> = React.createElement(
      BobCircularRepository,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context as any,
        services: this._services,
        serverRelativeUrl: this.context.pageContext.legacyPageContext.webServerRelativeUrl,
        circularListID: this._circularRepoListID,
        isUserMaker: this.isUserMaker,
        isUserCompliance: this.isUserCompliance,
        isUserChecker: this.isUserChecker
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // return this._getEnvironmentMessage().then(message => {
    //   this._environmentMessage = message;
    // });

    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit().then(async _ => {

      

      this._services = new Services(this.context);

      let currentUserEmail = this.context.pageContext.user.email;

      await this._services.getListInfo(this.context.pageContext.legacyPageContext.webServerRelativeUrl, Constants.circularList).
        then((value: IListInfo) => {
          this._circularRepoListID = value.Id
        }).catch((error) => {
          console.log(error);
        });

      await this._services.checkIfUserBelongToGroup(Constants.makerGroup, currentUserEmail).then((val) => {
        this.isUserMaker = val;
      }).catch((error) => {
        console.log(error)
      });

      await this._services.checkIfUserBelongToGroup(Constants.checkerGroup, currentUserEmail).then((val) => {
        this.isUserChecker = val;
      }).catch((error) => {
        console.log(error)
      });

      await this._services.checkIfUserBelongToGroup(Constants.complianceGroup, currentUserEmail).then((val) => {
        this.isUserCompliance = val;
      }).catch((error) => {
        console.log(error)
      })

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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
