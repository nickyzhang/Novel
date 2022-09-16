import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'NovelWebPartStrings';
import Novel from './components/Novel';
import { IBpmTodoApply, INovelProps } from './components/INovelProps';

import { AadHttpClient, HttpClientResponse,IHttpClientOptions } from '@microsoft/sp-http';

export interface INovelWebPartProps {
  description: string;
}

export default class NovelWebPart extends BaseClientSideWebPart<INovelWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private aadHttpClient: AadHttpClient;


  public render(): void {
    console.log("Current User Email => "+ this.context.pageContext.user.email);
    const getBpmTodoListRequest = {"email":this.context.pageContext.user.email}
    
    const options: IHttpClientOptions = { 
      headers: new Headers({
          'Accept':'application/json',
        }
      ),
      body: JSON.stringify(getBpmTodoListRequest)
    };
    this.aadHttpClient.post('https://kupotech.mynatapp.cc/page/api/getBpmTodoList',AadHttpClient.configurations.v1,options)
    // this.aadHttpClient.get('https://kupotech.azurewebsites.net/page/api/getBpmTodoList',AadHttpClient.configurations.v1)
      .then((response: HttpClientResponse):Promise<any> => {
        return response.json();
      })
      .then((novels:any):void => {
        const element: React.ReactElement<INovelProps> = React.createElement(
          Novel,
          {
            description: this.properties.description,
            isDarkTheme: this._isDarkTheme,
            environmentMessage: this._environmentMessage,
            hasTeamsContext: !!this.context.sdks.microsoftTeams,
            userDisplayName: this.context.pageContext.user.displayName,
            datalist:novels
          }
        );
    
        ReactDom.render(element, this.domElement);
      },(err:any):void => {
          console.error("Fail to get bpm todo list: ",err);
      });
    
    
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        // .getClient('https://kupotech.azurewebsites.net')
        .getClient('1e247003-d377-4eb8-a8ea-d0db8d4ad962')
        .then((client: AadHttpClient): void => {
          this.aadHttpClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
