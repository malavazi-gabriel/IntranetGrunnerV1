import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CentralAtalhosGrunnerWebPartStrings';
import CentralAtalhosGrunner from './components/CentralAtalhosGrunner';
import { ICentralAtalhosGrunnerProps } from './components/ICentralAtalhosGrunnerProps';

export interface ICentralAtalhosGrunnerWebPartProps {
  description: string;
}

export default class CentralAtalhosGrunnerWebPart extends BaseClientSideWebPart<ICentralAtalhosGrunnerWebPartProps> {
  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<ICentralAtalhosGrunnerProps> = React.createElement(
      CentralAtalhosGrunner,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: '',
        hasTeamsContext: false,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
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