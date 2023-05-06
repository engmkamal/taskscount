import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'TaskscountWebPartStrings';
import Taskscount from './components/Taskscount';
import { ITaskscountProps } from './components/ITaskscountProps';

import { getSP } from '../../pnpjsConfig';

export interface ITaskscountWebPartProps {
  description: string;
}

export default class TaskscountWebPart extends BaseClientSideWebPart<ITaskscountWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ITaskscountProps> = React.createElement(
      Taskscount,
      {        
        //isDarkTheme: this._isDarkTheme,
        //environmentMessage: this._environmentMessage,
        //hasTeamsContext: !!this.context.sdks.microsoftTeams,
        //userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> { 

    getSP(this.context);

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
