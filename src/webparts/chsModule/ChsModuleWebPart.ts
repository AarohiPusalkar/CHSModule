import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'ChsModuleWebPartStrings';
import ChsModule from './components/ChsModule';
import { IChsModuleProps } from './components/IChsModuleProps';
export interface IChsModuleWebPartProps {
  description: string;
}
export default class ChsModuleWebPart extends BaseClientSideWebPart<IChsModuleWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IChsModuleProps > = React.createElement(
      ChsModule,
      {
        description: this.properties.description,
        currentSPContext: this.context,
        selectedOuterTab:null
      }
    );
    ReactDom.render(element, this.domElement);
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }
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
