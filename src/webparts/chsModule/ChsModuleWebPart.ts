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
// import { sp } from 'sp-pnp-js';
import { isUserInGroup } from '../services/bal/utilities';
import SPCRUDOPS from '../services/dal/spcrudops';
import { sp } from '@pnp/sp/presets/all'; // 👈 for PnPjs

export interface IChsModuleWebPartProps {
  description: string;
}

export default class ChsModuleWebPart extends BaseClientSideWebPart<IChsModuleWebPartProps> {
  isHrGroup1: boolean;
  userLoginName: string;

  // private _isHrGroup1: boolean = false;
  // private _isHrGroup2: boolean = false;
  // private _userLoginName: string = "";
  public async onInit(): Promise<void> {
    await super.onInit();

    sp.setup({
      spfxContext: this.context
    });
  
    // const sp = spfi().using(SPFx(this.context));
    // const currentUser = await sp.web.currentUser();
    const userLoginName: string = this.context.pageContext.legacyPageContext.systemUserKey;
   // ✅ Check group membership
  //  const isHrGroup1: boolean = await isUserInGroup(userLoginName, "HR1_Group");
  //  const isHrGroup2: boolean = await isUserInGroup(userLoginName, "HR2_Group");
  }

  public render(): void {
    const element: React.ReactElement<IChsModuleProps> = React.createElement(
      ChsModule,
      {
        description: this.properties.description,
        currentSPContext: this.context,
        selectedOuterTab: null,
        isHrGroup1: this.isHrGroup1,
        isHrGroup2: this.isHrGroup1,
        userLoginName: this.userLoginName
      }
    );

    ReactDom.render(element, this.domElement);
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
