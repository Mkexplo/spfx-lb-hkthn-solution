import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'LbEmpProfileWpWebPartStrings';
import LbEmpProfileWp from './components/LbEmpProfileWp';
import { ILbEmpProfileWpProps } from './components/ILbEmpProfileWpProps';

export interface ILbEmpProfileWpWebPartProps {
  description: string;
  userEmail:string
}

export default class LbEmpProfileWpWebPart extends BaseClientSideWebPart<ILbEmpProfileWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ILbEmpProfileWpProps> = React.createElement(
      LbEmpProfileWp,
      {
        description: this.properties.description,
        userEmail: this.context.pageContext.user.email,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
   return Promise.resolve();
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
