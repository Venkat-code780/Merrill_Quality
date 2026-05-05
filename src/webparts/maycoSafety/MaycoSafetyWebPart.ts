import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MaycoSafetyWebPartStrings';
import MaycoSafety from './components/MaycoSafety';
import { IMaycoSafetyProps } from './components/IMaycoSafetyProps';
import "./components/CSS/style.css";
import "./components/CSS/form-input-style.css";
import 'bootstrap/dist/css/bootstrap.min.css';

export interface IMaycoSafetyWebPartProps {
  description: string;
}

export default class MaycoSafetyWebPart extends BaseClientSideWebPart<IMaycoSafetyWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMaycoSafetyProps> = React.createElement(
      MaycoSafety,
      {
        description: this.properties.description,
        userDisplayName: this.context.pageContext.user.displayName,
        spHttpClient: this.context.spHttpClient,
        spContext: this.context.pageContext.legacyPageContext,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
