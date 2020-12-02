import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'M365ComponentsLibrarySampleWebPartStrings';
import M365ComponentsLibrarySample from './components/M365ComponentsLibrarySample';

export interface IM365ComponentsLibrarySampleWebPartProps {
  description: string;
}

export default class M365ComponentsLibrarySampleWebPart extends BaseClientSideWebPart<IM365ComponentsLibrarySampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<{}> = React.createElement(
      M365ComponentsLibrarySample,
      {
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
