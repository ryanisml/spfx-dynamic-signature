import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PdfESignatureWebPartStrings';
import PdfESignature from './components/PdfESignature';
import { IPdfESignatureProps } from './components/IPdfESignatureProps';

export interface IPdfESignatureWebPartProps {
  description: string;
}

export default class PdfESignatureWebPart extends BaseClientSideWebPart<IPdfESignatureWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPdfESignatureProps> = React.createElement(
      PdfESignature,
      {
        description: this.properties.description,
        spHttpClient: this.context.spHttpClient,
        currentWebUrl: this.context.pageContext.web.absoluteUrl
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
