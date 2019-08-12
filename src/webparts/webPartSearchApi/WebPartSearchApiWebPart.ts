import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { taxonomy } from '@pnp/sp-taxonomy';
import * as strings from 'WebPartSearchApiWebPartStrings';
import WebPartSearchApi from './components/WebPartSearchApi';
import { IWebPartSearchApiProps } from './components/IWebPartSearchApiProps';

export interface IWebPartSearchApiWebPartProps {
    idTermSet: string;
}

export default class WebPartSearchApiWebPart extends BaseClientSideWebPart<IWebPartSearchApiWebPartProps> {


    public render(): void {
        taxonomy.setup({
            spfxContext: this.context
        });
    const element: React.ReactElement<IWebPartSearchApiProps > = React.createElement(
      WebPartSearchApi,
      {
          idTermSet: this.properties.idTermSet || 'c1a40892-eea2-4fd1-b75e-179737829e20',
          context: this.context
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
                PropertyPaneTextField('nameList', {
                  label: strings.IdTermSetFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
