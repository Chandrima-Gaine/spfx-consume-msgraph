import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ConsumeMsGraphWebPartStrings';
import ConsumeMsGraph from './components/ConsumeMsGraph';
import { IConsumeMsGraphProps } from './components/IConsumeMsGraphProps';

export interface IConsumeMsGraphWebPartProps {
  description: string;
}

export default class ConsumeMsGraphWebPart extends BaseClientSideWebPart<IConsumeMsGraphWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IConsumeMsGraphProps > = React.createElement(
      ConsumeMsGraph,
      {
        description: this.properties.description,
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
