import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpaFeedWebpartWebPartStrings';
import SpaFeedWebpart from './components/SpaFeedWebpart';
import { ISpaFeedWebpartProps } from './components/ISpaFeedWebpartProps';

export interface ISpaFeedWebpartWebPartProps {
  description: string;
}

export default class SpaFeedWebpartWebPart extends BaseClientSideWebPart<ISpaFeedWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpaFeedWebpartProps > = React.createElement(
      SpaFeedWebpart,
      {
        spHttpClient:this.context.spHttpClient,
        siteUrl:this.context.pageContext.web.absoluteUrl,
        context:this.context
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
