import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SiteDesignsStudioWebPartStrings';
import SiteDesignsStudioV2, { ISiteDesignsStudioProps } from './components/SiteDesignsStudio';
import { IApplicationState, initialAppState } from '../../app/ApplicationState';
import configureServices from '../../app/configureServices';

export interface ISiteDesignsStudioWebPartProps {
  description: string;
}

export default class SiteDesignsStudioWebPart extends BaseClientSideWebPart<ISiteDesignsStudioWebPartProps> {

  private applicationState: IApplicationState = initialAppState;

  public onInit(): Promise<void> {
    return super.onInit()
      .then(() => configureServices(this.context))
      .then(serviceScope => {
        this.applicationState = {
          ...this.applicationState,
          serviceScope: serviceScope,
          componentContext: this.context
        };
      });
  }

  public render(): void {

    const element: React.ReactElement<ISiteDesignsStudioProps> = React.createElement(
      SiteDesignsStudioV2,
      {
        description: this.properties.description,
        applicationState: this.applicationState
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [

          ]
        }
      ]
    };
  }
}
