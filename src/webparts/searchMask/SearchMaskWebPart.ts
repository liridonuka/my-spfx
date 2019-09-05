import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "SearchMaskWebPartStrings";
import SearchMask from "./components/SearchMask";
import { ISearchMaskProps } from "./components/ISearchMaskProps";

// export interface ISearchMaskWebPartProps {
//   description: string;
// }

export default class SearchMaskWebPart extends BaseClientSideWebPart<
  ISearchMaskProps
> {
  public render(): void {
    const element: React.ReactElement<ISearchMaskProps> = React.createElement(
      SearchMask,
      {
        listName: this.properties.listName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("listName", {
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
