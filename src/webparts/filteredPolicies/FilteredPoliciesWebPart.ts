import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "FilteredPoliciesWebPartStrings";
import FilteredPolicies from "./components/FilteredPolicies";
import { IFilteredPoliciesProps } from "./components/IFilteredPoliciesProps";

// export interface IFilteredPoliciesWebPartProps {
//   description: string;
// }

export default class FilteredPoliciesWebPart extends BaseClientSideWebPart<
  IFilteredPoliciesProps
> {
  public render(): void {
    const element: React.ReactElement<
      IFilteredPoliciesProps
    > = React.createElement(FilteredPolicies, {
      listName: this.properties.listName,
      context: this.context
    });

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
