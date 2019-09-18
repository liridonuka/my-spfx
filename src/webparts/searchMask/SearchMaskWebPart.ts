import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
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
        joinListName: this.properties.joinListName,
        context: this.context,
        commentDialogTitle: this.properties.commentDialogTitle,
        commentDialogSubTitle: this.properties.commentDialogSubTitle,
        myWorkingSpace: this.properties.myWorkingSpace
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
                  label: strings.DescriptionLibrary
                }),
                PropertyPaneTextField("joinListName", {
                  label: strings.DescriptionJoinList
                }),
                PropertyPaneToggle("myWorkingSpace", {
                  label: strings.DescriptionMyWorkingSpace,
                  onText: "Yes",
                  offText: "No"
                }),
                PropertyPaneTextField("commentDialogTitle", {
                  label: strings.CommentDialogTitle
                }),
                PropertyPaneTextField("commentDialogSubTitle", {
                  label: strings.CommentDialogSubTitle
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
