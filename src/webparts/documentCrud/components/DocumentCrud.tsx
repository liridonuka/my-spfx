import * as React from "react";
import styles from "./DocumentCrud.module.scss";
import { IDocumentCrudProps } from "./IDocumentCrudProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp, Item, ItemAddResult, ItemUpdateResult } from "@pnp/sp";
import { IListItem } from "./IListItem";
import { IDocumentCrudState } from "./IDocumentCrudState";
import { Diving } from "./Diving";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";

export default class DocumentCrud extends React.Component<
  IDocumentCrudProps,
  IDocumentCrudState
> {
  // public state = {
  //   status: this.listNotConfigured(this.props)
  //     ? "Please configure list in Web Part properties"
  //     : "Ready",
  //   documentFiles: [],
  //   policyCategories: [],
  //   policyCategoryDropDown: []
  // };
  constructor(props: IDocumentCrudProps, state: IDocumentCrudState) {
    super(props);

    this.state = {
      status: this.listNotConfigured(this.props)
        ? "Please configure list in Web Part properties"
        : "Ready",
      documentFiles: [],
      joinPolicyCategoryItems: [],
      policyCategoryDropDown: [],
      stringPolicyCategory: []
    };
  }

  public render(): React.ReactElement<IDocumentCrudProps> {
    return (
      <div className={styles.documentCrud}>
        <div>
          <Dropdown
            placeHolder="Filter by business functions"
            onChanged={this.filteredFile.bind(this)}
            multiSelect
            options={this.state.policyCategoryDropDown}
            // title={this.state.titleCategory}
          />
        </div>
        {this.state.documentFiles.map(document => (
          <Diving
            key={document.Id}
            name={document.Name}
            id={document.Id}
            documentLink={document.DocumentLink}
          >
            {document.ApprovedDate}
          </Diving>
        ))}
      </div>
    );
  }

  public componentWillMount() {
    this.connectAndRead();
  }

  private connectAndRead(): void {
    this.setState({ documentFiles: [], status: "Loading all items..." });
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.expand("File")
      .getAll()
      .then(
        items => {
          // console.log(items);
          this.getPolicies(items);
        },
        (error: any): void => {
          this.setState({
            status: "Loading all items failed with error: " + error
          });
        }
      );
  }

  private getPolicies(items): void {
    let documentFiles = [];
    let joinPolicyCategoryItems = [];
    items.forEach(policy => {
      documentFiles.push({
        Id: policy.Id,
        Name: policy.File.Name,
        DocumentLink: policy.File.LinkingUrl,
        ApprovedDate: policy.Modified
      });
      policy.MyMetadata.forEach(policyCategory => {
        joinPolicyCategoryItems.push({
          Id: policy.Id,
          MetaData: policyCategory.Label.split(/:/)[1]
        });
      });
    });

    this.dropDownPolicyCategory(joinPolicyCategoryItems);

    this.setState({
      documentFiles,
      joinPolicyCategoryItems,
      status: `Successfully loaded ${documentFiles.length} items`
    });
  }

  private filteredFile(selectedItems) {
    let stringPolicyCategory = [...this.state.stringPolicyCategory];
    if (selectedItems.selected) {
      // add the option if it's checked
      stringPolicyCategory.push(selectedItems.text);
    } else {
      // remove the option if it's unchecked
      const currIndex = stringPolicyCategory.indexOf(selectedItems.text);
      if (currIndex > -1) {
        stringPolicyCategory.splice(currIndex, 1);
      }
    }

    this.setState({
      stringPolicyCategory
    });
    console.log(stringPolicyCategory);

    //finally filter files based on search fieds
    let conncatList = [];
    let filteredPolicies = [];
    this.state.joinPolicyCategoryItems.forEach(filtered => {
      if (filtered["MetaData"] !== undefined) {
        if (filtered["MetaData"].includes("Mjed"))
          filteredPolicies.push({
            Id: filtered.Id,
            MetaData: filtered["MetaData"]
          });
      }
    });

    console.log(filteredPolicies);
    // //remove duplicates and get unique values
    // let uniqueId = new Set(filteredPolicies.map(i => i.Id));
    // let notIn = [];
    // uniqueId.forEach(Id => {
    //   notIn.push(Id);
    // });

    // this.state.documentFiles.map(f =>
    //   notIn
    //     .filter(q => q === f.Id)
    //     .map(m =>
    //       conncatList.push({
    //         Id: f.Id,
    //         Name: f.Name,
    //         DocumentLink: f.DocumentLink,
    //         ApprovedDate: new Date(f.ApprovedDate).toLocaleDateString()
    //       })
    //     )
    // );
  }

  //dropdowns
  private dropDownPolicyCategory(items): void {
    let policyCategoryItems = [];
    items.forEach(item => {
      if (item.MetaData) {
        policyCategoryItems.push({
          key: item.MetaData,
          text: item.MetaData
        });
      }
    });

    let uniqueItems = new Set(policyCategoryItems.map(unique => unique.text));
    let policyCategoryDropDown = [];
    uniqueItems.forEach(uniqueItem => {
      policyCategoryDropDown.push({ key: uniqueItem, text: uniqueItem });
    });
    this.setState({ policyCategoryDropDown });

    // console.log(dropDownItems);
  }

  private listNotConfigured(props: IDocumentCrudProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
