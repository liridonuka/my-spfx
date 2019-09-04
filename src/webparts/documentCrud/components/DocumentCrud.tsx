import * as React from "react";
import styles from "./DocumentCrud.module.scss";
import { IDocumentCrudProps } from "./IDocumentCrudProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp, Item, ItemAddResult, ItemUpdateResult, Web } from "@pnp/sp";
import { IListItem } from "./IListItem";
import { IDocumentCrudState } from "./IDocumentCrudState";
import { Diving } from "./Diving";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption
} from "office-ui-fabric-react/lib/Dropdown";
// import pnp from "sp-pnp-js";
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
      internalPolicies: [],
      documentFiles: [],
      joinPolicyCategoryItems: [],
      policyCategoryDropDown: [],
      stringPolicyCategory: [],
      joinRegulatoryTopicItems: [],
      regulatoryTopicDropDown: [],
      stringRegulatoryTopic: []
    };
  }

  public render(): React.ReactElement<IDocumentCrudProps> {
    return (
      <div className={styles.documentCrud}>
        <div>
          <div className="row">
            <Dropdown
              placeHolder="Filter by business functions"
              onChanged={this.filteredFile.bind(this)}
              multiSelect
              options={this.state.policyCategoryDropDown}
              // title={this.state.titleCategory}
            />
            <Dropdown
              placeHolder="Filter by regulatory topic"
              onChanged={this.filteredFile.bind(this)}
              multiSelect
              options={this.state.regulatoryTopicDropDown}
              // title={this.state.titleCategory}
            />
          </div>
          <div>
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
        </div>
      </div>
    );
  }

  public componentWillMount() {
    this.connectAndRead();
    return new Promise<void>(
      (resolve: () => void, reject: (error?: any) => void): void => {
        sp.setup({
          sp: {
            headers: {
              Accept: "application/json; odata=nometadata"
            }
          }
        });
        resolve();
      }
    );
  }

  private connectAndRead(): void {
    this.setState({ documentFiles: [], status: "Loading all items..." });
    let web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists
      .getByTitle(this.props.listName)
      .items.expand("File")
      .getAll()
      .then(
        items => {
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
    const monthNames = this.monthName();
    let documentFiles = [];
    let joinPolicyCategoryItems = [];
    let joinRegulatoryTopicItems = [];
    let itemList = [];
    items.forEach(policy => {
      itemList.push({
        Id: policy.Id,
        Name: policy.File.Name,
        DocumentLink: policy.File.LinkingUrl,
        ApprovedDate: policy.Date_x0020_of_x0020_approval
      });
      policy.MyMetadata.forEach(policyCategory => {
        joinPolicyCategoryItems.push({
          Id: policy.Id,
          PolicyCategory: policyCategory.Label.split(/:/)[1]
        });
      });
      policy.RegulatoryTopic.forEach(regulatoryTopic => {
        joinRegulatoryTopicItems.push({
          Id: policy.Id,
          RegulatoryTopic: regulatoryTopic.Label.split(/:/)[1]
        });
      });
    });
    joinPolicyCategoryItems.forEach(j => {
      itemList
        .filter(f => f.Id === j.Id)
        .map(item => (item.PolicyCategory += ";" + j.PolicyCategory));
    });
    joinRegulatoryTopicItems.forEach(j => {
      itemList
        .filter(f => f.Id === j.Id)
        .map(item => (item.RegulatoryTopic += ";" + j.RegulatoryTopic));
    });
    itemList.forEach(policy => {
      documentFiles.push({
        Id: policy.Id,
        Name: policy.Name,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory
          ? policy.PolicyCategory.split("undefined;").pop()
          : "",
        RegulatoryTopic: policy.RegulatoryTopic
          ? policy.RegulatoryTopic.split("undefined;").pop()
          : ""
      });
    });
    // console.log(
    //   documentFiles.filter(
    //     f =>
    //       f.PolicyCategory.includes("Operations") ||
    //       f.PolicyCategory.includes("Audit")
    //   )
    // );
    // // console.log(joinPolicyCategoryItems);
    // documentFiles = documentFiles.filter(
    //   f =>
    //     f.PolicyCategory.includes("Operations") ||
    //     f.PolicyCategory.includes("Audit")
    // );
    this.setState({
      documentFiles,
      internalPolicies: documentFiles,
      joinPolicyCategoryItems,
      joinRegulatoryTopicItems,
      status: `Successfully loaded ${documentFiles.length} items`
    });
    this.dropDownPolicyCategory();
    // this.dropDownRegulatoryTopic();
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

    let filteredJoinPolicyCategories = [];
    let filteredList = [...this.state.internalPolicies];
    stringPolicyCategory.forEach(policy => {
      this.state.joinPolicyCategoryItems
        .filter(f => f["PolicyCategory"] === policy)
        .map(join =>
          filteredJoinPolicyCategories.push({
            Id: join.Id,
            PolicyCtaegory: join["PolicyCategory"]
          })
        );
    });
    //remove duplicates and get unique values for filtered policies in join Policy Category
    let uniqueId = new Set(filteredJoinPolicyCategories.map(i => i.Id));
    let notIn = [];
    uniqueId.forEach(Id => {
      notIn.push(Id);
    });

    const filteredPolicies = filteredList.filter(({ Id: Idv }) =>
      filteredJoinPolicyCategories.some(({ Id: idc }) => Idv === idc)
    );
    // console.log(filteredPolicies);
    this.setState({
      documentFiles:
        stringPolicyCategory.length > 0
          ? filteredPolicies
          : this.state.internalPolicies,
      joinRegulatoryTopicItems: filteredPolicies
    });
    console.log(this.state.joinRegulatoryTopicItems);
    // this.dropDownRegulatoryTopic(filteredPolicies);
  }

  //dropdowns
  private dropDownPolicyCategory(): void {
    let policyCategoryItems = [];

    this.state.joinPolicyCategoryItems.forEach(item => {
      if (item["PolicyCategory"]) {
        policyCategoryItems.push({
          key: item["PolicyCategory"],
          text: item["PolicyCategory"]
        });
      }
    });

    let uniqueItems = new Set(policyCategoryItems.map(unique => unique.text));
    let policyCategoryDropDown = [];
    uniqueItems.forEach(uniqueItem => {
      policyCategoryDropDown.push({ key: uniqueItem, text: uniqueItem });
    });
    this.setState({ policyCategoryDropDown });
  }

  private dropDownRegulatoryTopic(items): void {
    let regulatoryTopicItems = [];
    const filteredPolicies = this.state.joinRegulatoryTopicItems.filter(
      ({ Id: Idv }) => items.some(({ Id: idc }) => Idv === idc)
    );
    filteredPolicies.forEach(item => {
      if (item["RegulatoryTopic"]) {
        regulatoryTopicItems.push({
          key: item["RegulatoryTopic"],
          text: item["RegulatoryTopic"]
        });
      }
    });

    let uniqueItems = new Set(regulatoryTopicItems.map(unique => unique.text));
    let regulatoryTopicDropDown = [];
    uniqueItems.forEach(uniqueItem => {
      regulatoryTopicDropDown.push({ key: uniqueItem, text: uniqueItem });
    });
    this.setState({ regulatoryTopicDropDown });
  }

  private monthName() {
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December"
    ];
    return monthNames;
  }

  private listNotConfigured(props: IDocumentCrudProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
