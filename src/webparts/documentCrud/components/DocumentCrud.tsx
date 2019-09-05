import * as React from "react";
import styles from "./DocumentCrud.module.scss";
import { IDocumentCrudProps } from "./IDocumentCrudProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Web } from "@pnp/sp";
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
      stringRegulatoryTopic: [],
      monthDropDown: [],
      stringMonth: []
    };
  }

  public render(): React.ReactElement<IDocumentCrudProps> {
    return (
      <div className={styles.documentCrud}>
        <div>
          <div className="row">
            <Dropdown
              placeHolder="Filter by business function"
              onChanged={this.filterByPolicyCategory.bind(this)}
              multiSelect
              options={this.state.policyCategoryDropDown}
              // title={this.state.titleCategory}
            />
            <Dropdown
              placeHolder="Filter by regulatory topic"
              onChanged={this.filterByRegulatoryTopic.bind(this)}
              multiSelect
              options={this.state.regulatoryTopicDropDown}
              // title={this.state.titleCategory}
            />
            <Dropdown
              placeHolder="Filter by approval month"
              onChanged={this.filterByRegulatoryTopic.bind(this)}
              multiSelect
              options={this.state.monthDropDown}
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
      policy.Policy_x0020_Category.forEach(policyCategory => {
        joinPolicyCategoryItems.push({
          Id: policy.Id,
          PolicyCategory: policyCategory.Label.split(/:/)[1]
        });
      });
      policy.Regulatory_x0020_Topic.forEach(regulatoryTopic => {
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
          : "",
        Month: monthNames[new Date(policy.ApprovedDate).getMonth()]
      });
    });
    this.setState({
      documentFiles,
      internalPolicies: documentFiles,
      joinPolicyCategoryItems,
      joinRegulatoryTopicItems,
      status: `Successfully loaded ${documentFiles.length} items`
    });
    this.dropDownPolicyCategory();
    this.dropDownRegulatoryTopic();
    this.dropDownMonth();
  }

  private filterByPolicyCategory(selectedItems) {
    console.log(this.state.documentFiles);
    const stringPolicyCategory = this.selectItems(
      selectedItems,
      this.state.stringPolicyCategory
    );
    this.setState({
      stringPolicyCategory
    });
    const clonedList = this.clonedList();
    let filteredList = [];
    stringPolicyCategory.forEach(s => {
      clonedList
        .filter(f => f.PolicyCategory.includes(s))
        .map(item => filteredList.push(item));
    });
    let uniqueItems = new Set(filteredList.map(unique => unique));
    let documentFiles = [];
    uniqueItems.forEach(u => {
      documentFiles.push(u);
    });
    this.setState({
      documentFiles:
        stringPolicyCategory.length > 0 ? documentFiles : clonedList
    });
    this.dropDownRegulatoryTopic(
      stringPolicyCategory.length > 0 ? documentFiles : clonedList
    );
    this.dropDownMonth(
      stringPolicyCategory.length > 0 ? documentFiles : clonedList
    );
  }
  private filterByRegulatoryTopic(selectedItems) {
    const stringRegulatoryTopic = this.selectItems(
      selectedItems,
      this.state.stringPolicyCategory
    );
    this.setState({
      stringRegulatoryTopic
    });
    const clonedList = this.clonedList();
    let filteredList = [];
    stringRegulatoryTopic.forEach(s => {
      clonedList
        .filter(f => f.RegulatoryTopic.includes(s))
        .map(item => filteredList.push(item));
    });
    let uniqueItems = new Set(filteredList.map(unique => unique));
    let documentFiles = [];
    uniqueItems.forEach(u => {
      documentFiles.push(u);
    });
    this.setState({
      documentFiles:
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
    });
    this.dropDownPolicyCategory(
      stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
    );
  }
  private selectItems(selectedItems, stringList) {
    let result = [...stringList];
    if (selectedItems.selected) {
      // add the option if it's checked
      result.push(selectedItems.text);
    } else {
      // remove the option if it's unchecked
      const currIndex = result.indexOf(selectedItems.text);
      if (currIndex > -1) {
        result.splice(currIndex, 1);
      }
    }
    return result;
  }
  private clonedList() {
    let clonedList = [...this.state.internalPolicies];
    let list = [];
    clonedList.forEach(policy => {
      list.push({
        Id: policy.Id,
        Name: policy.Name,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory,
        RegulatoryTopic: policy.RegulatoryTopic
      });
    });
    return list;
  }

  //dropdowns
  private dropDownPolicyCategory(items?): void {
    const filedName = "PolicyCategory";
    const policyCategoryDropDown = this.fillDropDown(items, filedName);
    this.setState({ policyCategoryDropDown });
  }

  private dropDownRegulatoryTopic(items?): void {
    const filedName = "RegulatoryTopic";
    const regulatoryTopicDropDown = this.fillDropDown(items, filedName);
    this.setState({ regulatoryTopicDropDown });
  }

  private dropDownMonth(items?): void {
    const filedName = "Month";
    const monthDropDown = this.fillDropDown(items, filedName);
    this.setState({ monthDropDown });
  }

  private fillDropDown(items?, filedName?) {
    let documentFiles = [];
    items
      ? documentFiles.push(...items)
      : documentFiles.push(...this.state.documentFiles);
    this.setState({ documentFiles });
    let listBeforeSplit = [];
    let listNoUnique = [];
    documentFiles.forEach(item => {
      if (item[filedName]) {
        listBeforeSplit.push({
          text: item[filedName]
        });
      }
    });
    listBeforeSplit.forEach(element => {
      element.text.split(";").forEach(split => {
        listNoUnique.push(split);
      });
    });

    let uniqueItems = new Set(listNoUnique.map(unique => unique));
    let dropDowResult = [];
    uniqueItems.forEach(uniqueItem => {
      dropDowResult.push({ key: uniqueItem, text: uniqueItem });
    });
    return dropDowResult;
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
