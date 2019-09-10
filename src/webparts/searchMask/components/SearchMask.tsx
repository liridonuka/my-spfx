import * as React from "react";
import styles from "./SearchMask.module.scss";
import { ISearchMaskProps } from "./ISearchMaskProps";
import { escape, times } from "@microsoft/sp-lodash-subset";
import { ISearchMaskState } from "./ISearchMaskState";
import { Web } from "@pnp/sp";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { Rating, RatingSize } from "office-ui-fabric-react/lib/Rating";
export default class DocumentCrud extends React.Component<
  ISearchMaskProps,
  ISearchMaskState
> {
  // public state = {
  //   status: this.listNotConfigured(this.props)
  //     ? "Please configure list in Web Part properties"
  //     : "Ready",
  //   documentFiles: [],
  //   policyCategories: [],
  //   policyCategoryDropDown: []
  // };
  constructor(props: ISearchMaskProps, state: ISearchMaskState) {
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
      joinYearItems: [],
      yearDropDown: [],
      stringYear: [],
      joinMonthItems: [],
      monthDropDown: [],
      stringMonth: [],
      anyPolicyCategorySelected: false,
      anyRegulatoryTopicSelected: false,
      anyYearSelected: false,
      anyMonthSelected: false
    };
  }
  public render(): React.ReactElement<ISearchMaskProps> {
    return (
      <div className={styles.searchMask}>
        <div>
          <div>
            <div className={styles.droping}>
              <div className={styles.dropOne}>
                <Dropdown
                  placeHolder="Filter by business function"
                  onChanged={this.filterByPolicyCategory.bind(this)}
                  multiSelect
                  options={this.state.policyCategoryDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
              <div className={styles.dropTwo}>
                <Dropdown
                  placeHolder="Filter by regulatory topic"
                  onChanged={this.filterByRegulatoryTopic.bind(this)}
                  multiSelect
                  options={this.state.regulatoryTopicDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
            </div>
            <div className={styles.droping}>
              <div className={styles.dropOne}>
                <Dropdown
                  placeHolder="Filter by approval year"
                  onChanged={this.filterByYear.bind(this)}
                  multiSelect
                  options={this.state.yearDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
              <div className={styles.dropTwo}>
                <Dropdown
                  placeHolder="Filter by approval month"
                  onChanged={this.filterByMonth.bind(this)}
                  multiSelect
                  options={this.state.monthDropDown}
                  // title={this.state.titleCategory}
                />
              </div>
            </div>
          </div>
          <div>
            <div className={styles.statusDivStyle}>{this.state.status}</div>

            {this.state.documentFiles.map(document => (
              <div className={styles.rowDivStyle}>
                <Icon iconName="AddBookmark" title="Add to bookmark" />
                <Icon
                  iconName="EntryView"
                  title="Policy details"
                  style={{ color: "#8a2c49" }}
                />
                &nbsp;
                <Icon
                  title="New policy"
                  iconName={
                    parseFloat(document.Version) <= 1
                      ? document.NewDocumentExpired < 7
                        ? "glimmer"
                        : undefined
                      : undefined
                  }
                  style={{ color: "#c4c91a" }}
                />
                {document.PolicyNumber} {document.Name} {" v"}
                {document.Version}
                <div>
                  <div style={{ display: "inline-block" }}>
                    <Rating min={0} max={5} />
                  </div>
                  &nbsp;
                  <div
                    title="Approved date"
                    style={{ display: "inline-block", verticalAlign: "middle" }}
                  >
                    <Icon iconName="Calendar" style={{ color: "#8a2c49" }} />
                    &nbsp;
                    <div
                      style={{
                        display: "inline-block",
                        verticalAlign: "middle",
                        paddingBottom: 7,
                        fontSize: 12
                      }}
                    >
                      {document.ApprovedDate}
                    </div>
                  </div>
                  &nbsp;
                  <div
                    style={{
                      display: "inline-block",
                      verticalAlign: "middle",
                      paddingBottom: 7,
                      fontSize: 12
                    }}
                  >
                    {/* <Icon
                      iconName={
                        parseFloat(document.Version) <= 1 ? "glimmer" : ""
                      }
                    /> */}
                  </div>
                </div>
              </div>
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
          console.log(items);
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
        PolicyNumber: policy.Policy_x0020_Number,
        Name: policy.File.Name,
        Version: policy.OData__UIVersionString,
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
      const diffDays = this.getDiffDays(policy.ApprovedDate);
      documentFiles.push({
        Id: policy.Id,
        PolicyNumber: policy.PolicyNumber,
        Name: policy.Name,
        Version: policy.Version,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory
          ? policy.PolicyCategory.split("undefined;").pop()
          : "",
        RegulatoryTopic: policy.RegulatoryTopic
          ? policy.RegulatoryTopic.split("undefined;").pop()
          : "",
        Year: new Date(policy.ApprovedDate).getFullYear().toString(),
        Month: monthNames[new Date(policy.ApprovedDate).getMonth()],
        NewDocumentExpired: policy.ApprovedDate ? diffDays : undefined
      });
    });
    this.setState({
      documentFiles,
      internalPolicies: documentFiles,
      joinPolicyCategoryItems,
      joinRegulatoryTopicItems,
      status: `Successfully loaded ${documentFiles.length} items`
    });
    this.dropDownPolicyCategory(documentFiles);
    this.dropDownRegulatoryTopic();
    this.dropDownDateYear();
    this.dropDownDateMonth();
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
  private dropDownDateYear(items?): void {
    const filedName = "Year";
    const yearDropDown = this.fillDropDown(items, filedName);
    this.setState({ yearDropDown });
  }
  private dropDownDateMonth(items?): void {
    const filedName = "Month";
    const monthDropDown = this.fillDropDown(items, filedName);
    this.setState({ monthDropDown });
  }
  private fillDropDown(items?, filedName?) {
    let documentFiles = [];
    items
      ? documentFiles.push(...items)
      : documentFiles.push(...this.state.documentFiles);
    let listBeforeSplit = [];
    let listNoUnique = [];
    documentFiles.forEach(item => {
      if (item[filedName]) {
        listBeforeSplit.push({
          text: ";" + item[filedName]
        });
      }
    });

    listBeforeSplit.forEach(element => {
      element.text.split(";").forEach(split => {
        if (split) {
          listNoUnique.push(split);
        }
      });
    });
    let uniqueItems = new Set(listNoUnique.map(unique => unique));
    let dropDowResult = [];
    uniqueItems.forEach(uniqueItem => {
      dropDowResult.push({ key: uniqueItem, text: uniqueItem });
    });
    return dropDowResult;
  }
  private filterByPolicyCategory(selectedItems) {
    const stringPolicyCategory = this.selectItems(
      selectedItems,
      this.state.stringPolicyCategory
    );
    this.setState({
      stringPolicyCategory
    });
    const clonedList = this.clonedList(
      this.state.anyRegulatoryTopicSelected ||
        this.state.anyYearSelected ||
        this.state.anyMonthSelected
        ? stringPolicyCategory.length > 0
          ? this.state.joinPolicyCategoryItems
          : this.state.joinPolicyCategoryItems
        : this.state.internalPolicies
    );
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
        stringPolicyCategory.length > 0 ? documentFiles : clonedList,
      joinRegulatoryTopicItems:
        stringPolicyCategory.length > 0 ? documentFiles : clonedList,
      joinYearItems:
        stringPolicyCategory.length > 0 ? documentFiles : clonedList,
      joinMonthItems:
        stringPolicyCategory.length > 0 ? documentFiles : clonedList,
      anyPolicyCategorySelected: stringPolicyCategory.length > 0 ? true : false
    });
    if (!this.state.anyRegulatoryTopicSelected) {
      this.dropDownRegulatoryTopic(
        stringPolicyCategory.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyYearSelected) {
      this.dropDownDateYear(
        stringPolicyCategory.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyMonthSelected) {
      this.dropDownDateMonth(
        stringPolicyCategory.length > 0 ? documentFiles : clonedList
      );
    }
  }
  private filterByRegulatoryTopic(selectedItems) {
    const stringRegulatoryTopic = this.selectItems(
      selectedItems,
      this.state.stringRegulatoryTopic
    );
    this.setState({
      stringRegulatoryTopic
    });
    const clonedList = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyYearSelected ||
        this.state.anyMonthSelected
        ? stringRegulatoryTopic.length > 0
          ? this.state.joinRegulatoryTopicItems
          : this.state.joinRegulatoryTopicItems
        : this.state.internalPolicies
    );

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
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
      joinPolicyCategoryItems:
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
      joinYearItems:
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
      joinMonthItems:
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList,
      anyRegulatoryTopicSelected:
        stringRegulatoryTopic.length > 0 ? true : false
    });

    if (!this.state.anyPolicyCategorySelected) {
      this.dropDownPolicyCategory(
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyYearSelected) {
      this.dropDownDateYear(
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyMonthSelected) {
      this.dropDownDateMonth(
        stringRegulatoryTopic.length > 0 ? documentFiles : clonedList
      );
    }
  }
  private filterByYear(selectedItems) {
    const stringYear = this.selectItems(selectedItems, this.state.stringYear);
    this.setState({
      stringYear
    });
    const clonedList = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyRegulatoryTopicSelected ||
        this.state.anyMonthSelected
        ? stringYear.length > 0
          ? this.state.joinYearItems
          : this.state.joinYearItems
        : this.state.internalPolicies
    );

    let filteredList = [];
    stringYear.forEach(s => {
      clonedList
        .filter(f => f.Year.includes(s))
        .map(item => filteredList.push(item));
    });
    let uniqueItems = new Set(filteredList.map(unique => unique));
    let documentFiles = [];
    uniqueItems.forEach(u => {
      documentFiles.push(u);
    });
    this.setState({
      documentFiles: stringYear.length > 0 ? documentFiles : clonedList,
      joinPolicyCategoryItems:
        stringYear.length > 0 ? documentFiles : clonedList,
      joinRegulatoryTopicItems:
        stringYear.length > 0 ? documentFiles : clonedList,
      joinMonthItems: stringYear.length > 0 ? documentFiles : clonedList,
      anyYearSelected: stringYear.length > 0 ? true : false
    });

    if (!this.state.anyPolicyCategorySelected) {
      this.dropDownPolicyCategory(
        stringYear.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyRegulatoryTopicSelected) {
      this.dropDownRegulatoryTopic(
        stringYear.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyMonthSelected) {
      this.dropDownDateMonth(
        stringYear.length > 0 ? documentFiles : clonedList
      );
    }
  }
  private filterByMonth(selectedItems) {
    const stringMonth = this.selectItems(selectedItems, this.state.stringMonth);
    this.setState({
      stringMonth
    });
    const clonedList = this.clonedList(
      this.state.anyPolicyCategorySelected ||
        this.state.anyRegulatoryTopicSelected ||
        this.state.anyYearSelected
        ? stringMonth.length > 0
          ? this.state.joinMonthItems
          : this.state.joinMonthItems
        : this.state.internalPolicies
    );
    let filteredList = [];
    stringMonth.forEach(s => {
      clonedList
        .filter(f => f.Month.includes(s))
        .map(item => filteredList.push(item));
    });
    let uniqueItems = new Set(filteredList.map(unique => unique));
    let documentFiles = [];
    uniqueItems.forEach(u => {
      documentFiles.push(u);
    });
    this.setState({
      documentFiles: stringMonth.length > 0 ? documentFiles : clonedList,
      joinPolicyCategoryItems:
        stringMonth.length > 0 ? documentFiles : clonedList,
      joinRegulatoryTopicItems:
        stringMonth.length > 0 ? documentFiles : clonedList,
      joinYearItems: stringMonth.length > 0 ? documentFiles : clonedList,
      anyMonthSelected: stringMonth.length > 0 ? true : false
    });

    if (!this.state.anyPolicyCategorySelected) {
      this.dropDownPolicyCategory(
        stringMonth.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyRegulatoryTopicSelected) {
      this.dropDownRegulatoryTopic(
        stringMonth.length > 0 ? documentFiles : clonedList
      );
    }
    if (!this.state.anyYearSelected) {
      this.dropDownDateYear(
        stringMonth.length > 0 ? documentFiles : clonedList
      );
    }
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
  private clonedList(sourceList) {
    let clonedList = [...sourceList];
    let list = [];
    clonedList.forEach(policy => {
      list.push({
        Id: policy.Id,
        PolicyNumber: policy.PolicyNumber,
        Name: policy.Name,
        Version: policy.Version,
        DocumentLink: policy.DocumentLink,
        ApprovedDate: new Date(policy.ApprovedDate).toLocaleDateString(),
        PolicyCategory: policy.PolicyCategory,
        RegulatoryTopic: policy.RegulatoryTopic,
        Year: policy.Year,
        Month: policy.Month,
        NewDocumentExpired: policy.NewDocumentExpired
      });
    });
    return list;
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
  private getDiffDays(elementDate) {
    const date1 = new Date(elementDate);
    const date2 = new Date();
    const diffTime = Math.abs(date2.getTime() - date1.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }
  private listNotConfigured(props: ISearchMaskProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
