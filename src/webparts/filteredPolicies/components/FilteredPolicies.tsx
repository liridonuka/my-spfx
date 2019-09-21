import * as React from "react";
import styles from "./FilteredPolicies.module.scss";
import { IFilteredPoliciesProps } from "./IFilteredPoliciesProps";
import { IFilteredPoliciesState } from "./IFilteredPoliciesState";
import { Web, ItemAddResult, sp } from "@pnp/sp";
import { Nav, INavLinkGroup } from "office-ui-fabric-react/lib/Nav";
import { escape } from "@microsoft/sp-lodash-subset";

export default class FilteredPolicies extends React.Component<
  IFilteredPoliciesProps,
  IFilteredPoliciesState
> {
  constructor(props: IFilteredPoliciesProps, state: IFilteredPoliciesState) {
    super(props);

    this.state = {
      status: this.listNotConfigured(this.props)
        ? "Please configure list in Web Part properties"
        : "Ready",
      statusIndicator: 1,
      internalPolicies: [],
      documentFiles: [],
      joinPolicyCategoryItems: [],
      joinRegulatoryTopicItems: []
    };
  }
  public render(): React.ReactElement<IFilteredPoliciesProps> {
    return (
      <div className={styles.filteredPolicies}>
        {this.state.joinPolicyCategoryItems.map(items => (
          <Nav
            // styles={{ root: { width: 300 } }}
            expandButtonAriaLabel="Expand or collapse"
            ariaLabel="Nav example similiar to one found in this demo page"
            groups={[
              {
                name: `${items.PolicyCategory}`,
                collapseByDefault: true,
                links: [
                  {
                    key: `${items.links.map(i => i.key)}`,
                    name: `${items.links.map(i => i.name)}`,
                    url: `${items.links.map(i => i.url)}`
                  }
                ]
              }
            ]}
          />
        ))}
      </div>
    );
  }
  public componentWillMount() {
    this.connectAndReadPolicies();
  }
  private connectAndReadPolicies(): void {
    this.setState({
      documentFiles: [],
      status: "Loading all items...",
      statusIndicator: 0
    });
    const web = new Web(this.props.context.pageContext.web.absoluteUrl);
    web.lists
      .getByTitle(this.props.listName)
      .items.top(900)
      .select(
        "File/Name",
        "Id",
        "Policy_x0020_Number",
        "OData__UIVersionString",
        "ServerRedirectedEmbedUrl",
        "Date_x0020_of_x0020_approval",
        "Policy_x0020_Category",
        "Regulatory_x0020_Topic"
      )
      .expand("File")
      .orderBy("Date_x0020_of_x0020_approval", false)
      .get()
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
        DocumentLink: policy.ServerRedirectedEmbedUrl,
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
        PolicyNumber: policy.PolicyNumber,
        Name: policy.Name,
        Version: policy.Version,
        DocumentLink: policy.DocumentLink,
        ApprovedDate:
          new Date(policy.ApprovedDate).getFullYear() !== 1970
            ? new Date(policy.ApprovedDate).toLocaleDateString()
            : "No approved date",
        PolicyCategory: policy.PolicyCategory
          ? policy.PolicyCategory.split("undefined;").pop()
          : "",
        RegulatoryTopic: policy.RegulatoryTopic
          ? policy.RegulatoryTopic.split("undefined;").pop()
          : ""
      });
    });
    let myList = [];
    const a = this.fillDropDown(joinPolicyCategoryItems, "PolicyCategory");
    a.forEach((value, i) => {
      myList.push({ PolicyCategory: value, links: [] });
      documentFiles
        .filter(f => f.PolicyCategory.includes(value))
        .map(item =>
          myList[i].links.push({
            key: item.Name,
            name: item.Name,
            url: item.DocumentLink
          })
        );
    });

    console.log(myList);
    this.setState({
      documentFiles,
      internalPolicies: documentFiles,
      joinPolicyCategoryItems: myList,
      joinRegulatoryTopicItems,
      status: `Displaying ${documentFiles.length} items`,
      statusIndicator: 1
    });
  }
  private fillDropDown(items?, filedName?) {
    let documentFiles = [];
    documentFiles.push(...items);

    let listBeforeSplit = [];
    let listNoUnique = [];
    documentFiles.forEach(item => {
      if (item.Year !== "1970") {
        if (item[filedName]) {
          listBeforeSplit.push({
            text: ";" + item[filedName]
          });
        }
      }
    });

    listBeforeSplit.forEach(element => {
      element.text.split(";").forEach(split => {
        if (split && split !== "undefined") {
          listNoUnique.push(split);
        }
      });
    });
    let uniqueItems = new Set(listNoUnique.map(unique => unique));
    let dropDowResult = [];
    uniqueItems.forEach(uniqueItem => {
      dropDowResult.push(uniqueItem);
    });
    return dropDowResult;
  }
  private listNotConfigured(props: IFilteredPoliciesProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
