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
  public state = {
    status: this.listNotConfigured(this.props)
      ? "Please configure list in Web Part properties"
      : "Ready",
    documentFiles: [],
    policyCategories: [],
    policyCategoryDropDown: []
  };
  // constructor(props: IDocumentCrudProps, state: IDocumentCrudState) {
  //   super(props);

  //   this.state = {
  //     status: this.listNotConfigured(this.props)
  //       ? "Please configure list in Web Part properties"
  //       : "Ready",
  //     documents: []
  //   };
  // }

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
    let policyCategories = [];
    let conncatList = [];
    items.forEach(file => {
      documentFiles.push({
        Id: file.Id,
        Name: file.File.Name,
        DocumentLink: file.File.LinkingUrl,
        ApprovedDate: file.Modified
      });
      file.Policy_x0020_Category.forEach(metaData => {
        policyCategories.push({
          Id: file.Id,
          MetaData: metaData.Label.split(/:/)[1]
        });
      });
    });

    this.policyCategory(policyCategories);

    this.setState({
      documentFiles,
      policyCategories,
      status: `Successfully loaded ${documentFiles.length} items`
    });
  }

  private filteredFile(item: IDropdownOption): void {
    const newSelectedItems = [];
    if (item.selected) {
      // add the option if it's checked
      newSelectedItems.push(item.key as string);
    } else {
      // remove the option if it's unchecked
      const currIndex = newSelectedItems.indexOf(item.key as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
    }
    console.log(newSelectedItems);
    // this.setState({
    //   policyCategoryDropDown: newSelectedItems
    // });

    // console.log(newSelectedItems.length);
    // //finally filter files based on search fieds
    // let conncatList = [];
    // let filteredPolicies = [];
    // this.state.policyCategories.forEach(filtered => {
    //   if (filtered.Label.includes(parameter))
    //     filteredPolicies.push({
    //       Id: filtered.Id,
    //       MetaData: filtered.Label
    //     });
    // });
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

  //dropdown
  private policyCategory(policyCategory): void {
    let policyCategoryDropDown = [];
    policyCategory.forEach(item => {
      policyCategoryDropDown.push({ key: item.MetaData, text: item.MetaData });
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
