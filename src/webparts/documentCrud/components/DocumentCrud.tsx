import * as React from "react";
import styles from "./DocumentCrud.module.scss";
import { IDocumentCrudProps } from "./IDocumentCrudProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp, Item, ItemAddResult, ItemUpdateResult } from "@pnp/sp";
import { IListItem } from "./IListItem";
import { IDocumentCrudState } from "./IDocumentCrudState";
import { Diving } from "./Diving";

export default class DocumentCrud extends React.Component<
  IDocumentCrudProps,
  IDocumentCrudState
> {
  public state = {
    status: this.listNotConfigured(this.props)
      ? "Please configure list in Web Part properties"
      : "Ready",
    documentFile: [],
    metaDataFile: []
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
        {this.state.documentFile.map(document => (
          <Diving
            key={document.Id}
            name={document.Name}
            id={document.Id}
            documentLink={document.DocumentLink}
          >
            {this.state.metaDataFile
              .filter(f => f.Id === document.Id)
              .map(metaData => metaData.FileMetaData)}
          </Diving>
        ))}
      </div>
    );
  }

  public componentWillMount() {
    this.readItems();
  }

  private getMyItems(items): void {
    let documentFiles = [];
    let documentMetaData = [];
    items.forEach(file => {
      documentFiles.push({
        Id: file.Id,
        Name: file.File.Name,
        DocumentLink: file.File.LinkingUrl
      });
      file.MyMetadata.forEach(metaData => {
        documentMetaData.push({
          Id: file.Id,
          FileMetaData: metaData.Label + ";"
        });
      });
      // console.log(element);
      // documentList.push({
      //   Id: element.Id,
      //   Name: element.File.Name,
      //   DocumentLink: element.File.LinkingUrl
      // });
    });
    // console.log(documentList);
    this.setState({
      documentFile: documentFiles,
      metaDataFile: documentMetaData,
      status: `Successfully loaded ${documentFiles.length} items`
    });
  }

  private readItems(): void {
    let documentList = [];
    this.setState({ documentFile: [], status: "Loading all items..." });
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.expand("File")
      .getAll()
      .then(
        items => {
          // console.log(items);
          this.getMyItems(items);
        },
        (error: any): void => {
          this.setState({
            status: "Loading all items failed with error: " + error
          });
        }
      );
  }

  private listNotConfigured(props: IDocumentCrudProps): boolean {
    return (
      props.listName === undefined ||
      props.listName === null ||
      props.listName.length === 0
    );
  }
}
