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
    documents: []
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
        {this.state.documents.map(document => (
          <Diving
            key={document.Id}
            name={document.Name}
            id={document.Id}
            documentLink={document.DocumentLink}
          />
        ))}
      </div>
    );
  }

  public componentWillMount() {
    this.readItems();
  }

  private getMyItems(items): void {
    let documentList = [];
    items.forEach(element => {
      documentList.push({
        Id: element.Id,
        Name: element.File.Name,
        DocumentLink: element.File.LinkingUrl
      });
    });
    console.log(documentList);
    this.setState({
      documents: documentList,
      status: `Successfully loaded ${documentList.length} items`
    });
  }

  private readItems(): void {
    let documentList = [];
    this.setState({ documents: [], status: "Loading all items..." });
    sp.web.lists
      .getByTitle(this.props.listName)
      .items.expand("File")
      .getAll()
      .then(
        itema => {
          this.getMyItems(itema);
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
