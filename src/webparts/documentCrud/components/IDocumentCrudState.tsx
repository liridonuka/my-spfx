import { IListItem } from "./IListItem";
export interface IDocumentCrudState {
  status: string;
  documentFile: IListItem[];
  metaDataFile: IListItem[];
}
