import { IListItem } from "./IListItem";
export interface IDocumentCrudState {
  status: string;
  documents: IListItem[];
}
