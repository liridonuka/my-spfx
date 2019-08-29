import { IListItem } from "./IListItem";
export interface IDocumentCrudState {
  status: string;
  documentFiles: IListItem[];
  policyCategories: IListItem[];
  policyCategoryDropDown: object[];
}
