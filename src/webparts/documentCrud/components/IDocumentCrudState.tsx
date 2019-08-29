import { IListItem } from "./IListItem";
export interface IDocumentCrudState {
  status: string;
  documentFiles: IListItem[];
  joinPolicyCategoryItems: IListItem[];
  policyCategoryDropDown: object[];
  stringPolicyCategory: string[];
}
