import { IListItem } from "./IListItem";
export interface IFilteredPoliciesState {
  status: string;
  statusIndicator: number;
  documentFiles: IListItem[];
  navCategoryItems: IListItem[];
}
