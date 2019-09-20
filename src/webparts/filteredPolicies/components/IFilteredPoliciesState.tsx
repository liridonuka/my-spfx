import { IListItem } from "./IListItem";
export interface IFilteredPoliciesState {
  status: string;
  statusIndicator: number;
  internalPolicies: IListItem[];
  documentFiles: IListItem[];
  joinPolicyCategoryItems: IListItem[];
  joinRegulatoryTopicItems: IListItem[];
}
