import { IListItem } from "./IListItem";
export interface IDocumentCrudState {
  status: string;
  internalPolicies: IListItem[];
  documentFiles: IListItem[];
  joinPolicyCategoryItems: IListItem[];
  policyCategoryDropDown: object[];
  stringPolicyCategory: string[];
  joinRegulatoryTopicItems: IListItem[];
  regulatoryTopicDropDown: object[];
  stringRegulatoryTopic: string[];
  monthDropDown: object[];
  stringMonth: string[];
}
