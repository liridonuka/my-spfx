import { IListItem } from "./IListItem";
export interface ISearchMaskState {
  status: string;
  internalPolicies: IListItem[];
  documentFiles: IListItem[];
  joinPolicyCategoryItems: IListItem[];
  policyCategoryDropDown: object[];
  stringPolicyCategory: string[];
  joinRegulatoryTopicItems: IListItem[];
  regulatoryTopicDropDown: object[];
  stringRegulatoryTopic: string[];
  joinYearItems: IListItem[];
  yearDropDown: object[];
  stringYear: string[];
  monthDropDown: object[];
  stringMonth: string[];
  anyPolicyCategorySelected: boolean;
  anyRegulatoryTopicSelected: boolean;
  anyYearSelected: boolean;
}
