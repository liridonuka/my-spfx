import { IListItem, IPolicyUser } from "./IListItem";
export interface ISearchMaskState {
  status: string;
  statusIndicator: number;
  internalPolicies: IListItem[];
  documentFiles: IListItem[];
  policyUser: IPolicyUser[];
  joinPolicyCategoryItems: IListItem[];
  policyCategoryDropDown: object[];
  stringPolicyCategory: string[];
  joinRegulatoryTopicItems: IListItem[];
  regulatoryTopicDropDown: object[];
  stringRegulatoryTopic: string[];
  joinYearItems: IListItem[];
  yearDropDown: object[];
  stringYear: string[];
  joinMonthItems: IListItem[];
  monthDropDown: object[];
  stringMonth: string[];
  anyPolicyCategorySelected: boolean;
  anyRegulatoryTopicSelected: boolean;
  anyYearSelected: boolean;
  anyMonthSelected: boolean;
  hideDialog: boolean;
  commentState: string;
  policyNumber: number;
  showPanel: boolean;
  itemsLengthDisplayed: number;
}
