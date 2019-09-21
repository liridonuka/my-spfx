export interface IListItem {
  Id: number;
  PolicyNumber: string;
  Name: string;
  links: ISuper[];
  Version: string;
  DocumentLink: string;
  PolicyCategory: string;
  RegulatoryTopic: string;
  ApprovedDate: string;
  MonthDate: string;
  YearDate: string;
  NewDocumentExpired: number;
  Favorite: number;
  Rate: number;
  Comment: string;
}

export interface ISuper {
  key: string;
  name: string;
  url: string;
}
