export interface IListItem {
  Id: number;
  PolicyNumber: string;
  Name: string;
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

export interface IPolicyUser {
  Title?: string;
  Policy: string;
  PolicyLink: string;
  PolicyNumber: number;
  Favorite: number;
  Rate: number;
  Comment: string;
  Id: number;
}
