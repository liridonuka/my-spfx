declare interface ISearchMaskWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionLibrary: string;
  DescriptionJoinList: string;
  DescriptionMyWorkingSpace: string;
  CommentDialogTitle: string;
  CommentDialogSubTitle: string;
}

declare module "SearchMaskWebPartStrings" {
  const strings: ISearchMaskWebPartStrings;
  export = strings;
}
