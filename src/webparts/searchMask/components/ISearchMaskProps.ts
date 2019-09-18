import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISearchMaskProps {
  listName: string;
  joinListName: string;
  context: WebPartContext;
  commentDialogTitle: string;
  commentDialogSubTitle: string;
  myWorkingSpace: boolean;
}
