import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICrudOperationProps {
  // Title of the SharePoint list that stores leave requests, set via the property pane.
  listName: string;
  context: WebPartContext;
}
