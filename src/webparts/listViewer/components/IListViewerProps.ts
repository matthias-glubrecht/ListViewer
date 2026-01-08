import { IListViewerService } from '../service/IListViewerService';

export interface IListViewerProps {
  service: IListViewerService;
  listId: string;
  viewId: string;
  webPartTitle?: string;
  noEntriesText?: string;
}
