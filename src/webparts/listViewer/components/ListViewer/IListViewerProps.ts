import { IListViewerService } from '../../service/IListViewerService';

export interface IListViewerProps {
  service: IListViewerService;
  viewId: string;
  detailsViewId: string;
  webPartTitle?: string;
  noEntriesText?: string;
  showBodyCaptionInDetails: boolean;
}
