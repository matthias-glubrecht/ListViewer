import { IListViewerService } from '../../service/IListViewerService';

export interface IDetailsViewProps {
  isOpen: boolean;
  onDismiss: () => void;
  showBodyCaption: boolean;
  itemId: number;
  viewId: string;
  service: IListViewerService;
}
