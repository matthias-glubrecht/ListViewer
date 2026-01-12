import { IListViewerService } from '../../service/IListViewerService';

export interface IDetailsViewProps {
  isOpen: boolean;
  onDismiss: () => void;

  itemId: number;
  viewId: string;
  service: IListViewerService;
}
