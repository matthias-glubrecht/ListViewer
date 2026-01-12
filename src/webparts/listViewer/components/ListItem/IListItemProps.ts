import { IFieldInfo } from '../../service/IFieldInfo';

export interface IListItemProps {
    // tslint:disable-next-line:no-any
    item: any;
    fields: IFieldInfo[];
    showDetailsDialog: (itemId: number) => void;
}