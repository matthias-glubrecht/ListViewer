import { IFieldInfo } from '../../service/IFieldInfo';

export interface IListViewerState {
    loading: boolean;
    // tslint:disable-next-line:no-any
    items: any[];
    viewFields: IFieldInfo[];
    viewUrl?: string;
    configMissing: boolean;
    idForDialog: number;
}