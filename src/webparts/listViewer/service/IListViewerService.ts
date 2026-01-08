import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IFieldInfo } from './IFieldInfo';
import { IViewDefinition } from './IViewDefinition';

export interface IListViewerService {
    GetViewDefinition: (listId: string, viewId: string) => Promise<IViewDefinition>;
    GetViewsOfList: (libraryId: string) => Promise<IDropdownOption[]>;
    GetListFields: (libraryId: string) => Promise<IFieldInfo[]>;
    GetViewFields: (libraryId: string, viewId: string) => Promise<IFieldInfo[]>;
    // tslint:disable-next-line:no-any
    GetListItems: (listId: string, view: IViewDefinition) => Promise<any[]>;
    GetListTitle: (listId: string) => Promise<string>;
}