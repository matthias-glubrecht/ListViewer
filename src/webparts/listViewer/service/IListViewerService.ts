import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IFieldInfo } from './IFieldInfo';
import { IViewDefinition } from './IViewDefinition';

export interface IListViewerService {
    GetViewDefinition: (viewId: string) => Promise<IViewDefinition>;
    GetViewsOfList: () => Promise<IDropdownOption[]>;
    GetListFields: () => Promise<IFieldInfo[]>;
    GetViewFields: (viewId: string) => Promise<IFieldInfo[]>;
    // tslint:disable-next-line:no-any
    GetListItems: (view: IViewDefinition) => Promise<any[]>;
    GetListTitle: () => Promise<string>;
    // Returns list items with FieldValuesAsHtml
    // tslint:disable-next-line:no-any
    GetListItemsAsHtmlAndText: (view: IViewDefinition) => Promise<any[]>;
}