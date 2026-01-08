import { sp } from '@pnp/sp';
import { IListViewerService } from './IListViewerService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IFieldInfo } from './IFieldInfo';
import { IViewDefinition } from './IViewDefinition';

export class ListViewerService implements IListViewerService {
  constructor(context: WebPartContext) {
    sp.setup({
      spfxContext: context
    });
  }

  public async GetViewDefinition(listId: string, viewId: string): Promise<IViewDefinition> {
    try {
      const view: IViewDefinition = await sp.web.lists
        .getById(listId)
        .views.getById(viewId)
        .select('ViewQuery', 'ViewFields', 'ServerRelativeUrl')
        .expand('ViewFields')
        .get();
      return view;
    } catch (error) {
      console.error('Error retrieving view:', error);
      throw error;
    }
  }

  // tslint:disable-next-line:no-any
  public async GetListItems(listId: string, view: IViewDefinition): Promise<any[]> {
    try {
      // tslint:disable-next-line:no-any
      const list: any = sp.web.lists.getById(listId);
      // tslint:disable-next-line:no-any
      const items: any[] = await list.getItemsByCAMLQuery(
        {
          ViewXml:
            `<View><Query>${view.ViewQuery}</Query><ViewFields>${view.ViewFields.Items.map(
              (field: string) => `<FieldRef Name='${field}' />`
            ).join('')}</ViewFields></View>`
        },
        'FieldValuesAsText'
      );
      return items;
    } catch (error) {
      console.error('Error getting list items:', error);
      throw error;
    }
  }

  public async GetViewsOfList(listId: string): Promise<IDropdownOption[]> {
    if (listId) {
      // tslint:disable-next-line:no-any
      const views: any[] = await sp.web.lists
        .getById(listId)
        .views.filter('Hidden eq false')
        .get();
      return views.map(v => {
        return {
          key: v.Id,
          text: v.Title
        };
      });
    } else {
      return [];
    }
  }

  public async GetListFields(listId: string): Promise<IFieldInfo[]> {
    // tslint:disable-next-line:no-any
    const fields: any[] = await sp.web.lists
      .getById(listId)
      .fields.filter('Hidden eq false')
      .select('Title', 'InternalName', 'TypeAsString')
      .get();
    return fields.map(f => {
      return {
        InternalName: f.InternalName,
        Title: f.Title,
        Type: f.TypeAsString
      };
    });
  }

  public async GetViewFields(libraryId: string, viewId: string): Promise<IFieldInfo[]> {
    // tslint:disable-next-line:no-any
    const o: any = await sp.web.lists
      .getById(libraryId)
      .getView(viewId)
      .fields.get();
    return o.Items.map((f: string) => {
      return { Title: f, InternalName: f, Type: '' };
    });
  }

  public async GetListTitle(listId: string): Promise<string> {
    const o: { Title: string } = await sp.web.lists
      .getById(listId)
      .select('Title')
      .get();
    if (o) {
      return o.Title;
    }
  }
}