// tslint:disable:no-any
import { List, sp } from '@pnp/sp';
import { IListViewerService } from './IListViewerService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IFieldInfo } from './IFieldInfo';
import { IViewDefinition } from './IViewDefinition';

export class ListViewerService implements IListViewerService {
  private _listId: string;

  private _viewsPromise: Promise<IDropdownOption[]>;
  private _listFieldsPromise: Promise<IFieldInfo[]>;
  private _listTitlePromise: Promise<string>;
  private _enableAttachmentsPromise: Promise<boolean>;
  private _viewFieldsPromises: { [viewId: string]: Promise<IFieldInfo[]> } = {};
  private _viewDefinitionPromises: { [viewId: string]: Promise<IViewDefinition> } = {};

  constructor(context: WebPartContext, listId: string) {
    sp.setup({
      spfxContext: context
    });

    this._listId = listId;
  }

  public async GetViewDefinition(viewId: string): Promise<IViewDefinition> {
    if (!this._viewDefinitionPromises[viewId]) {
      this._viewDefinitionPromises[viewId] = (async (): Promise<IViewDefinition> => {
        try {
          const view: IViewDefinition = await sp.web.lists
            .getById(this._listId)
            .views.getById(viewId)
            .select('ViewQuery', 'ViewFields', 'ServerRelativeUrl')
            .expand('ViewFields')
            .get();
          view.ViewFields.Items = view.ViewFields.Items.map((fName) => {
            if (fName.startsWith('LinkTitle')) {
              return 'Title';
            } else {
              return fName;
            }});
          return view;
        } catch (error) {
          delete this._viewDefinitionPromises[viewId];
          console.error('Error retrieving view:', error);
          throw error;
        }
      })();
    }

    return this._viewDefinitionPromises[viewId];
  }

  // tslint:disable-next-line:no-any
  public async GetListItems(view: IViewDefinition): Promise<any[]> {
    try {
      const list: any = sp.web.lists.getById(this._listId);
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

  public async GetViewsOfList(): Promise<IDropdownOption[]> {
    if (!this._listId) {
      return [];
    }

    if (!this._viewsPromise) {
      this._viewsPromise = (async (): Promise<IDropdownOption[]> => {
        try {
          // tslint:disable-next-line:no-any
          const views: any[] = await sp.web.lists
            .getById(this._listId)
            .views.filter('Hidden eq false')
            .get();
          return views.map(v => {
            return {
              key: v.Id,
              text: v.Title
            };
          });
        } catch (error) {
          this._viewsPromise = undefined;
          throw error;
        }
      })();
    }

    return this._viewsPromise;
  }

  public async GetListFields(): Promise<IFieldInfo[]> {
    if (!this._listFieldsPromise) {
      this._listFieldsPromise = (async (): Promise<IFieldInfo[]> => {
        try {
          // tslint:disable-next-line:no-any
          const fields: any[] = await sp.web.lists
            .getById(this._listId)
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
        } catch (error) {
          this._listFieldsPromise = undefined;
          throw error;
        }
      })();
    }

    return this._listFieldsPromise;
  }

  public async GetViewFields(viewId: string): Promise<IFieldInfo[]> {
    if (!this._viewFieldsPromises[viewId]) {
      this._viewFieldsPromises[viewId] = (async (): Promise<IFieldInfo[]> => {
        try {
          // tslint:disable-next-line:no-any
          const o: any = await sp.web.lists
            .getById(this._listId)
            .getView(viewId)
            .fields.get();
          return o.Items.map((f: string) => {
            return { Title: f,
              InternalName: f.startsWith('LinkTitle') ? 'Title' : f,
              Type: '' };
          });
        } catch (error) {
          delete this._viewFieldsPromises[viewId];
          throw error;
        }
      })();
    }

    return this._viewFieldsPromises[viewId];
  }

  public async GetListTitle(): Promise<string> {
    if (!this._listTitlePromise) {
      this._listTitlePromise = (async (): Promise<string> => {
        try {
          const o: { Title: string } = await sp.web.lists
            .getById(this._listId)
            .select('Title')
            .get();
          if (o) {
            return o.Title;
          }
        } catch (error) {
          this._listTitlePromise = undefined;
          throw error;
        }
      })();
    }

    return this._listTitlePromise;
  }

  public async GetEnableAttachments(): Promise<boolean> {
    if (!this._enableAttachmentsPromise) {
      this._enableAttachmentsPromise = (async (): Promise<boolean> => {
        try {
          const o: { EnableAttachments: boolean } = await sp.web.lists
            .getById(this._listId)
            .select('EnableAttachments')
            .get();
          return o ? o.EnableAttachments : false;
        } catch (error) {
          this._enableAttachmentsPromise = undefined;
          throw error;
        }
      })();
    }

    return this._enableAttachmentsPromise;
  }

  // Returns list items with FieldValuesAsHtml, FieldValuesAsText and AttachmentFiles, if requested
  // tslint:disable-next-line:no-any
  public async GetListItemForDetailView(view: IViewDefinition): Promise<any[]> {
    try {
      const list: List = sp.web.lists.getById(this._listId);
      const expands: string[] = ['FieldValuesAsHtml', 'FieldValuesAsText'];
      if (view.ViewFields.Items.indexOf('Attachments') !== -1) {
        expands.push('AttachmentFiles');
      }
      const items: any[] = await list.getItemsByCAMLQuery(
        {
          ViewXml:
            `<View><Query>${view.ViewQuery}</Query><ViewFields>${view.ViewFields.Items.map(
              (field: string) => `<FieldRef Name='${field}' />`
            ).join('')}</ViewFields></View>`
        },
        ...expands
      );
      return items;
    } catch (error) {
      console.error('Error getting list items as HTML:', error);
      throw error;
    }
  }
}