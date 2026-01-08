import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

import * as strings from 'ListViewerWebPartStrings';
import ListViewer from './components/ListViewer';
import { IListViewerProps } from './components/IListViewerProps';
import { IListViewerService } from './service/IListViewerService';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import { ListViewerService } from './service/ListViewerService';
import Utility from './utility/utility';

export interface IListViewerWebPartProps {
  selectedList: string;
  selectedView: string;
  webPartTitle: string;
  noEntriesText?: string;
}

export default class ListViewerWebPart extends BaseClientSideWebPart<IListViewerWebPartProps> {
  private _service: IListViewerService;

  public render(): void {
    const element: React.ReactElement<IListViewerProps> = React.createElement(
      ListViewer,
      {
        service: this._service,
        listId: this.properties.selectedList,
        viewId: this.properties.selectedView,
        webPartTitle: this.properties.webPartTitle,
        noEntriesText: this.properties.noEntriesText
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._service = new ListViewerService(this.context);
    return super.onInit();
  }

  /// @ts-ignore
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: string | number,
    newValue: string | number
  ): void {
    if (propertyPath === 'selectedList') {
      this.properties.selectedView = undefined;
      if (newValue) {
        this._service.GetListTitle(newValue as string).then(this.setTitle);
      } else {
        this.setTitle('');
      }
    } else if (propertyPath === 'selectedView') {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: `Version ${this.context.manifest.version} - Einstellungen`
          },
          groups: [
            {
              groupName: strings.PropertyPaneGroupListAndView,
              groupFields: [
                PropertyFieldListPicker('selectedList', {
                  label: strings.PropertyPaneFieldListLabel,
                  selectedList: this.properties.selectedList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 200,
                  key: 'befehleLibraryFieldId'
                }),
                PropertyPaneAsyncDropdown('selectedView', {
                  label: strings.PropertyPaneFieldViewLabel,
                  loadOptions: this.loadViews,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  selectedKey: this.properties.selectedView,
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneButton('editView', {
                  text: strings.PropertyPaneButtonEditView,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: !this.properties.selectedList || !this.properties.selectedView,
                  onClick: this.editListView
                }),
                PropertyPaneTextField('noEntriesText', {
                  label: strings.PropertyPaneFieldNoEntriesLabel,
                  value: this.properties.noEntriesText,
                  placeholder: strings.PropertyPaneFieldNoEntriesPlaceholder,
                  description: strings.PropertyPaneFieldNoEntriesDescription
                }),
                PropertyPaneHorizontalRule()
              ]
            },
            {
              groupName: strings.PropertyPaneGroupWebPartTitle,
              groupFields: [
                PropertyPaneTextField('webPartTitle', {
                  label: strings.PropertyPaneFieldWebPartTitleLabel,
                  description: strings.PropertyPaneFieldWebPartTitleDescription,
                  value: this.properties.webPartTitle,
                  placeholder: strings.PropertyPaneFieldWebPartTitlePlaceholder
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private editListView = () => {
    const webUrl: string = Utility.trimEnd(this.context.pageContext.web.serverRelativeUrl, '/');
    const url: string =
      `${webUrl}/_layouts/15/ViewEdit.aspx?` +
      `View={${this.properties.selectedView.toUpperCase()}}&` +
      `List={${this.properties.selectedList.toUpperCase()}}`;
    window.open(encodeURI(url), '_blank');
  }

  private setTitle = (title: string): void => {
    this.properties.webPartTitle = title;
    if (this.context.propertyPane.isPropertyPaneOpen) {
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private loadViews = async (): Promise<IDropdownOption[]> => {
    return this._service.GetViewsOfList(this.properties.selectedList);
  }
}
