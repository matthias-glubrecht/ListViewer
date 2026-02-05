import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

import * as strings from 'ListViewerWebPartStrings';
import { IListViewerProps } from './components/ListViewer/IListViewerProps';
import { IListViewerService } from './service/IListViewerService';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import { ListViewerService } from './service/ListViewerService';
import ListViewer from './components/ListViewer/ListViewer';
import Utility from './utility/Utility';

export interface IListViewerWebPartProps {
  selectedList: string;
  selectedView: string;
  detailsView: string;
  webPartTitle: string;
  noEntriesText?: string;
  showBodyCaption: boolean;
}

export default class ListViewerWebPart extends BaseClientSideWebPart<IListViewerWebPartProps> {
  private _service: IListViewerService;

  public render(): void {
    const element: React.ReactElement<IListViewerProps> = React.createElement(
      ListViewer,
      {
        service: this._service,
        viewId: this.properties.selectedView,
        detailsViewId: this.properties.detailsView,
        webPartTitle: this.properties.webPartTitle,
        noEntriesText: this.properties.noEntriesText,
        showBodyCaptionInDetails: this.properties.showBodyCaption
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._service = new ListViewerService(this.context, this.properties.selectedList);
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
      this.properties.detailsView = undefined;
      if (newValue) {
        this._service = new ListViewerService(this.context, newValue as string);
        this._service.GetListTitle().then(this.setTitle);
      } else {
        this.setTitle('');
      }
    } else if (propertyPath === 'selectedView' || propertyPath === 'detailsView') {
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
                PropertyPaneButton('editListView', {
                  text: strings.PropertyPaneButtonEditView,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: !this.properties.selectedList || !this.properties.selectedView,
                  onClick: this.editListView
                }),
                PropertyPaneAsyncDropdown('detailsView', {
                  label: strings.PropertyPaneDetailsViewLabel,
                  loadOptions: this.loadViews,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  selectedKey: this.properties.detailsView,
                  disabled: !this.properties.selectedList
                }),
                PropertyPaneButton('editDetailsView', {
                  text: strings.PropertyPaneButtonEditView,
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: !this.properties.selectedList || !this.properties.detailsView,
                  onClick: this.editDetailsView
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
              groupName: strings.PropertyPaneGroupLabels,
              groupFields: [
                PropertyPaneTextField('webPartTitle', {
                  label: strings.PropertyPaneFieldWebPartTitleLabel,
                  description: strings.PropertyPaneFieldWebPartTitleDescription,
                  value: this.properties.webPartTitle,
                  placeholder: strings.PropertyPaneFieldWebPartTitlePlaceholder
                }),
                PropertyPaneToggle('showBodyCaption', {
                  label: strings.PropertyPaneFieldShowBodyCaptionLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private editView = (viewId: string) => {
    const webUrl: string = Utility.trimEnd(this.context.pageContext.web.serverRelativeUrl, '/');
    const url: string =
      `${webUrl}/_layouts/15/ViewEdit.aspx?` +
      `View={${viewId.toUpperCase()}}&` +
      `List={${this.properties.selectedList.toUpperCase()}}`;
    window.open(encodeURI(url), '_blank');
  }

  private editListView = () => {
    this.editView(this.properties.selectedView);
  }

  private editDetailsView = () => {
    this.editView(this.properties.detailsView);
  }

  private setTitle = (title: string): void => {
    this.properties.webPartTitle = title;
    if (this.context.propertyPane.isPropertyPaneOpen) {
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private loadViews = async (): Promise<IDropdownOption[]> => {
    return this._service.GetViewsOfList();
  }
}
