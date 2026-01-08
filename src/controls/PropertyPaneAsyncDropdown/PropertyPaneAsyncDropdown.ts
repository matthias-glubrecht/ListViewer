
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IPropertyPaneAsyncDropdownProps } from './IPropertyPaneAsyncDropdownProps';
import { IPropertyPaneAsyncDropdownInternalProps } from './IPropertyPaneAsyncDropdownInternalProps';
import AsyncDropdown from './components/AsyncDropdown';
import { IAsyncDropdownProps } from './components/IAsyncDropdownProps';

class PropertyPaneAsyncDropdownControl implements IPropertyPaneField<IPropertyPaneAsyncDropdownProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public properties: IPropertyPaneAsyncDropdownInternalProps;
  private elem: HTMLElement;

  constructor(public targetProperty: string, properties: IPropertyPaneAsyncDropdownProps) {
    this.properties = {
      ...properties,
      key: properties.label,
      onRender: this.onRender.bind(this)
    };
  }

  // tslint:disable-next-line:max-line-length no-any
  private onRender(elem: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<IAsyncDropdownProps> = React.createElement(AsyncDropdown, {
      label: this.properties.label,
      loadOptions: this.properties.loadOptions,
      onChanged: this.onChanged.bind(this),
      callback: (key: string | number) => {
        if (changeCallback) {
          changeCallback(this.targetProperty, key);
        }
      },
      selectedKey: this.properties.selectedKey,
      disabled: this.properties.disabled,
      // required to allow the component to be re-rendered by calling this.render() externally
      stateKey: new Date().toString()
    });
    ReactDom.render(element, elem);
  }

  private onChanged(option: IDropdownOption, index?: number): void {
    this.properties.onPropertyChange(this.targetProperty, option.key);
  }
}

// tslint:disable-next-line:max-line-length
export function PropertyPaneAsyncDropdown(targetProperty: string, properties: IPropertyPaneAsyncDropdownProps): IPropertyPaneField<IPropertyPaneAsyncDropdownProps> {
  return new PropertyPaneAsyncDropdownControl(targetProperty, properties);
}
