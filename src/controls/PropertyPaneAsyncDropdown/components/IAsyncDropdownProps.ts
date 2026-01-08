import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface IAsyncDropdownProps {
  label: string;
  loadOptions: () => Promise<IDropdownOption[]>;
  onChanged: (option: IDropdownOption, index?: number) => void;
  callback?: (selectedKey: string | number) => void;
  selectedKey: string | number;
  disabled: boolean;
  stateKey: string;
}