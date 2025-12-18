/* eslint-disable @typescript-eslint/no-explicit-any */
// import { IDropdownOption, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { IDropdownOption,IDropdownStyles } from "@fluentui/react";
export interface ISearchableDropdownProps {
    disabled?: boolean;

    labelText?: string;
    onRenderLabel?(): JSX.Element;

    multiSelect?: boolean;
    selectedItems?: any[]; // for multiple options
    selectedItem?: any;
    options: IDropdownOption[];
    placeholder?: string;
    required?: boolean;

    styles?: Partial<IDropdownStyles>;

    // Dropdown change handler.
    onChangeHandler(e: any, option: IDropdownOption): void;

    // When the dropdown window closes. 
    onDropdownDismiss?(): void;

    // to modify the option
    onRenderOption?(option: IDropdownOption): JSX.Element;
}