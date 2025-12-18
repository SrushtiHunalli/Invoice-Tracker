/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-unused-expressions */
import { Dropdown, SearchBox, IDropdownStyles, IDropdownOption, DropdownMenuItemType } from "@fluentui/react";
import * as  React from 'react';
import { ISearchableDropdownProps } from './ISearchableDropdownProps';
import { ISearchableDropdownState } from './ISearchableDropdownState';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdownItemsWrapper: { maxHeight: 350 }, dropdownItem: {
    minHeight: 32,
    maxHeight: 100,


  }, dropdownItemSelected: {
    minHeight: 32,
    maxHeight: 100,

  }
};

export class SearchableDropdown extends React.Component<ISearchableDropdownProps, ISearchableDropdownState> {
  constructor(props: ISearchableDropdownProps | Readonly<ISearchableDropdownProps>) {
    super(props);
    this.state = {
      searchText: ""
    };
  }

  // private renderSearchBox = (option: IDropdownOption): JSX.Element => {
  //   return (option.itemType === DropdownMenuItemType.Header && option.key === "FilterHeader") && this.props.options.length > 10 ?
  //     <SearchBox styles={{
  //   root: { height: 32 },
  //   field: { fontSize: 14, paddingTop: 4 },
  // }} onChange={(e, newText) => this.setState({ searchText: newText ?? "" })} placeholder='search' underlined />
  //     : this.props.onRenderOption ? this.props.onRenderOption(option) : <>{option.text}</>;
  // };

  private renderSearchBox = (option: IDropdownOption): JSX.Element => {
    const isFilterHeader =
      option.itemType === DropdownMenuItemType.Header &&
      option.key === "FilterHeader";

    if (isFilterHeader) {
      return (
        <SearchBox
          styles={{
            root: { height: 32, padding: "4px 8px" },
            field: { fontSize: 14, paddingTop: 0 },
          }}
          placeholder="Search..."
          underlined={false}
          onChange={(_, newText) =>
            this.setState({ searchText: newText ?? "" })
          }
        />
      );
    }

    return this.props.onRenderOption
      ? this.props.onRenderOption(option)
      : <>{option.text}</>;
  };

  render() {
    const { labelText, options, placeholder, multiSelect, required, selectedItem,
      disabled, selectedItems, styles, onDropdownDismiss, onChangeHandler, onRenderLabel } = this.props;
    const { searchText } = this.state;

    const modifiedOptions: IDropdownOption[] = [
      {
        key: "FilterHeader",
        text: "",
        itemType: DropdownMenuItemType.Header,
        disabled: true,
      },
      {
        key: "divider_filterHeader",
        text: "",
        itemType: DropdownMenuItemType.Divider,
        disabled: true,
      },
      ...options.map(o =>
        !searchText ||
          o.text.toLowerCase().includes(searchText.toLowerCase())
          ? o
          : { ...o, hidden: true }
      ),
    ];

    if (options.length < 5) {
      modifiedOptions.splice(0, 2);
    }

    return (
      <Dropdown
        label={labelText || ""}
        onRenderLabel={onRenderLabel}
        placeholder={placeholder}
        multiSelect={multiSelect || false}
        required={required}
        options={modifiedOptions}
        defaultSelectedKey={selectedItem}
        selectedKeys={selectedItems}
        onRenderOption={this.renderSearchBox}
        onChange={onChangeHandler}
        onDismiss={() => {
          this.setState({ searchText: "" });
          onDropdownDismiss && onDropdownDismiss();
        }}
        styles={{ ...dropdownStyles, ...styles }}
        disabled={disabled}
      />
    );
  }
}

export default SearchableDropdown;
