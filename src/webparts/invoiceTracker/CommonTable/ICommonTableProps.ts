import { IColumn } from '@fluentui/react';

export interface ICommonTableProps {
    tableContent: any;
    mainColumns: any;
    selectedItem?: any;
    selectionMode?: any;
    onRowClick?: (item: any) => void;
    onSelectionChange?: (selectedItems: any[]) => void;
    selectedRowId?: number;
    items?: any[];
    onActiveItemChanged?: (item: any) => void;
    selection?: any;
    columns?: any[];
    setKey?: string;
    onDataFilter?(items: any[]): void;
    onRenderItemColumn?: (item?: any, index?: number, column?: IColumn) => React.ReactNode;
    onItemInvoked?: (item?: any, index?: number, ev?: Event) => void;
    groups?: any;
    _onRenderGroupFooter?: any;
    RemoveSelection?: boolean;
    onColumnsChange?(selectedColumns: any[]): void;
    localStorageKey?: string;
}
