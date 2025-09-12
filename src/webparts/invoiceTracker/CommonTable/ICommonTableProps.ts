export interface ICommonTableProps {
    tableContent: any;
    mainColumns: any;
    selectedItem?: any;
    selectionMode?: any;
    onRowClick?: (item: any) => void;
    onSelectionChange?: (selectedItems: any[]) => void;
    selectedRowId?: number;
}
