import { SelectionMode } from "office-ui-fabric-react/lib/DetailsList";
import { IContextualMenuItem } from "office-ui-fabric-react";
import { ICustomDetailsListColumn } from "../Types/index";

export interface ICommonTableStates {
  // columns
  items: any[];
  columns: ICustomDetailsListColumn[];
  visibleColumns: ICustomDetailsListColumn[];

  // selection
  selectionMode: SelectionMode;
  RemoveSelection?: boolean;

  // grouping
  groupBy: string;
  groups: any[];

  // sorting
  sortColumnKey: string;
  sortDirection: "asc" | "desc" | null;

  // totals
  totalsKey: "sum" | null;
  totalColumnKey: string;

  // column filter menu
  isFilterPanelOpen: boolean;
  // key -> list of selected filter values
  columnFilters: { [key: string]: string[] };
  columnFilterMenu: {
    target: any;
    visible: boolean;
    columnKey: string | null;
    contextItems: IContextualMenuItem[];
  };

  // filter panel content
  filterColumnKey: string | null;
  filterColumnName: string | null;
  filterColumnValues: string[];
  filterSearchText: string;
  selectedFilterValues: string[];

  // column selection panel
  isColumnsPanelOpen: boolean;
}
