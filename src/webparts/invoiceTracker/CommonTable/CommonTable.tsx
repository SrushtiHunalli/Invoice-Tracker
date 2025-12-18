import * as React from 'react';
import styles from './CommonTable.module.scss';
import { ICommonTableProps } from './ICommonTableProps';
import { ICommonTableStates } from './ICommonTableStates';
import ColumnsViewPanel from '../CommonTable/ColumnsViewSettings/ColumnsViewPanel';
import { formatDate } from '../Common Functions/utils';
import {
  DetailsList,
  DetailsRow,
  IDetailsHeaderProps,
  Selection,
  SelectionMode,
} from 'office-ui-fabric-react/lib/DetailsList';
import {
  Checkbox,
  ContextualMenu,
  ContextualMenuItemType,
  DefaultButton,
  Icon,
  IContextualMenuItem,
  IRenderFunction,
  Panel,
  PanelType,
  PrimaryButton,
  Stack,
  Sticky,
  StickyPositionType,
  TextField,
} from 'office-ui-fabric-react';
import { ICustomDetailsListColumn } from '../Types/index';

export default class CommonTable extends React.Component<ICommonTableProps, ICommonTableStates> {
  private selection: Selection;

  constructor(props: ICommonTableProps) {
    super(props);

    this.state = {
      // columns
      columns: [],
      visibleColumns: [],

      // data
      items: props.tableContent ?? [],
      selectionMode: props.selectionMode ?? SelectionMode.single,

      // grouping / totals / sorting
      groupBy: '',
      groups: [],
      totalsKey: null,
      totalColumnKey: '',
      sortColumnKey: '',
      sortDirection: null,

      // filters (per column)
      columnFilters: {},
      columnFilterMenu: {
        target: null,
        visible: false,
        columnKey: null,
        contextItems: [],
      },

      // filter panel
      filterColumnKey: null,
      filterColumnName: null,
      filterColumnValues: [],
      filterSearchText: '',
      selectedFilterValues: [],
      isFilterPanelOpen: false,

      // columns panel
      isColumnsPanelOpen: false,
    };

    this.selection = new Selection({
      onSelectionChanged: () => this._getNewSelectionDetails(),
    });
  }

  public componentDidMount(): void {
    this.setTableData();
  }

  public componentDidUpdate(prevProps: Readonly<ICommonTableProps>): void {
    // data changed
    if (prevProps.tableContent !== this.props.tableContent) {
      this.props.onDataFilter && this.props.onDataFilter(this.props.tableContent ?? []);
      this.setState({
        items: this.applyAllFilters(this.props.tableContent ?? [], this.state.columnFilters as any),
      });
    }

    // columns changed
    if (prevProps.mainColumns !== this.props.mainColumns) {
      this.props.onColumnsChange && this.props.onColumnsChange(this.props.mainColumns);

      const columns = this.decorateColumns(this.props.mainColumns);
      const visibleColumns = this.getVisibleColumnsFromStorage(columns);

      this.setState({
        columns,
        visibleColumns,
      });
    }
  }

  // ----------------------- init -----------------------

  private decorateColumns(columns: ICustomDetailsListColumn[]) {
    return [...columns].map((c) => ({
      ...c,
      isResizable: true,
      onRenderHeader: () => this.renderColumnHeaderName(c.key, c.name),
    }));
  }

  private getVisibleColumnsFromStorage(columns: ICustomDetailsListColumn[]) {
    const saved = this.props.localStorageKey
      ? localStorage.getItem(this.props.localStorageKey)
      : null;
    const selected = saved ? JSON.parse(saved) : null;

    if (!selected || !Array.isArray(selected)) {
      return columns;
    }

    return selected
      .map((keyObj: any) => columns.find((col) => col.key === keyObj.key))
      .filter((col) => col) as ICustomDetailsListColumn[];
  }

  private setTableData() {
    const columns = this.decorateColumns(this.props.mainColumns as any);
    const visibleColumns = this.getVisibleColumnsFromStorage(columns);

    this.props.onColumnsChange && this.props.onColumnsChange(visibleColumns);
    this.props.onDataFilter && this.props.onDataFilter(this.props.tableContent ?? []);

    this.setState({
      columns,
      visibleColumns,
      items: this.props.tableContent ?? [],
    });
  }

  // ----------------------- selection -----------------------

  private _getNewSelectionDetails() {
    const selected = this.selection.getSelection();
    if (this.props.selectedItem) {
      if (selected.length) this.props.selectedItem(selected[0]);
      else this.props.selectedItem(false);
    }
  }

  // ----------------------- header (sticky) -----------------------

  private onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) return null;
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        {defaultRender!({ ...props })}
      </Sticky>
    );
  };

  private renderColumnHeaderName = (key: string, displayName: string): JSX.Element => {
    const isFiltered = !!(this.state.columnFilters as any)[key];
    const isSorted = this.state.sortColumnKey === key;
    const isGrouped = this.state.groupBy === key;

    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
        <span>{displayName}</span>
        {isFiltered && (
          <Icon iconName="Filter" styles={{ root: { color: '#0078d4', fontSize: 12 } }} />
        )}
        {isSorted && (
          <Icon
            iconName={this.state.sortDirection === 'asc' ? 'SortUp' : 'SortDown'}
            styles={{ root: { color: '#0078d4', fontSize: 12 } }}
          />
        )}
        {isGrouped && (
          <Icon
            iconName="GroupedDescending"
            styles={{ root: { color: '#0078d4', fontSize: 12 } }}
          />
        )}
      </Stack>
    );
  };

  // ----------------------- filters -----------------------

  private openFilterPanelFromMenu = (): void => {
    const columnKey = this.state.columnFilterMenu.columnKey as string;
    const column = this.state.columns.find((col: any) => col.key === columnKey) as
      | ICustomDetailsListColumn
      | undefined;
    const columnName = column?.name || columnKey;

    const values = this.getUniqueColumnValues(column as any);
    const currentFilter = (this.state.columnFilters as any)[columnKey] || [];

    this.setState({
      isFilterPanelOpen: true,
      filterColumnKey: columnKey,
      filterColumnName: columnName,
      filterColumnValues: values,
      filterSearchText: '',
      selectedFilterValues: currentFilter,
    });
  };

  private clearColumnFilter = (columnKey: string) => {
    const { [columnKey]: _omit, ...restFilters } = this.state.columnFilters as any;
    this.setState({
      columnFilters: restFilters,
      columnFilterMenu: { visible: false, target: null, columnKey: '', contextItems: [] },
      items: this.applyAllFilters(this.props.tableContent ?? [], restFilters),
    });
  };

  private clearFilter(columnKey: string, value: string) {
    const filters = { ...(this.state.columnFilters as any) };
    const newFilterValues = (filters[columnKey] || []).filter((v: string) => v !== value);

    if (newFilterValues.length > 0) {
      filters[columnKey] = newFilterValues;
    } else {
      delete filters[columnKey];
    }

    this.setState({
      columnFilters: filters,
      columnFilterMenu: { visible: false, target: null, columnKey: '', contextItems: [] },
      items: this.applyAllFilters(this.props.tableContent ?? [], filters),
    });
  }

  private applyAllFilters(items: any[], filters: { [key: string]: string[] }) {
    if (!items || !items.length) {
      this.props.onDataFilter && this.props.onDataFilter([]);
      return [];
    }

    const filteredItems = items.filter((item) =>
      Object.keys(filters).every((colKey) => {
        const filterValues = filters[colKey];
        if (!filterValues || filterValues.length === 0) return true;

        const itemValue = (item as any)[colKey];
        const column = this.state.columns.find((c: any) => c.key === colKey) as
          | ICustomDetailsListColumn
          | undefined;

        if (column?.type === 'date') {
          const formattedValue = itemValue ? formatDate(itemValue, true) : 'Blank';
          return filterValues.includes(formattedValue);
        }

        if (column?.type === 'yesNo') {
          const formattedValue = itemValue ? 'Yes' : 'No';
          return filterValues.includes(formattedValue);
        }

        if (filterValues.includes('Blank')) {
          // eslint-disable-next-line eqeqeq
          if (itemValue == null || itemValue == '' || itemValue == undefined) {
            return true;
          }
        }

        const nonBlankFilters = filterValues.filter((v) => v !== 'Blank');
        if (nonBlankFilters.length > 0) {
          return nonBlankFilters.includes(String(itemValue));
        }

        return filterValues.includes('Blank');
      }),
    );

    this.props.onDataFilter && this.props.onDataFilter(filteredItems);
    return filteredItems;
  }

  private toggleFilterValueSelection = (value: string, checked: boolean = false) => {
    const current = [...this.state.selectedFilterValues];
    const updated = checked ? [...current, value] : current.filter((v) => v !== value);
    this.setState({ selectedFilterValues: updated });
  };

  private applySelectedValuesFilters = (): void => {
    const columnKey = this.state.filterColumnKey as string;
    const values = [...this.state.selectedFilterValues];
    const column = this.state.columns.find((c: any) => c.key === columnKey) as
      | ICustomDetailsListColumn
      | undefined;

    const filtered = (this.props.tableContent ?? []).filter((item: any) => {
      if (values.length === 0) return true;

      const itemValue = (item as any)[columnKey];

      if (column?.type === 'date') {
        const formattedValue = itemValue ? formatDate(itemValue, true) : 'Blank';

        if (values.includes('Blank') && !itemValue) return true;
        return values.includes(formattedValue);
      }

      if (column?.type === 'yesNo') {
        const formattedValue = itemValue ? 'Yes' : 'No';
        return values.includes(formattedValue);
      }

      if (values.includes('Blank')) {
        // eslint-disable-next-line eqeqeq
        if (itemValue == null || itemValue == '' || itemValue == undefined) {
          return true;
        }
      }

      const nonBlankValues = values.filter((v) => v !== 'Blank');
      if (nonBlankValues.length > 0 && nonBlankValues.includes(String(itemValue))) {
        return true;
      }

      return false;
    });

    this.props.onDataFilter && this.props.onDataFilter(filtered);

    this.setState({
      columnFilters: {
        ...(this.state.columnFilters as any),
        [columnKey]: values,
      },
      items: filtered,
      isFilterPanelOpen: false,
    });
  };

  private clearPendingFilters = () => {
    this.setState({ selectedFilterValues: [] });
  };

  private getUniqueColumnValues(column: ICustomDetailsListColumn): string[] {
    const columnKey = column.key;
    const columnType = column.type;

    const values = (this.props.tableContent ?? []).map((item: any) => (item as any)[columnKey]);

    const processedValues = values.map((value: any) => {
      if (columnType === 'yesNo') {
        return value ? 'Yes' : 'No';
      }

      if (value == null || value === '' || value === undefined) {
        return 'Blank';
      }

      if (columnType === 'date' && value) {
        try {
          return formatDate(value, true);
        } catch {
          return String(value);
        }
      }

      return String(value);
    });

    const uniqueValues: string[] = Array.from(new Set(processedValues));
    return uniqueValues.sort();
  }

  // ----------------------- sorting / grouping / totals -----------------------

  private sortColumn = (columnKey: string, type: 'text' | 'date' | 'number', direction: 'asc' | 'desc') => {
    let arr = [...this.state.items];

    arr.sort((a: any, b: any) => {
      const aRaw = a[columnKey];
      const bRaw = b[columnKey];

      const aVal = type === 'date' ? new Date(aRaw as any) : aRaw ?? '';
      const bVal = type === 'date' ? new Date(bRaw as any) : bRaw ?? '';

      if (type === 'date') {
        const res = (aVal as any).getTime() - (bVal as any).getTime();
        return direction === 'asc' ? res : -res;
      }

      const aNum = typeof aVal === 'number' ? aVal : parseFloat(aVal);
      const bNum = typeof bVal === 'number' ? bVal : parseFloat(bVal);

      const aIsNumber = !isNaN(aNum) && aVal !== '' && aVal != null;
      const bIsNumber = !isNaN(bNum) && bVal !== '' && bVal != null;

      if (aIsNumber && bIsNumber) {
        return direction === 'asc' ? aNum - bNum : bNum - aNum;
      }

      const aStr = String(aVal).toLowerCase();
      const bStr = String(bVal).toLowerCase();
      const cmp = aStr.localeCompare(bStr, undefined, { numeric: true, sensitivity: 'base' });
      return direction === 'asc' ? cmp : -cmp;
    });

    let sortColumnKey: string | null;
    let sortDirection: 'asc' | 'desc' | null;

    if (this.state.sortColumnKey === columnKey) {
      const clickedSame = this.state.sortDirection === 'asc' && direction === 'asc';
      sortDirection = clickedSame ? null : direction;
      sortColumnKey = clickedSame ? '' : this.state.sortColumnKey;
    } else {
      sortColumnKey = columnKey;
      sortDirection = direction;
    }

    this.setState({
      items: arr,
      columnFilterMenu: { visible: false, target: null, columnKey: '', contextItems: [] },
      sortColumnKey: sortColumnKey || '',
      sortDirection,
    });
  };

  private _onRenderGroupHeader = (props: any) => {
    const { totalsKey, totalColumnKey } = this.state;

    if (!props) return null;

    const isAdditionalDetailsShown = totalsKey === 'sum';
    const column: ICustomDetailsListColumn | undefined = this.state.columns.find(
      (c: any) => c.key === totalColumnKey,
    ) as any;

    let additionDetail: JSX.Element | string = '';

    if (isAdditionalDetailsShown && column) {
      additionDetail =
        typeof props.group[totalColumnKey] !== 'undefined' ? (
          <p>
            <strong>{column.name}</strong>: {props.group[totalColumnKey]}
          </p>
        ) : (
          ''
        );
    }

    return (
      <div style={{ display: 'flex', gap: '1em' }}>
        <p>
          <strong>
            {props.group.key ? props.group.key : 'Blank'} ({props.group.count})
          </strong>
        </p>
        {additionDetail}
      </div>
    );
  };

  private groupBy(col: ICustomDetailsListColumn, type: 'text' | 'date' | 'number') {
    let items = [...this.state.items];

    if (!this.state.groupBy) {
      this.setState({ items, groups: [] });
      return;
    }

    const column: ICustomDetailsListColumn | undefined = this.state.columns.find(
      (c: any) => c.key === this.state.totalColumnKey,
    ) as any;

    const valueArray = [...items]
      .map((item) => (item as any)[col.key])
      .filter((value, index, self) => self.indexOf(value) === index && value != null);

    const groups = valueArray.map((e) => {
      const key = type === 'date' ? formatDate(e, true) : e;
      const startIndex = [...items].map((j) => (j as any)[col.key]).indexOf(e);
      const count = [...items].filter((j) => (j as any)[col.key] === e).length;

      const total =
        column && column.key
          ? [...items]
            .filter((j) => (j as any)[col.key] === e)
            .reduce((acc, total) => {
              return (acc += parseFloat((total as any)[column.key]));
            }, 0)
          : undefined;

      return {
        key,
        startIndex,
        level: 0,
        name: e,
        [column?.key as string]: total,
        count,
      };
    });

    this.setState({
      items,
      groups,
    });
  }

  // ----------------------- column menu -----------------------

  private _onColumnHeaderClick = (e: any, col: ICustomDetailsListColumn) => {
    if (!col) return;

    const sortAtoZ: IContextualMenuItem = {
      key: 'asc',
      text: 'Sort A → Z',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'asc',
      onClick: () => this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'text', 'asc'),
    };

    const sortZtoA: IContextualMenuItem = {
      key: 'desc',
      text: 'Sort Z → A',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'desc',
      onClick: () => this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'text', 'desc'),
    };

    const sortNewest: IContextualMenuItem = {
      key: 'sortNewest',
      text: 'Sort Newest',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'desc',
      onClick: () =>
        this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'date', 'desc'),
    };

    const sortOldest: IContextualMenuItem = {
      key: 'sortOldest',
      text: 'Sort Oldest',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'asc',
      onClick: () =>
        this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'date', 'asc'),
    };

    const sortSmallerToLarger: IContextualMenuItem = {
      key: 'sortSmallerToLarger',
      text: 'Sort Smaller To Larger',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'asc',
      onClick: () =>
        this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'number', 'asc'),
    };

    const sortLargerToSmaller: IContextualMenuItem = {
      key: 'sortLargerToSmaller',
      text: 'Sort Larger To Smaller',
      canCheck: true,
      checked: this.state.sortColumnKey === col.key && this.state.sortDirection === 'desc',
      onClick: () =>
        this.sortColumn(this.state.columnFilterMenu.columnKey as string, 'number', 'desc'),
    };

    const filter: IContextualMenuItem = {
      key: 'filter',
      text: 'Filter...',
      canCheck: true,
      checked: Object.keys(this.state.columnFilters ?? {}).includes(col.key),
      onClick: () => this.openFilterPanelFromMenu(),
    };

    const clearFilter: IContextualMenuItem = {
      key: 'clear',
      text: 'Clear Filter',
      onClick: () => this.clearColumnFilter(this.state.columnFilterMenu.columnKey as string),
    };

    const showHideColumns: IContextualMenuItem = {
      key: 'showHideColumns',
      text: 'Show or Hide Columns',
      onClick: () => this.openColumnsPanel(),
    };

    const groupBy = (type: 'text' | 'date'): IContextualMenuItem => ({
      key: 'groupBy',
      name: 'Group by ' + col.name,
      canCheck: true,
      checked: this.state.groupBy === col.key,
      onClick: () => {
        this.setState(
          { groupBy: this.state.groupBy === col.key ? '' : col.key },
          () => this.groupBy(col, type),
        );
      },
    });

    const totals: IContextualMenuItem = {
      key: 'totals',
      text: 'Totals',
      subMenuProps: {
        items: [
          {
            key: 'sum',
            text: 'Sum',
            canCheck: true,
            checked: this.state.totalsKey === 'sum',
            onClick: () => {
              this.setState((prev) => ({
                totalsKey: prev.totalsKey ? null : 'sum',
                totalColumnKey: prev.totalColumnKey ? '' : col.key,
              }));
            },
          },
        ],
      },
    };

    const divider: IContextualMenuItem = { key: 'divider', itemType: ContextualMenuItemType.Divider };

    const contextItems: IContextualMenuItem[] = [];

    switch (col.type) {
      case 'text':
        contextItems.push(sortAtoZ, sortZtoA, divider, filter, clearFilter, divider, groupBy('text'));
        break;
      case 'date':
        contextItems.push(sortNewest, sortOldest, divider, filter, clearFilter, divider, groupBy('date'));
        break;
      case 'number':
        contextItems.push(
          sortSmallerToLarger,
          sortLargerToSmaller,
          divider,
          filter,
          clearFilter,
          divider,
          totals,
        );
        break;
      case 'multilinetext':
        break;
      case 'yesNo':
        contextItems.push(filter, clearFilter);
        break;
      default:
        break;
    }

    contextItems.length
      ? contextItems.push(divider, showHideColumns)
      : contextItems.push(showHideColumns);

    this.setState({
      columnFilterMenu: {
        target: e?.currentTarget || null,
        visible: true,
        columnKey: col.key,
        contextItems,
      },
    });
  };

  // ----------------------- columns panel -----------------------

  private openColumnsPanel = () => {
    this.setState({ isColumnsPanelOpen: true });
  };

  private closeColumnsPanel = () => {
    this.setState({ isColumnsPanelOpen: false });
  };

  private onColumnsUpdate = (updatedColumns: ICustomDetailsListColumn[]) => {
    if (!updatedColumns.length) return;

    this.props.onColumnsChange && this.props.onColumnsChange(updatedColumns);

    this.setState({ visibleColumns: updatedColumns });
    if (this.props.localStorageKey) {
      localStorage.setItem(this.props.localStorageKey, JSON.stringify(updatedColumns));
    }
  };

  // ----------------------- render -----------------------

  public render(): React.ReactElement<ICommonTableProps> {
    const extraProps: any = {};
    if (this.props.onItemInvoked) extraProps.onItemInvoked = this.props.onItemInvoked;
    if (this.props.onRenderItemColumn) extraProps.onRenderItemColumn = this.props.onRenderItemColumn;

    return (
      <div className={styles.commonTable + ' ms-Grid'}>
        <div className={'ms-Grid-row ' + styles.topSpace}>
          <div className={'ms-Grid-col ms-sm12 ms-md12 ms-lg12 '}>

            {/* applied filters chips */}
            <div
              className="remove-scrollbar"
              style={{
                display: 'flex',
                overflowX: 'auto',
                overflowY: 'hidden',
                whiteSpace: 'nowrap',
                padding: 0,
                gap: '1.5em',
              }}
            >
              {Object.keys(this.state.columnFilters ?? {}).map((columnKey) => {
                const column = this.state.columns.find((c: any) => c.key === columnKey) as
                  | ICustomDetailsListColumn
                  | undefined;
                const key = column?.name ?? columnKey;
                const values = ((this.state.columnFilters as any)[columnKey] ?? []).map(
                  (v: string) => (
                    <span
                      key={v}
                      title="clear filter"
                      style={{
                        padding: '0.25em 0.75em',
                        borderRadius: '1em',
                        cursor: 'pointer',
                        background: '#f3f2f1',
                      }}
                      onClick={() => this.clearFilter(columnKey, v)}
                    >
                      {v.length > 20 ? `${v.substring(0, 20)}...` : v}
                      <Icon styles={{ root: { fontSize: 8, marginLeft: 4 } }} iconName="ChromeClose" />
                    </span>
                  ),
                );

                if (values.length) {
                  return (
                    <div
                      key={columnKey}
                      style={{ display: 'flex', gap: '0.15em', alignItems: 'center' }}
                    >
                      <label>{key}</label>:
                      <div style={{ display: 'flex', gap: '0.25em', alignItems: 'center' }}>
                        {values}
                      </div>
                    </div>
                  );
                }
                return null;
              })}
            </div>

            {/* table with sticky header */}
            <div className={styles.tableContainer}>
              <DetailsList
                columns={this.state.visibleColumns}
                items={this.state.items}
                selection={this.selection}
                selectionMode={this.props.selectionMode ?? SelectionMode.single}
                groups={this.state.groups.length ? this.state.groups : undefined}
                onRenderDetailsHeader={this.onRenderDetailsHeader}
                compact={true}
                groupProps={{
                  onRenderHeader:
                    this.state.groups && this.state.groups.length
                      ? this._onRenderGroupHeader
                      : undefined,
                }}
                isHeaderVisible={true}
                styles={{
                  headerWrapper: {
                    position: 'sticky',
                    top: 0,
                    zIndex: 1,
                    background: 'white',
                    boxShadow: '0px 2px 4px rgba(0,0,0,0.1)',
                    selectors: {
                      '.ms-DetailsHeader-cellTitle': { fontWeight: 700 },
                    },
                  },
                }}
                onRenderRow={(props) => (
                  <DetailsRow
                    {...props}
                    styles={{
                      root: {
                        backgroundColor:
                          props.itemIndex % 2 === 0 ? '#f3f2f1' : 'white',
                      },
                    }}
                  />
                )}
                onColumnHeaderClick={this._onColumnHeaderClick}
                {...extraProps}
              />
            </div>
          </div>

          {this.state.columnFilterMenu.visible && (
            <ContextualMenu
              items={this.state.columnFilterMenu.contextItems}
              target={this.state.columnFilterMenu.target}
              onDismiss={() =>
                this.setState({
                  columnFilterMenu: {
                    visible: false,
                    target: null,
                    columnKey: '',
                    contextItems: [],
                  },
                })
              }
            />
          )}

          {!this.props.tableContent?.length && (
            <div className={'ms-Grid-col ms-sm12 ms-md12 ms-lg12 ' + styles.emptyList}>
              <label>No items found</label>
            </div>
          )}

          <Panel
            isOpen={this.state.isFilterPanelOpen}
            onDismiss={() => this.setState({ isFilterPanelOpen: false })}
            headerText={`Filter by ${this.state.filterColumnName || this.state.filterColumnKey}`}
            type={PanelType.smallFixedFar}
            onRenderFooterContent={() => (
              <div style={{ display: 'flex', justifyContent: 'space-around' }}>
                <PrimaryButton
                  text="Apply"
                  onClick={this.applySelectedValuesFilters}
                  disabled={this.state.selectedFilterValues.length === 0}
                />
                <DefaultButton text="Clear all" onClick={this.clearPendingFilters} />
              </div>
            )}
          >
            <TextField
              placeholder="Type text to find a filter"
              value={this.state.filterSearchText}
              onChange={(_, newVal) => this.setState({ filterSearchText: newVal || '' })}
              styles={{ root: { marginBottom: 12 } }}
            />

            <Stack tokens={{ childrenGap: 8 }} style={{ maxHeight: '80vh', overflowY: 'auto' }}>
              {this.state.filterColumnValues
                .filter((val) =>
                  val.toLowerCase().includes(this.state.filterSearchText.toLowerCase()),
                )
                .map((val, index) => (
                  <Checkbox
                    key={index}
                    label={val}
                    checked={this.state.selectedFilterValues?.includes(val)}
                    onChange={(_, checked) => this.toggleFilterValueSelection(val, !!checked)}
                  />
                ))}
            </Stack>
          </Panel>

          {this.state.isColumnsPanelOpen && (
            <ColumnsViewPanel
              allColumns={this.state.columns}
              visibleColumns={this.state.visibleColumns}
              onClose={this.closeColumnsPanel}
              onUpdateColumns={this.onColumnsUpdate}
            />
          )}
        </div>
      </div>
    );
  }
}
