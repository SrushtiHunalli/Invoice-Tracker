import * as React from "react";
import { useState, useEffect } from "react";
import {
  SearchBox,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  Spinner,
  Stack,
  IColumn,
  Panel,
  TextField,
  DetailsList,
  Selection,
  SelectionMode,
  PanelType,
  IconButton,
  Text,
  ScrollablePane,
  IDetailsHeaderProps,
  IRenderFunction,
  Sticky,
  StickyPositionType,
  ContextualMenu,
  ContextualMenuItemType,
  Separator,
  Dropdown,
  Label,
  TooltipHost,
  DetailsListLayoutMode,
  Icon,
  Checkbox
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import { MSGraphClient } from '@microsoft/sp-http';
import styles from "./CreateView.module.scss"
import DocumentViewer from '../DocumentViewer';
interface CreateViewProps {
  sp: SPFI;
  context: any;
  projectsp: SPFI;
  effectiveUserLogin?: string;
}
type PurchaseOrderItem = {
  Id: number;
  POID: string;
  ProjectName?: string;
  POAmount?: string;
  Currency?: string;
  POComments?: string;
  CostCenter?: string;
};
type ChildPOItem = {
  Id: number;
  POID: string;
  POAmount: string;
  ParentPOIndex: number;
  POIndex: number;
};
type InvoiceRequest = {
  Id: number;
  PurchaseOrderPO: string;
  Amount: number;
  Status: string;
  ProjectName?: string;
  POItemTitle?: string;
  POItemValue?: number;
  CustomerContact?: string;
  Comments?: string;
  PMCommentsHistory?: string;
  FinanceCommentsHistory?: string;
  Created?: string;
  CreatedBy?: string;
  Modified?: string;
  ModifiedBy?: string;
  CurrentStatus?: string;
  InvoicedAmountsJSON?: string;
};
type InvoiceFormState = {
  POID: string;
  PurchaseOrder: string;
  ProjectName: string;
  POAmount: string;
  POItemTitle: string;
  POItemValue: string;
  InvoiceAmount: string;
  CustomerContact: string;
  Comments: string;
  Attachment: File | null;
  CostCenter?: string;
};
type LineAllocation = {
  poItemId: string;          // childPO.POID
  title: string;             // "Line item 1" etc.
  poItemValue: number;
  remaining: number;
  invoiceAmount: number;     // user-entered
  error?: string;
};
const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";
const CreateView: React.FC<CreateViewProps> = ({ sp, projectsp, context, effectiveUserLogin }) => {
  // const [mergedItems, ] = useState<PurchaseOrderItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [filters, setFilters] = useState({ search: "" });
  // const [customerOptions, ] = useState<IDropdownOption[]>([]);
  const [selectedItem, setSelectedItem] = useState<PurchaseOrderItem | null>(null);
  const [error, setError] = useState<string>("");
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [childPOItems, setChildPOItems] = useState<ChildPOItem[]>([]);
  const [fetchingChildPOs, setFetchingChildPOs] = useState(false);
  const [invoiceRequests, setInvoiceRequests] = useState<InvoiceRequest[]>([]);
  // const [invoiceRequestsForPercent, setInvoiceRequestsForPercent] = useState<InvoiceRequest[]>([]);
  const [fetchingInvoices, setFetchingInvoices] = useState(false);
  const [activePOIDFilter, setActivePOIDFilter] = useState<string | null>(null);
  const [childPOSelection] = useState(new Selection());
  const [invoiceAmountError, setInvoiceAmountError] = useState<string | undefined>(undefined);
  const [isDragActive, setIsDragActive] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [previewFileIdx, setPreviewFileIdx] = useState<number | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitDialogState, setSubmitDialogState] = useState<'idle' | 'submitting' | 'success'>('idle');
  const [, setSubmitDialogMessage] = useState('');
  const [selection] = useState(() =>
    new Selection({

      onSelectionChanged: () => {
        const sel = selection.getSelection()[0];
        setSelectedItem(sel ? (sel as PurchaseOrderItem) : null);
      },
    })
  );
  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");
  const [dialogType, setDialogType] = useState<"success" | "error">("success");
  const [lineAllocations, setLineAllocations] = useState<LineAllocation[]>([]);
  const [currentUserEmail, setCurrentUserEmail] = useState<string>("");
  const [isAdminUser, setIsAdminUser] = useState<boolean>(false);
  const [allProjects, setAllProjects] = useState<any[]>([]); // to hold Projects list data
  const [userGroups, setUserGroups] = useState<string[]>([]);
  // Invoice panel state
  const [isInvoicePanelOpen, setIsInvoicePanelOpen] = useState(false);
  const [invoicePanelPO, setInvoicePanelPO] = useState<ChildPOItem | null>(null);
  const [, setAllInvoicePOs] = useState<any[]>([]);
  const [invoiceCurrency, setInvoiceCurrency] = useState<string>("");
  const [mainPOs, setMainPOs] = useState<PurchaseOrderItem[]>([]);
  // const [columnFilterMenu, setColumnFilterMenu] = React.useState<{ visible: boolean; target: HTMLElement | null; columnKey: string | null }>({ visible: false, target: null, columnKey: null });
  // Main PO list (existing)
  const [mainColumnFilterMenu, setMainColumnFilterMenu] = useState({
    visible: false, target: null, columnKey: null
  });

  // Child PO list
  const [childColumnFilterMenu, setChildColumnFilterMenu] = useState({
    visible: false, target: null, columnKey: null
  });

  // Invoice list
  // const [invoiceColumnFilterMenu, setInvoiceColumnFilterMenu] = useState({
  //   visible: false, target: null, columnKey: null
  // });
  interface ColumnFilterMenu {
    visible: boolean;
    target?: HTMLElement;
    columnKey: string | null;
  }

  const [columnFilterMenu, setColumnFilterMenu] = React.useState<ColumnFilterMenu>({
    visible: false,
    target: undefined,
    columnKey: null,
  });

  const [visibleColumns, setVisibleColumns] = useState<string[]>([]);
  const [columnOrder,] = useState<Record<string, number>>({});
  const [isColumnPanelOpen, setIsColumnPanelOpen] = useState<boolean>(false);

  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState<boolean>(false);
  const [currentFilterColumn, setCurrentFilterColumn] = useState<string>('');

  const [isReadOnlyInvoicePanel, setIsReadOnlyInvoicePanel] = useState(false);
  const [isInvoiceRequestViewPanelOpen, setIsInvoiceRequestViewPanelOpen] = useState(false);
  const [selectedInvoiceRequest, setSelectedInvoiceRequest] = useState<InvoiceRequest | null>(null);
  const onInvoiceRequestClicked = (item: InvoiceRequest) => {
    setSelectedInvoiceRequest(item);
    setIsInvoiceRequestViewPanelOpen(true);
  };
  const [filterMode, setFilterMode] = useState<'mainPO' | 'childPO'>('mainPO');
  const [invoiceStatusFilter, setInvoiceStatusFilter] = React.useState<string | null>(null); // "NotPaid" | "PartiallyInvoiced" | "CompletelyInvoiced"
  const invoiceStatusOptions = [
    { key: "NotPaid", text: "Not Paid" },
    { key: "PartiallyInvoiced", text: "Partially Invoiced" },
    { key: "CompletelyInvoiced", text: "Completely Invoiced" },
  ];
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [previewFileName, setPreviewFileName] = useState<string>('');
  const [isViewerOpen, setIsViewerOpen] = useState(false);
  // Example function to open preview, call this on file click
  const openPreview = (url: string, fileName: string) => {
    setPreviewUrl(url);
    setPreviewFileName(fileName);
    setIsViewerOpen(true);
  };

  // Example function to close preview
  const closePreview = () => {
    setIsViewerOpen(false);
    setPreviewUrl(null);
    setPreviewFileName('');
  };
  const isFilterApplied = !!filters.search || !!invoiceStatusFilter;
  const [invoiceFormState, setInvoiceFormState] = useState<InvoiceFormState>({
    POID: "",
    PurchaseOrder: "",
    ProjectName: "",
    POItemTitle: "",
    POItemValue: "",
    POAmount: "",
    InvoiceAmount: "",
    CustomerContact: "",
    Comments: "",
    Attachment: null,
    CostCenter: "",
  });
  const onColumnHeaderClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn) => {
    if (column && ev) {
      setColumnFilterMenu({
        visible: true,
        target: ev.currentTarget,
        columnKey: column.key as string
      });
    }
  };

  const handleChildPOSort = (column?: IColumn) => {
    if (!column) return;
    const isDesc = !column.isSortedDescending;
    const key = column.fieldName as keyof ChildPOItem;

    const sorted = [...sortedChildItems].sort((a, b) => {
      const av = (a as any)[key];
      const bv = (b as any)[key];
      if (av == null) return -1;
      if (bv == null) return 1;
      const sa = av.toString();
      const sb = bv.toString();
      return isDesc ? sb.localeCompare(sa) : sa.localeCompare(sb);
    });

    const newCols = sortedChildColumns.map(c => ({
      ...c,
      isSorted: c.key === column.key,
      isSortedDescending: c.key === column.key ? isDesc : false,
    }));

    setSortedChildItems(sorted);
    setSortedChildColumns(newCols);
  };

  const columns: IColumn[] = [
    { key: "POID", name: "Purchase Order", fieldName: "POID", minWidth: 100, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
    { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 150, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
    { key: "POComments", name: "PO Comments", fieldName: "POComments", minWidth: 70, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
    {
      key: 'Customer',
      name: 'Customer',
      fieldName: 'Customer',
      minWidth: 100,
      isResizable: true,
      isCollapsible: true,
      onColumnClick: onColumnHeaderClick
    },
    {
      key: "POAmount", name: "PO Amount", fieldName: "POAmount", minWidth: 100, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick, onRender: (item) => {
        // return `${item.POAmount} ${item.Currency ?? ''}`.trim();
        const currencyCode = item.Currency && item.Currency.trim() !== "" ? item.Currency : "USD";
        const symbol = getCurrencySymbol(currencyCode);
        return <span>{symbol} {item.POAmount}</span>;
      }
    },
    {
      key: 'InvoicedPercent',
      name: 'Invoiced %',
      minWidth: 100,
      isResizable: true,
      isCollapsible: true,
      onColumnClick: onColumnHeaderClick,
      onRender: item => {
        return `${calculateInvoicedPercent(item.POID, item.POAmount).toFixed(0)}%`;
      }
    },
    {
      key: "RequestedAmount",
      name: "Requested Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: PurchaseOrderItem) => {
        const { requested } = getTotalsFromJsonForPO(item.POID, invoiceRequests)
        const symbol = getCurrencySymbol(item.Currency || 'USD')
        return requested > 0 ? <span>{symbol}{requested.toLocaleString()}</span> : <span>-</span>
      }
    },
    {
      key: "InvoicedAmount",
      name: "Invoiced Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: PurchaseOrderItem) => {
        const { invoiced } = getTotalsFromJsonForPO(item.POID, invoiceRequests)
        const symbol = getCurrencySymbol(item.Currency || 'USD')
        return invoiced > 0 ? <span>{symbol}{invoiced.toLocaleString()}</span> : <span>-</span>
      }
    },
    {
      key: "PaidAmount",
      name: "Paid Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: PurchaseOrderItem) => {
        const { paid } = getTotalsFromJsonForPO(item.POID, invoiceRequests)
        const symbol = getCurrencySymbol(item.Currency || 'USD')
        return paid > 0 ? <span>{symbol}{paid.toLocaleString()}</span> : <span>-</span>
      }
    },
  ];
  useEffect(() => {
    const initialVisible = columns.map(col => col.key as string);
    setVisibleColumns(initialVisible);
  }, []);
  const invoiceColumnsView: IColumn[] = [
    { key: "Purchase Order", name: "Purchase Order", fieldName: "PurchaseOrder", minWidth: 130, isResizable: true },
    {
      key: "RequestedAmount",
      name: "Requested Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        const { requested } = getTotalsFromJsonForInvoice(item)
        const symbol = getCurrencySymbol(invoiceCurrency || 'USD')

        return requested > 0
          ? <span>{symbol}{requested.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    {
      key: "InvoicedAmountCalc",
      name: "Invoiced Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        const { invoiced } = getTotalsFromJsonForInvoice(item)
        const symbol = getCurrencySymbol(invoiceCurrency || 'USD')

        return invoiced > 0
          ? <span>{symbol}{invoiced.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    {
      key: "PaidAmountCalc",
      name: "Paid Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        const { paid } = getTotalsFromJsonForInvoice(item)
        const symbol = getCurrencySymbol(invoiceCurrency || 'USD')

        return paid > 0
          ? <span>{symbol}{paid.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    {
      key: "CancelledAmountCalc",
      name: "Cancelled Amount",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        const { cancelled } = getTotalsFromJsonForInvoice(item)
        const symbol = getCurrencySymbol(invoiceCurrency || 'USD')
        return cancelled > 0
          ? <span>{symbol}{cancelled.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 140, isResizable: true },

    { key: "CurrentStatus", name: "Current Status", fieldName: "CurrentStatus", minWidth: 140, isResizable: true },
    {
      key: "PMCommentsHistory",
      name: "Requestor Comments",
      fieldName: "PMCommentsHistory",
      minWidth: 200,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        if (!item.PMCommentsHistory) return "No Requestor Comments";
        try {
          const comments = formatCommentHistory(item.PMCommentsHistory);
          return comments
        } catch {
          return "Invalid PMCommentsHistory";
        }
      }
    },
    {
      key: "FinanceCommentsHistory",
      name: "Finance Comments",
      fieldName: "FinanceCommentsHistory",
      minWidth: 200,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        if (!item.FinanceCommentsHistory) return "No Finance Comments";
        try {
          const comments = formatCommentHistory(item.FinanceCommentsHistory);
          return comments
        } catch {
          return "Invalid FinanceCommentsHistory";
        }
      }
    },
    {
      key: "Created",
      name: "Created",
      fieldName: "Created",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => new Date(item.Created).toLocaleString()
    },
    {
      key: "CreatedBy",
      name: "Created By",
      fieldName: "CreatedBy",
      minWidth: 150,
      isResizable: true
    },
    {
      key: "Modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 120,
      isResizable: true,
      onRender: (item: InvoiceRequest) => new Date(item.Modified).toLocaleString()
    },
    {
      key: "ModifiedBy",
      name: "Modified By",
      fieldName: "ModifiedBy",
      minWidth: 150,
      isResizable: true
    }
  ];
  const childPOColumns: IColumn[] = [
    {
      key: "POID",
      name: "PO Item Title",
      fieldName: "POID",
      minWidth: 150,
      isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => (
        <span style={{ color: "#0078d4", cursor: "pointer", fontWeight: 500 }}>{item.POID}</span>
      ),
    },
    {
      key: "POItemValue",
      name: `PO Item Value`,
      fieldName: "POItemValue",
      minWidth: 120,
      // isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => {
        const currencyCode = invoiceCurrency && invoiceCurrency.trim() !== "" ? invoiceCurrency : "USD";
        const symbol = getCurrencySymbol(currencyCode);
        return <span>{symbol} {item.POAmount}</span>;
      }
    },

    {
      key: "POAmount", name: `Remaining Item Value`, fieldName: "POAmount", minWidth: 120, isResizable: true, onColumnClick: (_ev, col) => handleChildPOSort(col), onRender: (item: ChildPOItem) => {
        const currencyCode = invoiceCurrency && invoiceCurrency.trim() !== "" ? invoiceCurrency : "USD";
        const symbol = getCurrencySymbol(currencyCode);
        const remaining = getRemainingPOAmount(item, invoiceRequests);
        return <span>{symbol} {remaining}</span>;
      }
    },
    {
      key: "RequestedAmountItem",
      name: "Requested Amount",
      minWidth: 120,
      isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => {
        const { requested } = getTotalsFromJsonForPOItem(
          selectedItem?.POID ?? '',
          item.POID,
          invoiceRequests
        )

        const currencyCode = invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
        const symbol = getCurrencySymbol(currencyCode)

        return requested > 0
          ? <span>{symbol}{requested.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    {
      key: "InvoicedAmountItem",
      name: "Invoiced Amount",
      minWidth: 120,
      isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => {
        const { invoiced } = getTotalsFromJsonForPOItem(
          selectedItem?.POID ?? '',
          item.POID,
          invoiceRequests
        )

        const currencyCode = invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
        const symbol = getCurrencySymbol(currencyCode)

        return invoiced > 0
          ? <span>{symbol}{invoiced.toLocaleString()}</span>
          : <span>-</span>
      }
    },
    {
      key: "PaidAmountItem",
      name: "Paid Amount",
      minWidth: 120,
      isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => {
        const { paid } = getTotalsFromJsonForPOItem(
          selectedItem?.POID ?? '',
          item.POID,
          invoiceRequests
        )

        const currencyCode = invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
        const symbol = getCurrencySymbol(currencyCode)

        return paid > 0
          ? <span>{symbol}{paid.toLocaleString()}</span>
          : <span>-</span>
      }
    },

    {
      key: 'InvoicedPercentItem',
      name: 'Invoiced %',
      minWidth: 100,
      isResizable: true,
      onColumnClick: (_ev, col) => handleChildPOSort(col),
      onRender: (item: ChildPOItem) => {
        const invoicedPercent = calculateInvoicedPercentForItem(item.POID, parseFloat(item.POAmount));
        return `${invoicedPercent.toFixed(0)}%`;
      }
    },

  ];
  const [invoicePanelLoading, setInvoicePanelLoading] = useState(false);
  const [sortedChildColumns, setSortedChildColumns] = useState<IColumn[]>(childPOColumns);
  const [sortedChildItems, setSortedChildItems] = useState<ChildPOItem[]>(childPOItems);

  useEffect(() => {
    setSortedChildItems(childPOItems);
  }, [childPOItems]);
  // const menuItems = [
  //   {
  //     key: 'sortAsc',
  //     text: 'Sort A→Z',
  //     iconProps: { iconName: 'SortUp' },
  //     onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc')
  //   },
  //   {
  //     key: 'sortDesc',
  //     text: 'Sort Z→A',
  //     iconProps: { iconName: 'SortDown' },
  //     onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc')
  //   },
  //   { key: 'divider1', itemType: ContextualMenuItemType.Divider },
  //   {
  //     key: 'filter',
  //     text: 'Filter Column',
  //     iconProps: { iconName: 'Filter' },
  //     onClick: () => {
  //       setCurrentFilterColumn(columnFilterMenu.columnKey!);
  //       setIsFilterPanelOpen(true);
  //     }
  //   },
  //   {
  //     key: 'clearFilter',
  //     text: 'Clear Filter',
  //     iconProps: { iconName: 'ClearFilter' },
  //     onClick: () => clearColumnFilter(columnFilterMenu.columnKey!)
  //   },
  //   { key: 'divider2', itemType: ContextualMenuItemType.Divider },
  //   {
  //     key: 'columns',
  //     text: 'Manage Columns',
  //     iconProps: { iconName: 'Columns' },
  //     onClick: () => setIsColumnPanelOpen(true)
  //   }
  // ];
  const mainMenuItems = [
    {
      key: 'sortAsc', text: 'Sort A-Z', iconProps: { iconName: 'SortUp' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc')
    },
    {
      key: 'sortDesc', text: 'Sort Z-A', iconProps: { iconName: 'SortDown' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc')
    },
    { key: 'divider1', itemType: ContextualMenuItemType.Divider },
    {
      key: 'filter', text: 'Filter Column', iconProps: { iconName: 'Filter' },
      onClick: () => {
        setCurrentFilterColumn(columnFilterMenu.columnKey!);
        setIsFilterPanelOpen(true);
      }
    },
    {
      key: 'clearFilter', text: 'Clear Filter', iconProps: { iconName: 'ClearFilter' },
      onClick: () => clearColumnFilter(columnFilterMenu.columnKey!)
    },
    { key: 'divider2', itemType: ContextualMenuItemType.Divider },
    {
      key: 'columns', text: 'Manage Columns', iconProps: { iconName: 'Columns' },
      onClick: () => setIsColumnPanelOpen(true)
    },
  ];

  const simpleMenuItems = [
    {
      key: 'sortAsc', text: 'Sort A-Z', iconProps: { iconName: 'SortUp' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc')
    },
    {
      key: 'sortDesc', text: 'Sort Z-A', iconProps: { iconName: 'SortDown' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc')
    },
    { key: 'divider1', itemType: ContextualMenuItemType.Divider },
    {
      key: 'filter', text: 'Filter Column', iconProps: { iconName: 'Filter' },
      onClick: () => {
        setCurrentFilterColumn(columnFilterMenu.columnKey!);
        setIsFilterPanelOpen(true);
      }
    },
    {
      key: 'clearFilter', text: 'Clear Filter', iconProps: { iconName: 'ClearFilter' },
      onClick: () => clearColumnFilter(columnFilterMenu.columnKey!)
    },
  ];

  const handleInvoicePanelDismiss = () => {
    setIsInvoicePanelOpen(false);
    setInvoicePanelPO(null);
    // Clear uploaded attachments on panel close
    setInvoiceFormState(prev => ({
      ...prev,
      Attachment: null,
    }));
    setUploadedFiles([]);
  };
  const handlePOIDDoubleClick = async (item: PurchaseOrderItem) => {
    if (!selectedItem) return;
    setInvoiceCurrency(selectedItem.Currency || "");
    setFetchingChildPOs(true);
    setFetchingInvoices(true);
    setChildPOItems([]);
    setInvoiceRequests([]);
    setActivePOIDFilter(selectedItem?.POID || null);
    setIsPanelOpen(false);
    setIsInvoicePanelOpen(false);
    setIsReadOnlyInvoicePanel(true);

    try {
      const allSalesRecords = await fetchAllPORecords();
      console.log("Fetched all sales records:", allSalesRecords);

      const rec = allSalesRecords.find((r: any) => findPOIndex(r, selectedItem.POID) !== null);

      if (!rec) {

        const invoices = await fetchInvoiceRequests(sp, [selectedItem.POID]);
        setInvoiceRequests(invoices);

        await handleOpenInvoicePanelSinglePO(selectedItem, "");
        return;
      }
      const poIndex = findPOIndex(rec, selectedItem.POID)!;
      let children = getChildItemsForPO(rec, poIndex);
      console.log("Child POs from JSON or ChildPO fields:", children);
      if (!children.length) {
        children = findChildPOsByParentPOID(allSalesRecords, selectedItem.POID);
        console.log("Child POs found by scanning ParentPOID columns:", children);
      }
      setChildPOItems(children);
      const poids = children.length > 0 ? [selectedItem.POID, ...children.map(c => c.POID)] : [selectedItem.POID];
      const invoices = await fetchInvoiceRequests(sp, poids);
      setInvoiceRequests(invoices);

      const poamount = poIndex !== null ? (rec[`POAmount${poIndex === 0 ? "" : poIndex + 1}`] || "") : "";

      if (children.length === 0) {
        await handleOpenInvoicePanelSinglePO(selectedItem, poamount);
      } else {
        setIsPanelOpen(true);
      }
    } catch (err) {
      console.error("Error during PO handling", err);
    } finally {
      setFetchingChildPOs(false);
      setFetchingInvoices(false);
    }
  };

  useEffect(() => {
    (async () => {
      setLoading(true);
      setError("");
      try {
        const items = await sp.web.lists.getByTitle("InvoicePO")
          .items
          .select("ID", "POID", "ParentPOID", "POAmount", "LineItemsJSON", "ProjectName", "Currency", "POComments", "Customer", "CostCenter")();

        setAllInvoicePOs(items);

        // Build a POID-to-item map
        const poidMap = new Map(items.map(item => [item.POID, item]));

        // Filter Main POs (ParentPOID empty)
        const mains = items.filter(item => !item.ParentPOID || item.ParentPOID.trim() === "")
          .map(item => ({
            Id: item.ID,
            POID: item.POID,
            ProjectName: item.ProjectName || "",
            POAmount: item.POAmount || "",
            Currency: item.Currency || "",
            POComments: item.POComments || "",
            Customer: item.Customer || "",
            CostCenter: item.CostCenter || "",
          }));

        const childrenByMainPO = new Map<string, ChildPOItem[]>();

        for (const main of mains) {
          const children: ChildPOItem[] = [];

          // Find direct child POs with ParentPOID = main POID
          items.forEach(item => {
            if (item.ParentPOID && item.ParentPOID.trim() === main.POID.trim()) {
              children.push({
                Id: item.ID,
                POID: item.POID,
                POAmount: item.POAmount || "",
                ParentPOIndex: 0,
                POIndex: 0,
              });
            }
          });

          // Parse and add LineItemsJSON if present
          const mainItem = poidMap.get(main.POID);
          if (mainItem?.LineItemsJSON) {
            try {
              const lineItems = JSON.parse(mainItem.LineItemsJSON);
              if (Array.isArray(lineItems)) {
                lineItems.forEach((li: any, idx: number) => {
                  children.push({
                    Id: mainItem.ID * 1000 + idx, // unique ID for line items
                    POID: li.POID || `LineItem${idx + 1}`,
                    POAmount: li.POAmount || "0",
                    ParentPOIndex: 0,
                    POIndex: 0,
                  });
                });
              }
            } catch {
              // Ignore JSON parse errors silently
            }
          }

          childrenByMainPO.set(main.POID, children);
        }

        setMainPOs(mains);
      } catch (e: any) {
        setError("Error loading POs: " + (e.message || e));
        setMainPOs([]);

      } finally {
        setLoading(false);
      }
    })();
  }, [sp]);

  useEffect(() => {
    async function fetchUserInfo() {
      try {
        // const email = context.pageContext.user.email.toLowerCase();
        const userInfo = getEffectiveUser(context, effectiveUserLogin);
        const email = userInfo.email.toLowerCase();
        setCurrentUserEmail(email);

        const userGroups = await sp.web.currentUser.groups();
        const groupsLower = userGroups.map(g => g.Title.toLowerCase());
        const isAdmin = groupsLower.includes("admin");
        setIsAdminUser(isAdmin);
      } catch (error) {
        console.error('Error fetching user info:', error);
        setIsAdminUser(false);
      }
    }
    fetchUserInfo();
  }, [context, sp]);

  useEffect(() => {
    async function fetchProjects() {
      try {
        const projects = await projectsp.web.lists.getByTitle("Projects")
          .items
          .select(
            "Id", "Title",
            "POID/Id",
            "PM/EMail",
            "DM/EMail",
            "DH/EMail"
          )
          .expand("POID", "PM", "DM", "DH")
          .top(500)();

        setAllProjects(projects);
      } catch (error) {
        console.error('Error loading projects:', error);
        setAllProjects([]);
      }
    }
    fetchProjects();
  }, [sp]);

  function getEffectiveUser(context: any, effectiveUserLogin?: string) {
    const login = effectiveUserLogin || context.pageContext.user.loginName;
    const email = effectiveUserLogin || context.pageContext.user.email;
    const displayName = effectiveUserLogin || context.pageContext.user.displayName;

    return { login, email, displayName };
  }

  // Update your existing filteredMainPOs to include column filters
  const filteredMainPOs = React.useMemo(() => {
    const searchText = filters.search?.toLowerCase() || '';

    return mainPOs.filter(po => {
      // Global search
      const matchesSearch = !searchText || columns.some(col => {
        const fieldName = col.fieldName;
        if (!fieldName) return false;
        const fieldValue = (po as any)[fieldName];
        return fieldValue?.toString().toLowerCase().includes(searchText);
      });

      const matchesColumnFilters = Object.entries(columnFilters).every(
        ([colKey, selectedVals]) => {
          if (!selectedVals || selectedVals.length === 0) return true;
          const col = columns.find(c => c.key === colKey);
          if (!col || !col.fieldName) return true;
          const value = (po as any)[col.fieldName];
          if (value === null || value === undefined || value === '') return false;
          const vStr = value.toString();
          return selectedVals.includes(vStr);
        }
      );

      if (invoiceStatusFilter) {
        const percent = calculateInvoicedPercent(
          po.POID,
          parseFloat(po.POAmount || "0")
        );
        const epsilon = 0.0001;

        // NotPaid: ≈ 0%
        if (invoiceStatusFilter === "NotPaid" && Math.abs(percent) > epsilon) {
          return false;
        }

        // PartiallyInvoiced: strictly between 0 and 100
        if (
          invoiceStatusFilter === "PartiallyInvoiced" &&
          (percent <= epsilon || percent >= 100 - epsilon)
        ) {
          return false;
        }

        // CompletelyInvoiced: ≈ 100%
        if (
          invoiceStatusFilter === "CompletelyInvoiced" &&
          Math.abs(percent - 100) > epsilon
        ) {
          return false;
        }
      }

      // Existing project access filter
      if (isAdminUser) return matchesSearch && matchesColumnFilters;
      if (!po.ProjectName) return false;
      const project = allProjects.find((p: any) => p.Title === po.ProjectName);
      if (!project) return false;
      const userEmail = currentUserEmail.toLowerCase();
      const isUserPM = project.PM?.EMail?.toLowerCase() === userEmail;
      const isUserDM = project.DM?.EMail?.toLowerCase() === userEmail;
      const isUserDH = project.DH?.EMail?.toLowerCase() === userEmail;
      const isInPMGroup = userGroups.includes('pm');
      const isInDMGroup = userGroups.includes('dm');
      const isInDHGroup = userGroups.includes('dh');

      return (isInPMGroup || isUserPM || isInDMGroup || isUserDM || isInDHGroup || isUserDH) &&
        matchesSearch && matchesColumnFilters;
    });
  }, [mainPOs, filters.search, columnFilters, invoiceStatusFilter, allProjects, currentUserEmail, userGroups, columns]);

  const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
    // const isAmountField = ['POItemx0020Value', 'InvoiceAmount'].includes(columnKey)
    const isAmountField = ['POAmount', 'Amount'].includes(columnKey);

    const sortedItems = [...filteredMainPOs].sort((a: any, b: any) => {
      let aVal = a[columnKey]
      let bVal = b[columnKey]

      // EMPTY/NULL FIRST in ASC (0 first for numbers)
      if (aVal === null || aVal === undefined || aVal === '') {
        return direction === 'asc' ? -1 : 1
      }
      if (bVal === null || bVal === undefined || bVal === '') {
        return direction === 'asc' ? 1 : -1
      }

      // NUMERIC FIELDS - 0 first in ASC
      if (isAmountField) {
        const aNum = Number(aVal) || 0
        const bNum = Number(bVal) || 0
        return direction === 'asc' ? aNum - bNum : bNum - aNum
      }

      // DATES
      if (aVal instanceof Date) aVal = aVal.getTime()
      if (bVal instanceof Date) bVal = bVal.getTime()
      const aAsDate = Date.parse(aVal as any)
      const bAsDate = Date.parse(bVal as any)
      if (!isNaN(aAsDate) && !isNaN(bAsDate)) {
        return direction === 'asc' ? aAsDate - bAsDate : bAsDate - aAsDate
      }

      // STRINGS (default)
      const aStr = aVal?.toString() ?? ''
      const bStr = bVal?.toString() ?? ''
      return direction === 'asc' ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr)
    })

    setMainPOs(sortedItems);
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  }
  // Get distinct values for column filter
  const getColumnDistinctValues = (columnKey: string): string[] => {
    const col = columns.find(c => c.key === columnKey);
    if (!col || !col.fieldName) return [];
    const field = col.fieldName;

    const values = Array.from(
      new Set(
        mainPOs
          .map(i => (i as any)[field])
          .filter(v => v !== null && v !== undefined && v !== '')
          .map(v => v.toString())
      )
    );
    return values.sort((a, b) => a.localeCompare(b));
  };

  // Column visibility management
  const getVisibleColumns = (): IColumn[] => {
    return columns
      .filter(col => visibleColumns.includes(col.key as string))
      .sort((a, b) => {
        const aOrder = columnOrder[a.key as string] ?? visibleColumns.indexOf(a.key as string);
        const bOrder = columnOrder[b.key as string] ?? visibleColumns.indexOf(b.key as string);
        return aOrder - bOrder;
      });
  };

  const toggleColumnVisibility = (columnKey: string) => {
    setVisibleColumns(prev =>
      prev.includes(columnKey)
        ? prev.filter(k => k !== columnKey)
        : [...prev, columnKey]
    );
  };

  const moveColumn = (columnKey: string, direction: 'up' | 'down') => {
    const currentIndex = visibleColumns.indexOf(columnKey);
    if (direction === 'up' && currentIndex > 0) {
      const newOrder = [...visibleColumns];
      [newOrder[currentIndex - 1], newOrder[currentIndex]] = [newOrder[currentIndex], newOrder[currentIndex - 1]];
      setVisibleColumns(newOrder);
    } else if (direction === 'down' && currentIndex < visibleColumns.length - 1) {
      const newOrder = [...visibleColumns];
      [newOrder[currentIndex], newOrder[currentIndex + 1]] = [newOrder[currentIndex + 1], newOrder[currentIndex]];
      setVisibleColumns(newOrder);
    }
  };

  const clearColumnFilter = (columnKey: string) => {
    setColumnFilters(prev => {
      const newFilters = { ...prev };
      delete newFilters[columnKey];
      return newFilters;
    });
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  };

  function decodeHtmlEntities(str: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = str;
    return txt.value;
  }

  function formatCommentHistory(historyJson?: string) {
    let arr = [];
    try {
      if (!historyJson) return "";

      // Decode any HTML entities to get valid JSON string
      const decodedJson = decodeHtmlEntities(historyJson);

      arr = JSON.parse(decodedJson);
    } catch {
      arr = [];
    }
    if (!Array.isArray(arr)) arr = [];

    return arr
      .map((entry: any) => {
        const dateObj = entry.Date ? new Date(entry.Date) : null;
        const dateStr = dateObj ? dateObj.toLocaleDateString('en-GB') : '';
        const timeStr = dateObj ? dateObj.toLocaleTimeString('en-US', { hour: 'numeric', minute: 'numeric', second: 'numeric', hour12: true }) : '';
        const user = entry.User || 'Unknown User';
        const role = entry.Role ? ` (${entry.Role})` : '';
        const title = entry.Title || '';
        const comment = entry.Data || '';

        // Format as: [date time] user (role) - title:\ncomment
        return `[${dateStr} ${timeStr}]${user}${role} \n${title}: ${comment}`;
      })
      .join('\n\n'); // two line breaks between entries
  }

  const handleOpenPanel = async () => {
    if (!selectedItem) return;
    setInvoiceCurrency(selectedItem.Currency || "");
    setFetchingChildPOs(true);
    setFetchingInvoices(true);
    setChildPOItems([]);
    setInvoiceRequests([]);
    setActivePOIDFilter(selectedItem?.POID || null);
    setIsPanelOpen(false);
    setIsInvoicePanelOpen(false);
    setIsReadOnlyInvoicePanel(false);

    try {
      const allSalesRecords = await fetchAllPORecords();
      console.log("Fetched all sales records:", allSalesRecords);

      const rec = allSalesRecords.find((r: any) => findPOIndex(r, selectedItem.POID) !== null);

      if (!rec) {

        const invoices = await fetchInvoiceRequests(sp, [selectedItem.POID]);
        setInvoiceRequests(invoices);
        await handleOpenInvoicePanelSinglePO(selectedItem, "");
        return;
      }
      const poIndex = findPOIndex(rec, selectedItem.POID)!;
      let children = getChildItemsForPO(rec, poIndex);
      console.log("Child POs from JSON or ChildPO fields:", children);
      if (!children.length) {
        children = findChildPOsByParentPOID(allSalesRecords, selectedItem.POID);
        console.log("Child POs found by scanning ParentPOID columns:", children);
      }
      setChildPOItems(children);
      // const poids = children.length > 0 ? [selectedItem.POID, ...children.map(c => c.POID)] : [selectedItem.POID];
      const invoices = await fetchInvoiceRequests(sp, [selectedItem.POID]);
      setInvoiceRequests(invoices);

      const poamount = poIndex !== null ? (rec[`POAmount${poIndex === 0 ? "" : poIndex + 1}`] || "") : "";

      if (children.length === 0) {
        await handleOpenInvoicePanelSinglePO(selectedItem, poamount);
      } else {
        setIsPanelOpen(true);
      }
    } catch (err) {
      console.error("Error during PO handling", err);
    } finally {
      setFetchingChildPOs(false);
      setFetchingInvoices(false);
    }
  };

  const renderValue = (value: any) =>
    value !== null && value !== undefined
      ? value
      : <span style={{ color: '#999' }}>—</span>;

  useEffect(() => {
    async function fetchUserGroups() {
      try {
        const groups = await sp.web.currentUser.groups();
        setUserGroups(groups.map(g => g.Title.toLowerCase()));
      } catch (error) {
        console.error('Error fetching user groups:', error);
        setUserGroups([]);
      }
    }
    fetchUserGroups();
  }, [sp]);

  useEffect(() => {
    async function loadInvoiceRequests() {
      const poids = mainPOs.map(po => po.POID); // or whatever source POIDs you have
      const invoices = await fetchInvoiceRequests(sp, poids);
      setInvoiceRequests(invoices);
    }
    loadInvoiceRequests();
  }, [mainPOs]);

  useEffect(() => {
    const style = document.createElement('style');
    style.innerHTML = '[class*="contentContainer-"]';
    document.head.appendChild(style);
    return () => { document.head.removeChild(style); };
  }, []);


  const handleOpenInvoicePanelSinglePO = async (poItem: PurchaseOrderItem, poAmount: string) => {
    setInvoicePanelPO(null);
    setIsInvoicePanelOpen(true);
    setInvoiceCurrency(poItem.Currency || "");

    setInvoiceFormState({
      POID: poItem.POID,
      PurchaseOrder: poItem.POID,
      ProjectName: poItem.ProjectName,
      POItemTitle: "",
      POItemValue: "",
      InvoiceAmount: "",
      POAmount: poAmount,
      CustomerContact: "",
      Comments: "",
      Attachment: null,
      CostCenter: poItem.CostCenter || "",
    });
  };
  const handlePanelDismiss = () => {
    setIsPanelOpen(false);
    setChildPOItems([]);
    setInvoiceRequests([]);
    setSelectedItem(null);
    selection.setAllSelected(false);
    childPOSelection.setAllSelected(false);
    setIsReadOnlyInvoicePanel(false);
    window.history.replaceState(null, '', window.location.pathname);
  };
  const handleChildPORowClick = (item?: ChildPOItem) => {
    if (item) {
      console.log(item)
      setActivePOIDFilter(item.POID);
      setFilterMode('childPO');
      // childPOSelection.setKeySelected(item.Id.toString(), true, false);
    }
  };
  const filteredInvoiceRequests = activePOIDFilter
    ? invoiceRequests.filter(ir => {
      if (!ir.InvoicedAmountsJSON) return false;

      let rows: any[] = [];
      try {
        const decoded = decodeHtmlEntities(ir.InvoicedAmountsJSON);
        rows = JSON.parse(decoded);
      } catch {
        return false;
      }

      return Array.isArray(rows) &&
        rows.some(r => r.poItemTitle === activePOIDFilter && r.invoicedAmount > 0);
    })
    : invoiceRequests;

  const showInvoices =
    filterMode === 'mainPO'
      ? invoiceRequests.filter(ir => ir.PurchaseOrderPO === activePOIDFilter)
      : filteredInvoiceRequests;


  const handleInvoiceFormChange = (field: keyof InvoiceFormState, value: any) => {
    setInvoiceFormState((prev) => ({
      ...prev,
      [field]: value,
    }));
  };

  async function sendMailWithGraph(graphClient: MSGraphClient, to: string | string[], subject: string, body: string): Promise<void> {
    const recipients = (Array.isArray(to) ? to : [to]).map(address => ({
      emailAddress: { address }
    }));

    const mail = {
      message: {
        subject,
        body: {
          contentType: 'HTML',
          content: body
        },
        toRecipients: recipients
      }
    };
    await graphClient.api('/me/sendMail').post(mail);
  }

  function getRemainingPOAmount(
    childPO: ChildPOItem,
    invoiceRequests: InvoiceRequest[]
  ): number {
    const poId = selectedItem?.POID?.trim() ?? '';
    const poItemId = childPO.POID.trim();

    let used = 0;

    for (const ir of invoiceRequests) {
      if (
        ir.PurchaseOrderPO?.trim() !== poId ||
        (ir.CurrentStatus ?? '').toLowerCase() === 'cancelled'
      ) continue;

      if (!ir.InvoicedAmountsJSON) continue;

      try {
        const parsed = JSON.parse(ir.InvoicedAmountsJSON);
        if (!Array.isArray(parsed)) continue;

        const match = parsed.find(
          (r: any) => r.poItemTitle?.trim() === poItemId
        );

        if (match && !isNaN(match.invoicedAmount)) {
          used += Number(match.invoicedAmount);
        }
      } catch (e) {
        console.warn('Invalid InvoicedAmountsJSON', e);
      }
    }

    const original = Number(childPO.POAmount) || 0;
    return Math.max(0, original - used);
  }

  function calculateInvoicedPercent(
    poId: string,
    mainPOAmount: number | string
  ): number {
    const amount =
      typeof mainPOAmount === 'string'
        ? Number(mainPOAmount.replace(/[^\d.]/g, ''))
        : Number(mainPOAmount)

    if (!amount) return 0

    const { invoiced } = getTotalsFromJsonForPO(poId, invoiceRequests)
    return Math.round((invoiced / amount) * 100)
  }


  function calculateInvoicedPercentForItem(
    poItemPOID: string,
    poItemAmount: number
  ): number {
    if (!poItemAmount) return 0

    const { invoiced } = getTotalsFromJsonForPOItem(
      selectedItem?.POID ?? '',
      poItemPOID,
      invoiceRequests
    )

    return Math.round((invoiced / poItemAmount) * 100)
  }

  const handleOpenInvoicePanel = async (item: ChildPOItem) => {
    if (!selectedItem) return;
    setInvoicePanelLoading(true);
    setInvoicePanelPO(item);
    setIsInvoicePanelOpen(true);
    setInvoiceCurrency(selectedItem.Currency || "");

    const parentPOID = selectedItem.POID;
    const parentPOIDId = selectedItem.Id;
    const projectName = await getProjectNameByPOID(context, parentPOIDId, selectedItem);
    initMultiLineAllocations(childPOItems)

    setInvoiceFormState({
      POID: parentPOID,
      PurchaseOrder: parentPOID,
      ProjectName: projectName,
      POItemTitle: item.POID,
      POAmount: item.POAmount,
      POItemValue: item.POAmount,
      InvoiceAmount: "",
      CustomerContact: "",
      Comments: "",
      CostCenter: selectedItem.CostCenter || "",
      Attachment: null,
    });
    setInvoicePanelLoading(false);
  };
  function findPOIndex(record: any, poidToFind: string): number | null {
    for (let i = 0; i < 26; i++) {
      const fieldName = i === 0 ? "POID" : `POID${i + 1}`;
      const value = record[fieldName];
      if (value && value.toString().trim() === poidToFind.trim()) {
        return i;
      }
    }
    return null;
  }
  // const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
  //   if (!props) return null;
  //   return (
  //     <Sticky stickyPosition={StickyPositionType.Header}>
  //       {defaultRender!({ ...props })}
  //     </Sticky>
  //   );
  // };
  const onRenderHeaderGeneric: IRenderFunction<IDetailsHeaderProps> = (props: any, defaultRender: any) => {
    if (!props || !defaultRender) return null;
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        {defaultRender({
          ...props,
          onColumnClick: (ev?: React.MouseEvent<HTMLElement>, column?: IColumn) => {
            if (!column || !ev) return;

            // Use the correct setMenuState for THIS list
            setMainColumnFilterMenu({  // ← Change per list
              visible: true,
              target: ev.currentTarget as HTMLElement,
              columnKey: column.key as string,
            });
          },
        })}
      </Sticky>
    );
  };


  // Utility to approximate pixel width of text (rough)
  function estimateWidth(text: string) {
    return Math.min(300, Math.max(50, text.length * 10)); // min 50px, max 300px
  }
  // Compute column widths given items and columns config
  function computeColumnWidths(items: any[], columns: IColumn[]): IColumn[] {
    return columns.map(col => {
      const headerWidth = estimateWidth(col.name);
      const maxItemLength = items.reduce((max, item) => {
        const val = item[col.fieldName as keyof typeof item];
        const valStr = val !== undefined && val !== null ? String(val) : "";
        return Math.max(max, valStr.length);
      }, 0);
      const dataWidth = estimateWidth('W'.repeat(maxItemLength));
      const width = Math.max(headerWidth, dataWidth);

      return {
        ...col,
        minWidth: width,
        maxWidth: width,
        isResizable: true,
      };
    });
  }

  const handleFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.from(e.target.files) : [];
    setUploadedFiles(prev => [...prev, ...files]);
    setInvoiceFormState(prev => ({ ...prev, Attachment: files[0] }));
  };

  const handleDropMulti = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragActive(false);
    const files = Array.from(e.dataTransfer.files);
    setUploadedFiles(prev => [...prev, ...files]);
    setInvoiceFormState(prev => ({ ...prev, Attachment: files[0] }));
  };

  const removeAttachment = (idx: number) => {
    setUploadedFiles(prev => {
      const updated = prev.filter((_, i) => i !== idx);
      setInvoiceFormState(form => ({
        ...form,
        Attachment: updated[0] ?? null // sets to first file, or null if none left
      }));
      return updated;
    });
  };

  function getChildItemsForPO(record: any, index: number): ChildPOItem[] {
    const childPOField = index === 0 ? "ParentPOID" : `ParentPOID${index + 1}`;
    const poAmountField = index === 0 ? "POAmount" : `POAmount${index + 1}`;
    const lineItemsField = index === 0 ? "LineItemsJSON" : `LineItemsJSON${index + 1}`;

    if (record[childPOField] && record[childPOField].toString().trim() !== "") {
      return [{
        Id: record.Id,
        POID: record[childPOField],
        POAmount: record[poAmountField] || "",
        ParentPOIndex: index + 1,
        POIndex: index + 1
      }];
    }

    if (record[lineItemsField]) {
      try {
        // 🔑 decode RichText first
        const decoded = decodeHtmlEntities(record[lineItemsField]);
        const items = JSON.parse(decoded);
        if (Array.isArray(items)) {
          return items.map((item: any, i: number) => ({
            Id: i + 1,
            POID: item.Title || `LineItem${i + 1}`,   // use Title instead of POID
            POAmount: item.Value?.toString() || "0", // use Value instead of POAmount
            ParentPOIndex: 0,
            POIndex: 0,
          }));
        }
      } catch (err) {
        console.warn("Error parsing LineItemsJSON:", err, record[lineItemsField]);
      }
    }

    return [];
  }

  function findChildPOsByParentPOID(allRecords: any[], poid: string): ChildPOItem[] {
    const parentPOIDCols = Array.from({ length: 25 }, (_, i) => (i === 0 ? "ParentPOID" : `ParentPOID${i + 1}`));
    const poIDCols = Array.from({ length: 25 }, (_, i) => (i === 0 ? "POID" : `POID${i + 1}`));
    const poAmountCols = Array.from({ length: 25 }, (_, i) => (i === 0 ? "POAmount" : `POAmount${i + 1}`));
    const childPOs: ChildPOItem[] = [];

    allRecords.forEach(record => {
      for (let idx = 0; idx < parentPOIDCols.length; idx++) {
        const parentVal = record[parentPOIDCols[idx]];
        if (parentVal && parentVal.toString().trim() === poid.trim()) {
          const childPOID = record[poIDCols[idx]] || "";
          const childAmount = record[poAmountCols[idx]] || "";
          childPOs.push({
            Id: record.Id,
            POID: childPOID,
            POAmount: childAmount,
            ParentPOIndex: idx + 1,
            POIndex: idx + 1,
          });
        }
      }
    });

    return childPOs;
  }
  // function isPoItemUsedInAnyInvoice(
  //   parentPOID: string,
  //   poItemTitle: string,
  //   allRequests: InvoiceRequest[]
  // ): boolean {
  //   const normalizedParent = parentPOID.trim().toLowerCase();
  //   const normalizedItem = poItemTitle.trim().toLowerCase();

  //   const relevant = allRequests.filter(
  //     ir =>
  //       ir.PurchaseOrderPO?.trim().toLowerCase() === normalizedParent &&
  //       !!ir.InvoicedAmountsJSON
  //   );

  //   for (const ir of relevant) {
  //     try {
  //       const rows = JSON.parse(ir.InvoicedAmountsJSON!);
  //       if (!Array.isArray(rows)) continue;
  //       const match = rows.some((r: any) =>
  //         (r.poItemTitle ?? r.POItemTitle ?? "")
  //           .toString()
  //           .trim()
  //           .toLowerCase() === normalizedItem
  //       );
  //       if (match) return true;
  //     } catch {
  //       continue;
  //     }
  //   }
  //   return false;
  // }

  async function fetchAllPORecords(): Promise<any[]> {
    // const remote = spfi(PROJECTS_SITE_URL).using(SPFx(context));
    const allItems: any[] = [];

    try {
      // Fetch all items, selecting only relevant InvoicePO fields
      const pagedItems = sp.web.lists.getByTitle("InvoicePO")
        .items
        .select("ID", "POID", "ParentPOID", "POAmount", "LineItemsJSON", "ProjectName", "CostCenter")
        .top(500); // you can adjust top count or implement paging if needed

      // Use async iterator to get all pages
      for await (const batch of pagedItems) {
        allItems.push(...batch);
      }
    } catch (err) {
      console.error("Error fetching InvoicePO records:", err);
    }
    return allItems;
  }

  async function fetchInvoiceRequests(sp: SPFI, poIds: string[]): Promise<InvoiceRequest[]> {
    const validPoIds = poIds.filter(po => po && po.toLowerCase() !== "null");
    if (validPoIds.length === 0) return [];

    const batchSize = 15;                        // tune as needed
    const allItems: any[] = [];
    try {
      for (let i = 0; i < validPoIds.length; i += batchSize) {
        const chunk = validPoIds.slice(i, i + batchSize)
          .map(po => po.replace(/'/g, "''"));
        const filter = `(${chunk.map(po => `PurchaseOrder eq '${po}'`).join(" or ")})`;

        const items = await sp.web.lists
          .getByTitle("Invoice Requests")
          .items.filter(filter)
          .select(
            "Id",
            "PurchaseOrder",
            "InvoiceAmount",
            "Comments",
            "Customer_x0020_Contact",
            "POItem_x0020_Title",
            "POItem_x0020_Value",
            "ProjectName",
            "Status",
            "PMCommentsHistory",
            "FinanceCommentsHistory",
            "Created",
            "Author/Title",
            "Modified",
            "Editor/Title",
            "CurrentStatus",
            "InvoicedAmountsJSON",
          )
          .expand("Author", "Editor")();
        allItems.push(...items);
      }

      return allItems.map(item => ({
        Id: item.Id,
        PurchaseOrderPO: item.PurchaseOrder,
        PurchaseOrder: item.PurchaseOrder,
        Amount: item.InvoiceAmount,
        Status: item.Status,
        ProjectName: item.ProjectName,
        POItemTitle: item.POItem_x0020_Title,
        POItemValue: item.POItem_x0020_Value,
        CustomerContact: item.Customer_x0020_Contact,
        Comments: item.Comments,
        PMCommentsHistory: item.PMCommentsHistory,
        FinanceCommentsHistory: item.FinanceCommentsHistory,
        Created: item.Created,
        CreatedBy: item.Author?.Title ?? "",
        Modified: item.Modified,
        ModifiedBy: item.Editor?.Title ?? "",
        CurrentStatus: item.CurrentStatus,
        InvoicedAmountsJSON: item.InvoicedAmountsJSON,
      }));
    } catch (err) {
      console.error("Error in fetchInvoiceRequests:", err);
    }
  }

  async function getCurrentUserRole(context: any, poId: any): Promise<string> {
    try {
      // const sp = spfi(PROJECTS_SITE_URL).using(SPFx(context));
      const userInfo = getEffectiveUser(context, effectiveUserLogin);
      const currentUserEmail = userInfo.email.toLowerCase();
      // const currentUserEmail = context.pageContext.user.email.toLowerCase();

      const projects = await projectsp.web.lists.getByTitle("Projects")
        .items
        .filter(`Title eq '${poId?.ProjectName?.replace(/'/g, "''")}'`)
        .select(
          "POID/Id",
          "PM/EMail",
          "DM/EMail",
          "DH/EMail",
        )
        .expand("POID", "PM", "DM", "DH")
        ();

      if (isAdminUser) {
        return "Admin";
      }
      const matchedProject = projects[0];
      if (!matchedProject) {
        return "";
      }
      if (matchedProject.DH?.EMail.toLowerCase() === currentUserEmail) return "DH";
      if (matchedProject.DM?.EMail.toLowerCase() === currentUserEmail) return "DM";
      if (matchedProject.PM?.EMail.toLowerCase() === currentUserEmail) return "PM";
      return "";
    } catch (error) {
      console.error("Error determining user role:", error);
      return "";
    }
  }

  function getCurrencySymbol(currencyCode: string, locale = "en-US") {
    if (!currencyCode || currencyCode.trim() === "") {
      // Return a default symbol or empty string if no currency code provided
      return "USD";
    }
    try {
      return new Intl.NumberFormat(locale, {
        style: "currency",
        currency: currencyCode,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0,
      })
        .formatToParts(1)
        .find(part => part.type === "currency")?.value ?? currencyCode;
    } catch (error) {
      // Fallback if currency code invalid
      console.warn(`Invalid currency code: ${currencyCode}`, error);
      return currencyCode;
    }
  }

  const initMultiLineAllocations = (items: ChildPOItem[]) => {
    const rows = items.map((c, idx) => {
      const remaining = getRemainingPOAmount(
        { POID: c.POID, POAmount: c.POAmount, Id: c.Id, ParentPOIndex: c.ParentPOIndex, POIndex: c.POIndex },
        invoiceRequests
      );
      return {
        poItemId: c.POID,
        title: `Line item ${idx + 1}`,
        poItemValue: Number(c.POAmount) || 0,
        remaining,
        invoiceAmount: 0
      };
    });
    setLineAllocations(rows);
    setInvoiceFormState(f => ({ ...f, InvoiceAmount: "0" }));
  };

  const handleLineAmountChange = (index: number, value: string) => {
    setLineAllocations(prev => {
      const copy = [...prev];
      const parsed = parseFloat(value);

      copy[index].invoiceAmount = !value || isNaN(parsed) ? 0 : parsed;

      if (!value || parsed <= 0) {
        copy[index].error = "Please enter a valid number.";
      } else if (parsed > copy[index].remaining) {
        copy[index].error =
          `Invoiced Amount cannot exceed remaining amount: ${copy[index].remaining}`;
      } else {
        copy[index].error = undefined;
      }

      const total = copy.reduce((s, r) => s + (r.invoiceAmount || 0), 0);
      setInvoiceFormState(f => ({ ...f, InvoiceAmount: total.toString() }));

      return copy;
    });
  };

  function getTotalsFromJsonForPOItem(
    poId: string,
    poItemTitle: string,
    invoiceRequests: InvoiceRequest[]
  ) {
    let requested = 0
    let invoiced = 0
    let paid = 0

    const normalizedPoId = poId.trim().toLowerCase()
    const normalizedItem = poItemTitle.trim().toLowerCase()

    const relevant = invoiceRequests.filter(ir =>
      ir.PurchaseOrderPO &&
      ir.PurchaseOrderPO.trim().toLowerCase() === normalizedPoId &&
      (ir.CurrentStatus ?? '').toLowerCase() !== 'cancelled' &&
      !!ir.InvoicedAmountsJSON
    )

    for (const ir of relevant) {
      let rows: any[] = []

      try {
        rows = JSON.parse(ir.InvoicedAmountsJSON!)
      } catch {
        continue
      }

      if (!Array.isArray(rows)) continue

      const match = rows.find(row =>
        (row.poItemTitle ?? row.POItemTitle ?? '')
          .trim()
          .toLowerCase() === normalizedItem
      )

      if (!match) continue

      const amount = Number(
        match.invoicedAmount ??
        match.InvoiceAmount ??
        match.Value ??
        0
      )

      if (!amount || isNaN(amount)) continue

      const status = (ir.Status ?? '').toLowerCase()
      const currentStatus = (ir.CurrentStatus ?? '').toLowerCase()

      if (
        status === 'invoice requested' ||
        currentStatus === 'request submitted'
      ) {
        requested += amount
      } else if (
        status === 'invoice raised' ||
        status === 'pending payment'
      ) {
        invoiced += amount
      } else if (status === 'payment received') {
        paid += amount
      }
    }

    return { requested, invoiced, paid }
  }

  function getTotalsFromJsonForPO(
    poId: string,
    invoiceRequests: InvoiceRequest[]
  ) {
    let requested = 0
    let invoiced = 0
    let paid = 0

    for (const ir of invoiceRequests) {
      if (
        ir.PurchaseOrderPO?.trim() !== poId.trim() ||
        !ir.InvoicedAmountsJSON ||
        (ir.CurrentStatus ?? '').toLowerCase() === 'cancelled'
      ) continue

      let total = 0

      try {
        const rows = JSON.parse(ir.InvoicedAmountsJSON)
        if (!Array.isArray(rows)) continue

        total = rows.reduce((s: number, r: any) => {
          const amt = Number(r.invoicedAmount ?? r.Value ?? 0)
          return !isNaN(amt) ? s + amt : s
        }, 0)
      } catch {
        continue
      }

      const status = (ir.Status ?? '').toLowerCase()
      const current = (ir.CurrentStatus ?? '').toLowerCase()

      if (
        status === 'invoice requested' ||
        current === 'request submitted'
      ) {
        requested += total
      } else if (
        status === 'invoice raised' ||
        status === 'pending payment'
      ) {
        invoiced += total
      } else if (status === 'payment received') {
        paid += total
      }
    }
    return { requested, invoiced, paid }
  }

  function getTotalsFromJsonForInvoice(ir: InvoiceRequest) {
    let requested = 0
    let invoiced = 0
    let paid = 0
    let cancelled = 0

    if (!ir.InvoicedAmountsJSON) {
      return { requested, invoiced, paid, cancelled }
    }

    let rows: any[] = []

    try {
      rows = JSON.parse(ir.InvoicedAmountsJSON)
    } catch {
      return { requested, invoiced, paid, cancelled }
    }

    if (!Array.isArray(rows)) {
      return { requested, invoiced, paid, cancelled }
    }

    const lineTotal = rows.reduce((sum, r) => {
      const amt = Number(
        r.invoicedAmount ??
        r.InvoiceAmount ??
        r.Value ??
        0
      )
      return !isNaN(amt) ? sum + amt : sum
    }, 0)

    const status = (ir.Status ?? '').toLowerCase()
    const currentStatus = (ir.CurrentStatus ?? '').toLowerCase()

    if (
      status === 'invoice requested' ||
      currentStatus === 'request submitted'
    ) {
      requested = lineTotal
    } else if (
      status === 'invoice raised' ||
      status === 'pending payment' || status === 'overdue'
    ) {
      invoiced = lineTotal
    } else if (status === 'payment received') {
      paid = lineTotal
    } else if (status === 'cancelled') {
      cancelled = lineTotal
    }

    return { requested, invoiced, paid, cancelled }
  }

  async function getProjectNameByPOID(context: any, poId: number, poItem: any): Promise<string> {
    try {
      const getAllProjectsq = async (): Promise<any[]> => {
        const allItems: any[] = [];

        const pagedItems = projectsp.web.lists.getByTitle("Projects")
          .items
          .select(
            "Id", "Title",
            "POID/Id", "POID/Title",
            "PM/Id", "PM/Title", "PM/EMail",
            "DM/Id", "DM/Title", "DM/EMail",
            "DH/Id", "DH/Title", "DH/EMail"
          )
          .expand("POID", "PM", "DM", "DH")
          .top(100);

        for await (const batch of pagedItems) {
          allItems.push(...batch);
        }

        return allItems;
      };

      const allItems = await getAllProjectsq();
      const projectNameToMatch = poItem?.ProjectName?.trim().toLowerCase();
      const matchedItem = allItems.find(item =>
        item.Title && item.Title.trim().toLowerCase() === projectNameToMatch
      );

      return matchedItem && matchedItem.Title ? matchedItem.Title : "";

    } catch (error) {
      console.error("Error fetching projects or filtering:", error);
      return "";
    }
  }

  const amountLabel = React.useMemo(() => {
    const status = selectedInvoiceRequest?.Status?.toLowerCase() ?? "";
    const current = selectedInvoiceRequest?.CurrentStatus?.toLowerCase() ?? "";

    // Align with your JSON-total helpers logic
    if (status === "invoice requested" || current === "request submitted") {
      return "Requested Amount";
    }
    if (status === "invoice raised" || status === "pending payment" || status === "overdue") {
      return "Invoiced Amount";
    }
    if (status === "payment received") {
      return "Paid Amount";
    }
    if (status === "cancelled") {
      return "Cancelled Amount";
    }
    return "Amount";
  }, [selectedInvoiceRequest]);

  const handleInvoiceFormSubmit = async () => {
    // if (isSubmitting) return;
    if (isSubmitting || submitDialogState !== 'idle') return;

    setSubmitDialogState('submitting');
    setIsSubmitting(true);

    let addedItemId: number | null = null;
    try {
      if (invoiceAmountError || !invoiceFormState.InvoiceAmount) {
        setDialogType('error');
        setDialogMessage(invoiceAmountError || "Invoiced Amount is required.");
        setDialogVisible(true);
        setIsSubmitting(false);
        return;
      }
      const userRole = await getCurrentUserRole(context, selectedItem);

      const financeStatusValue = "Invoice Requested";
      const nonZeroLines = lineAllocations.filter(l => l.invoiceAmount > 0)
      const total = nonZeroLines.reduce((s, r) => s + r.invoiceAmount, 0)

      if (total === 0) {
        setInvoiceAmountError('Enter at least one line amount.')
        setIsSubmitting(false)
        return
      }

      // child PO + single PO cases
      const allPOItemsData = childPOItems.length > 0
        ? childPOItems.map((childPO: ChildPOItem) => ({
          poItemTitle: childPO.POID,
          poItemValue: Number(childPO.POAmount) || 0,
          invoicedAmount:
            lineAllocations.find(la => la.poItemId === childPO.POID)?.invoiceAmount || 0
        }))
        : [{
          poItemTitle: invoiceFormState.POID,
          poItemValue: Number(invoiceFormState.POAmount) || 0,
          invoicedAmount: Number(invoiceFormState.InvoiceAmount) || 0
        }];

      if (invoiceFormState.Comments && invoiceFormState.Comments.trim().length > 0) {
        const userInfo = getEffectiveUser(context, effectiveUserLogin);
        console.log("User Info:", userInfo);
        const newCommentEntry = {
          Date: new Date().toISOString(),
          Title: "Comment",
          User: userInfo.displayName,
          Role: userRole,
          Data: invoiceFormState.Comments.trim()
        };

        const pmCommentsHistoryArray = [newCommentEntry];

        const eff = await sp.web.ensureUser(
          getEffectiveUser(context, effectiveUserLogin).login
        );
        const initialStatusHistory = JSON.stringify([
          {
            index: 1,
            status: 'Invoice Requested',
            date: new Date().toISOString(),
            user: context.pageContext.user.displayName || 'Admin',
          },
        ]);
        const added1 = await sp.web.lists.getByTitle("Invoice Requests").items.add({
          PurchaseOrder: invoiceFormState.POID,
          ProjectName: invoiceFormState.ProjectName,
          POAmount: invoiceFormState.POAmount ? Number(invoiceFormState.POAmount) : null,
          // POItem_x0020_Title: invoicePanelPO === null ? null : invoiceFormState.POItemTitle,
          // POItem_x0020_Value: invoicePanelPO === null ? null : (invoiceFormState.POItemValue ? Number(invoiceFormState.POItemValue) : null),
          InvoiceAmount: invoiceFormState.InvoiceAmount ? Number(invoiceFormState.InvoiceAmount) : null,
          Customer_x0020_Contact: invoiceFormState.CustomerContact,
          Comments: invoiceFormState.Comments,
          PMStatus: "Submitted",
          FinanceStatus: "Pending",
          Status: financeStatusValue,
          Currency: invoiceCurrency,
          CurrentStatus: `Request Submitted`,
          PMCommentsHistory: JSON.stringify(pmCommentsHistoryArray),
          RequestCreatedDate: new Date().toISOString(),
          RequestedCreatedById: eff.Id,
          InvoicedAmountsJSON: JSON.stringify(allPOItemsData),
          StatusHistory: initialStatusHistory,
        });
        addedItemId = added1.Id;

      } else {
        const eff = await sp.web.ensureUser(
          getEffectiveUser(context, effectiveUserLogin).login
        );
        const initialStatusHistory = JSON.stringify([
          {
            index: 1,
            status: 'Invoice Requested',
            date: new Date().toISOString(),
            user: context.pageContext.user.displayName || 'Admin',
          },
        ]);

        const added2 = await sp.web.lists.getByTitle("Invoice Requests").items.add({
          PurchaseOrder: invoiceFormState.POID,
          ProjectName: invoiceFormState.ProjectName,
          POAmount: invoiceFormState.POAmount ? Number(invoiceFormState.POAmount) : null,
          POItem_x0020_Title: invoicePanelPO === null ? null : invoiceFormState.POItemTitle,
          POItem_x0020_Value: invoicePanelPO === null ? null : (invoiceFormState.POItemValue ? Number(invoiceFormState.POItemValue) : null),
          InvoiceAmount: invoiceFormState.InvoiceAmount ? Number(invoiceFormState.InvoiceAmount) : null,
          Customer_x0020_Contact: invoiceFormState.CustomerContact,
          Comments: invoiceFormState.Comments,
          PMStatus: "Submitted",
          FinanceStatus: "Pending",
          Status: financeStatusValue,
          Currency: invoiceCurrency,
          CurrentStatus: `Request Submitted`,
          RequestCreatedDate: new Date().toISOString(),
          RequestedCreatedById: eff.Id,
          InvoicedAmountsJSON: JSON.stringify(allPOItemsData),
          StatusHistory: initialStatusHistory,
        });
        addedItemId = added2.Id;
      }

      if (invoicePanelPO === null && invoiceFormState.POItemValue) {
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).update({
          POAmount: Number(invoiceFormState.POItemValue)
        });
      }
      if (uploadedFiles.length > 0) {
        for (const file of uploadedFiles) {
          const fileExt = file.name.slice(file.name.lastIndexOf('.') + 1);
          const fileNameWithoutExt = file.name.slice(0, file.name.lastIndexOf('.'));
          const fileNameWithSuffix = `${fileNameWithoutExt}Requestor.${fileExt}`;
          const fileContent = await file.arrayBuffer();

          await sp.web.lists.getByTitle("Invoice Requests")
            .items.getById(addedItemId)
            .attachmentFiles.add(fileNameWithSuffix, fileContent);
        }
      }
      const financeConfigItems = await sp.web.lists.getByTitle("InvoiceConfiguration").items.filter("Title eq 'FinanceEmail'")();
      const financeEmails = financeConfigItems.length > 0 ? financeConfigItems[0].Value : "";

      const eff = getEffectiveUser(context, effectiveUserLogin);
      const creatorEmail = eff.email.toLowerCase();

      // const creatorEmail = context.pageContext.user.email;

      const siteUrl = context.pageContext.web.absoluteUrl;
      const pageName = context.pageContext.site.serverRequestPath.split('/').pop() || 'InvoiceTracker.aspx';
      const appPageUrl = `${siteUrl}/SitePages/${pageName}`;

      const itemLink = `${appPageUrl}#myrequests?selectedInvoice=${addedItemId}`;
      const siteTitle = context.pageContext.web.title;
      // ✅ Add these BEFORE your email templates
      // const totalInvoiceAmount = lineAllocations.reduce((sum, row) => sum + (row.invoiceAmount || 0), 0);
      const currencySymbol = getCurrencySymbol(invoiceCurrency || 'USD');

      // TypeScript: Compose HTML for "created user" email
      //       const createdUserEmailBody = `
      // <div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;background:#f9f9f9;border-radius:10px;padding:24px;">
      //   <div style="font-size:18px;font-weight:600;color:#0078d4;margin-bottom:16px;">
      //     Invoice Request Created
      //   </div>
      //   <div style="font-size:16px;color:#444;margin-bottom:18px;">
      //     Your new invoice request has been created and is now being tracked in the system.
      //   </div>
      //   <table style="width:100%;border-collapse:collapse;font-size:15px;color:#333;margin-bottom:20px;">
      //     <tr>
      //       <td style="font-weight:600;padding:6px 0;">PO ID:</td>
      //       <td>${invoiceFormState.POID}</td>
      //     </tr>
      //     <tr>
      //       <td style="font-weight:600;padding:6px 0;">Project Name:</td>
      //       <td>${invoiceFormState.ProjectName}</td>
      //     </tr>
      //     <tr>
      //       <td style="font-weight:600;padding:6px 0;">PO Item Title:</td>
      //       <td>${invoiceFormState.POItemTitle}</td>
      //     </tr>
      //     <tr>
      //       <td style="font-weight:600;padding:6px 0;">Comments:</td>
      //       <td>${invoiceFormState.Comments || "—"}</td>
      //     </tr>
      //   </table>
      //   <div style="margin-bottom:24px;">
      //     <a href="${itemLink}" style="font-size:15px;color:#0078d4;text-decoration:underline;">
      //       Click here to view the invoice request
      //     </a>
      //   </div>
      //   <div style="border-top:1px solid #eee;margin-top:22px;padding-top:10px;font-size:13px;color:#999;">
      //     Invoice Tracker | SACHA Group
      //   </div>
      // </div>
      // `;
      // const createdUserEmailBody = (
      //   <div style={{ fontFamily: 'Segoe UI,Arial,sans-serif', maxWidth: '600px', background: '#f9f9f9', borderRadius: '10px', padding: '24px' }}>
      //     <div style={{ fontSize: '18px', fontWeight: 600, color: '#0078d4', marginBottom: '16px' }}>
      //       Invoice Request Created
      //     </div>
      //     <div style={{ fontSize: '16px', color: '#444', marginBottom: '18px' }}>
      //       Your new invoice request has been created and is now being tracked in the system.
      //     </div>

      //     {/* ✅ NEW PO ITEMS TABLE */}
      //     <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '15px', color: '#333', marginBottom: '20px' }}>
      //       <thead>
      //         <tr style={{ background: '#f8f9fa' }}>
      //           <th style={{ padding: '12px 8px', textAlign: 'left', fontWeight: 600, borderBottom: '2px solid #dee2e6' }}>PO ID</th>
      //           <th style={{ padding: '12px 8px', textAlign: 'left', fontWeight: 600, borderBottom: '2px solid #dee2e6' }}>Project</th>
      //           <th style={{ padding: '12px 8px', textAlign: 'right', fontWeight: 600, borderBottom: '2px solid #dee2e6' }}>PO Items</th>
      //           <th style={{ padding: '12px 8px', textAlign: 'right', fontWeight: 600, borderBottom: '2px solid #dee2e6' }}>Total Amount</th>
      //         </tr>
      //       </thead>
      //       <tbody>
      //         <tr>
      //           <td style={{ padding: '12px 8px', fontWeight: 600, borderBottom: '1px solid #eee' }}>{invoiceFormState.POID}</td>
      //           <td style={{ padding: '12px 8px', borderBottom: '1px solid #eee' }}>{invoiceFormState.ProjectName}</td>
      //           <td style={{ padding: '12px 8px', textAlign: 'right', borderBottom: '1px solid #eee' }}>
      //             <table style={{ width: '100%', margin: '0 auto' }}>
      //               <tbody>
      //                 {allPOItemsData.map((item, idx) => (
      //                   <tr key={idx}>
      //                     <td style={{ padding: '4px 8px', fontSize: '13px' }}>
      //                       {item.poItemTitle}
      //                       <span style={{ fontSize: '11px', color: '#666' }}> ({currencySymbol}{item.poItemValue?.toLocaleString()})</span>
      //                     </td>
      //                     <td style={{ padding: '4px 8px', textAlign: 'right', fontSize: '13px', fontWeight: 600 }}>
      //                       {currencySymbol}{item.invoicedAmount?.toLocaleString()}
      //                     </td>
      //                   </tr>
      //                 ))}
      //                 <tr style={{ borderTop: '1px solid #ddd', marginTop: '4px' }}>
      //                   <td style={{ padding: '6px 8px', fontWeight: 700 }}>TOTAL</td>
      //                   <td style={{ padding: '6px 8px', textAlign: 'right', fontWeight: 700, fontSize: '14px' }}>
      //                     {currencySymbol}{totalInvoiceAmount.toLocaleString()}
      //                   </td>
      //                 </tr>
      //               </tbody>
      //             </table>
      //           </td>
      //           <td style={{ padding: '12px 8px', textAlign: 'right', fontWeight: 700, fontSize: '16px', color: '#28a745' }}>
      //             {currencySymbol}{Number(invoiceFormState.InvoiceAmount || 0).toLocaleString()}
      //           </td>
      //         </tr>
      //       </tbody>
      //     </table>

      //     {invoiceFormState.Comments && (
      //       <div style={{ marginBottom: '24px' }}>
      //         <strong>Comments:</strong>
      //         <div style={{ background: '#f8f9fa', padding: '12px', borderRadius: '6px', fontSize: '14px', marginTop: '8px' }}>
      //           {invoiceFormState.Comments}
      //         </div>
      //       </div>
      //     )}

      //     <div style={{ marginBottom: '24px' }}>
      //       <a href={itemLink} style={{ fontSize: '15px', color: '#0078d4', textDecoration: 'underline' }}>
      //         Click here to view the invoice request
      //       </a>
      //     </div>

      //     <div style={{ borderTop: '1px solid #eee', marginTop: '22px', paddingTop: '10px', fontSize: '13px', color: '#999' }}>
      //       Invoice Tracker - SACHA Group
      //     </div>
      //   </div>
      // );
      // ✅ FIXED: Convert to string for sendMailWithGraph
      const createdUserEmailBody = `
<div style="font-family: 'Segoe UI,Arial,sans-serif'; max-width: 600px; background: #f9f9f9; border-radius: 10px; padding: 24px">
  <div style="font-size: 18px; font-weight: 600; color: #0078d4; margin-bottom: 16px">
    Invoice Request Created
  </div>
  <div style="font-size: 16px; color: #444; margin-bottom: 18px">
    Your new invoice request has been created and is now being tracked in the system.
  </div>
  
// In BOTH email templates - replace the PO Items table section
<table style="width:100%;margin:0 auto">
  <tbody>
    ${allPOItemsData
          .map((item: any) => `
        <tr>
          <td style="padding:4px 8px;font-size:13px">
            ${item.poItemTitle}
            <span style="font-size:11px;color:#666">
              (${currencySymbol}${(item.poItemValue || 0).toLocaleString()})
            </span>
          </td>
          <td style="padding:4px 8px;text-align:right;font-size:13px;font-weight:600">
            ${currencySymbol}${(item.invoicedAmount || 0).toLocaleString()}
          </td>
        </tr>`)
          .join("")}
  </tbody>
</table>

  ${invoiceFormState.Comments ? `
    <div style="margin-bottom: 24px">
      <strong>Comments:</strong>
      <div style="background: #f8f9fa; padding: 12px; border-radius: 6px; font-size: 14px; margin-top: 8px">
        ${invoiceFormState.Comments}
      </div>
    </div>
  ` : ''}

  <div style="margin-bottom: 24px">
    <a href="${itemLink}" style="font-size: 15px; color: #0078d4; text-decoration: underline">
      Click here to view the invoice request
    </a>
  </div>
  
  <div style="border-top: 1px solid #eee; margin-top: 22px; padding-top: 10px; font-size: 13px; color: #999">
    Invoice Tracker - SACHA Group
  </div>
</div>
`.trim();

      const sendNotificationEmail = async () => {
        const subject = `[${siteTitle}]New Invoice Request for ${invoiceFormState.PurchaseOrder}`;
        try {
          const graphClient = await context.msGraphClientFactory.getClient();
          await sendMailWithGraph(graphClient, creatorEmail, subject, createdUserEmailBody);
          setDialogType("success");
          setDialogMessage("Invoice request submitted successfully!");
          // setDialogVisible(true);
          setIsSubmitting(false);

          // Refresh invoiceRequests data after update
          const allPOIDs = mainPOs.map(po => po.POID);
          const updatedInvoices = await fetchInvoiceRequests(sp, allPOIDs);
          setInvoiceRequests(updatedInvoices);
          setActivePOIDFilter(selectedItem?.POID);
          // setInvoiceRequestsForPercent(updatedInvoices);
          setIsInvoicePanelOpen(false);
          setInvoicePanelPO(null);
          // setActivePOIDFilter(selectedItem?.POID);
          setFilterMode('mainPO');
          setUploadedFiles([]);
          setInvoiceFormState(prev => ({ ...prev, Attachment: null }));

        } catch (error) {
          if (addedItemId !== null) {
            await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).delete();
          }
          setDialogType("error");
          setDialogMessage("Error submitting invoice request: " + (error as any)?.message);
          setDialogVisible(true);
          setIsSubmitting(false);
        }
      };

      await sendNotificationEmail();
      // TypeScript: Compose HTML for "finance" email
      if (financeEmails) {
        const financeEmailArray = financeEmails.split(",").map((e: any) => e.trim());
        const financelink = `${appPageUrl}#updaterequests?selectedInvoice=${addedItemId}`;
        const financeEmailBody = `
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;background:#f9f9f9;border-radius:10px;padding:24px">
  <div style="font-size:18px;font-weight:600;color:#0078d4;margin-bottom:16px">
    Invoice Request Submitted
  </div>
  <div style="font-size:16px;color:#444;margin-bottom:18px">
    An invoice request has been submitted and is waiting for your review.
  </div>

  // In BOTH email templates - replace the PO Items table section
<table style="width:100%;margin:0 auto">
  <tbody>
    ${allPOItemsData
            .map((item: any) => `
        <tr>
          <td style="padding:4px 8px;font-size:13px">
            ${item.poItemTitle}
            <span style="font-size:11px;color:#666">
              (${currencySymbol}${(item.poItemValue || 0).toLocaleString()})
            </span>
          </td>
          <td style="padding:4px 8px;text-align:right;font-size:13px;font-weight:600">
            ${currencySymbol}${(item.invoicedAmount || 0).toLocaleString()}
          </td>
        </tr>`)
            .join("")}
  </tbody>
</table>

        </td>
        <td style="padding:12px 8px;text-align:right;font-weight:700;font-size:16px;color:#28a745">
          ${currencySymbol}${Number(invoiceFormState.InvoiceAmount || 0).toLocaleString()}
        </td>
      </tr>
    </tbody>
  </table>

  ${invoiceFormState.Comments
            ? `
  <div style="margin-bottom:24px">
    <strong>Comments</strong>
    <div style="background:#f8f9fa;padding:12px;border-radius:6px;font-size:14px;margin-top:8px">
      ${invoiceFormState.Comments}
    </div>
  </div>`
            : ""
          }

  <div style="margin-bottom:24px">
    <a href="${financelink}" style="font-size:15px;color:#0078d4;text-decoration:underline">
      Click here to update the invoice request
    </a>
  </div>

  <div style="border-top:1px solid #eee;margin-top:22px;padding-top:10px;font-size:13px;color:#999">
    Invoice Tracker - SACHA Group
  </div>
</div>
`.trim();
        const subject = `[${siteTitle}]Invoice Request Submitted`;
        const graphClient = await context.msGraphClientFactory.getClient();
        await sendMailWithGraph(graphClient, financeEmailArray, subject, financeEmailBody);
      }

      const allPOIDs = mainPOs.map(po => po.POID);
      const updatedInvoices = await fetchInvoiceRequests(sp, allPOIDs);  // ✅ This await works
      setInvoiceRequests(updatedInvoices);
      // setInvoiceRequestsForPercent(updatedInvoices);

      setSubmitDialogState('success');
      setSubmitDialogMessage('Request submitted successfully!');

      // Close panels
      setIsInvoicePanelOpen(false);
      setInvoicePanelPO(null);
      setUploadedFiles([]);
      setInvoiceFormState(prev => ({ ...prev, Attachment: null }));

      // Auto-dismiss after 2 seconds (NO AWAITS here)
      setTimeout(() => {
        setSubmitDialogState('idle');
        setIsSubmitting(false);
      }, 4000);
    } catch (error) {
      setSubmitDialogState('idle');
      setIsSubmitting(false);
      if (addedItemId !== null) {
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).delete();
      }
      setDialogType('error');
      setDialogMessage(`Error submitting invoice request: ${error} `);
      setDialogVisible(true);
    }
  };
  // Inside the CreateView component function, before return:

  const adjustedChildColumns = computeColumnWidths(childPOItems, childPOColumns);
  // const filteredChildPOItemsForCreate = React.useMemo(
  //   () =>
  //     childPOItems.filter(item =>
  //       isPoItemUsedInAnyInvoice(
  //         selectedItem?.POID ?? "",
  //         item.POID,
  //         invoiceRequests
  //       )
  //     ),
  //   [childPOItems, invoiceRequests, selectedItem]
  // );
  // console.log('Filtered Child PO Items for Create:', filteredChildPOItemsForCreate);

  // For invoice requests, you have invoiceRequests as state; apply filter if needed:
  // const filteredInvoiceRequests = activePOIDFilter
  //   ? invoiceRequests.filter((ir) => ir.POItemTitle === activePOIDFilter)
  //   : invoiceRequests;

  const adjustedInvoiceColumns = computeColumnWidths(filteredInvoiceRequests, invoiceColumnsView);
  let invoicedItems: { poItemTitle: string; poItemValue: number; invoicedAmount: number }[] = []

  if (selectedInvoiceRequest?.InvoicedAmountsJSON) {
    try {
      const parsed = JSON.parse(selectedInvoiceRequest.InvoicedAmountsJSON)
      if (Array.isArray(parsed)) {
        invoicedItems = parsed.map((r: any) => ({
          poItemTitle: r.poItemTitle ?? r.POItemTitle ?? '',
          poItemValue: Number(r.poItemValue ?? r.Value ?? 0),
          invoicedAmount: Number(r.invoicedAmount ?? r.InvoiceAmount ?? 0)
        }))
      }
    } catch (e) {
      console.warn('Error parsing InvoicedAmountsJSON', e)
    }
  }

  const currencyCodeForDetails =
    invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
  const currencySymbolForDetails = getCurrencySymbol(currencyCodeForDetails)

  const poTotalsFromJson = React.useMemo(
    () =>
      selectedItem?.POID
        ? getTotalsFromJsonForPO(selectedItem.POID, invoiceRequests)
        : { requested: 0, invoiced: 0, paid: 0, cancelled: 0 },
    [selectedItem?.POID, invoiceRequests]
  )

  return (
    <section style={{ background: "#fff", borderRadius: 8, padding: 16 }}>
      <div>
        <h2 style={{ marginBottom: 20 }}>Create Invoice Request</h2>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: "nowrap", overflowX: "auto", paddingBottom: 8 } }}>
          <div>
            <Label>Search</Label>
            <SearchBox
              placeholder="Search"
              value={filters.search}
              onChange={(ev, newVal) => setFilters((f) => ({ ...f, search: newVal || "" }))}
              styles={{ root: { width: 250, minWidth: 250 } }}
            />
          </div>
          <div>
            <Dropdown
              placeholder="Filter by Invoice Status"
              options={invoiceStatusOptions}
              selectedKey={invoiceStatusFilter}
              onChange={(e, option) => {
                setInvoiceStatusFilter(option?.key ? option.key.toString() : null);
                selection.setAllSelected(false);      // Remove selection from main list
                setSelectedItem(null);                // Also clear selected main PO item in state
              }}
              styles={{ root: { width: 250, minWidth: 250 } }}
            />
          </div>
          <div>
            <PrimaryButton
              text="Clear Filters"
              onClick={() => {
                setFilters({ search: "" });
                setInvoiceStatusFilter(null);
                selection.setAllSelected(false);      // Remove selection from main list
                setSelectedItem(null);                // Also clear selected main PO item in state
              }}
              disabled={!isFilterApplied}
              styles={{ root: { backgroundColor: primaryColor } }}
            // styles={{ root: { marginRight: 12 } }}
            />
          </div>
          <Stack.Item align="end" styles={{ root: { paddingLeft: 12 } }}>
            <IconButton
              iconProps={{ iconName: 'Columns' }}
              title="Manage Columns"
              ariaLabel="Manage Columns"
              onClick={() => setIsColumnPanelOpen(true)}
              styles={{ root: { color: primaryColor } }}
            />
          </Stack.Item>
          <div>
            <PrimaryButton text="Create Invoice Request" disabled={!selectedItem} onClick={handleOpenPanel} />
          </div>
        </Stack>
        {loading && <Spinner label="Loading data..." />}
        {error && <div style={{ color: "red" }}>{error}</div>}
        {!loading && !error && (
          <div className={`ms - Grid - row ${styles.detailsListContainer} `}>
            <div style={{ height: 300, position: 'relative' }}>
              <ScrollablePane>
                <div
                  className={`ms - Grid - col ms - sm12 ms - md12 ms - lg12 ${styles.detailsList_Scrollablepane_Container} `}
                >
                  <DetailsList
                    items={filteredMainPOs}
                    columns={getVisibleColumns()}
                    selection={selection}
                    selectionMode={SelectionMode.single}
                    setKey="mainPOsList"
                    isHeaderVisible={true}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    selectionPreservedOnEmptyClick={true}
                    onRenderDetailsHeader={onRenderHeaderGeneric}
                    onItemInvoked={handlePOIDDoubleClick}
                  />
                </div>
                {/* {columnFilterMenu.visible && (
                  <ContextualMenu
                    items={menuItems}
                    target={columnFilterMenu.target}
                    onDismiss={() => setColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                  />
                )} */}
                {mainColumnFilterMenu.visible && (
                  <ContextualMenu
                    items={mainMenuItems}
                    target={mainColumnFilterMenu.target}
                    shouldFocusOnContainer={true}
                    onDismiss={() => setMainColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                  />
                )}
              </ScrollablePane>
            </div>
          </div>
        )}
        <Dialog
          hidden={!dialogVisible}
          onDismiss={() => setDialogVisible(false)}
          dialogContentProps={{
            type: dialogType === "error" ? DialogType.largeHeader : DialogType.normal,
            title: dialogType === "error" ? "Error" : "Success",
            subText: dialogMessage,
          }}
          modalProps={{
            isBlocking: true,
          }}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={() => {
                setDialogVisible(false);
                if (dialogType !== 'error') {
                  setIsInvoicePanelOpen(false);
                  setInvoicePanelPO(null);
                }
              }}
              text="OK"
              styles={{ root: { backgroundColor: primaryColor } }}
            />
          </DialogFooter>
        </Dialog>
        <Panel
          isOpen={isPanelOpen}
          onDismiss={handlePanelDismiss}
          // headerText="Purchase Order"
          closeButtonAriaLabel="Close"
          type={PanelType.extraLarge}
          // customWidth="1000px"
          isLightDismiss={false}
          isBlocking={false}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 18 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
            <div style={{ display: "flex", flexDirection: "row", gap: 24, alignItems: "flex-start", marginTop: 0, marginBottom: 0 }}>
              <TextField
                label="Purchase Order"
                value={selectedItem?.POID}
                readOnly
                disabled
                styles={{ root: { maxWidth: 220, marginTop: 0, marginBottom: 0, fontSize: 15, fontWeight: 600 } }}
              />
              <TextField
                label={`PO Amount`}
                value={`${getCurrencySymbol(invoiceCurrency && invoiceCurrency.trim() !== "" ? invoiceCurrency : "USD")}${selectedItem?.POAmount} `}
                readOnly
                disabled
                styles={{ root: { maxWidth: 220, marginTop: 0, marginBottom: 0, fontSize: 15, fontWeight: 600 } }}
              />
              <TextField
                label="Invoiced Amount"
                value={`${getCurrencySymbol(
                  invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
                )}${poTotalsFromJson.invoiced.toFixed(2)} `}
                readOnly
                disabled
              />

              <TextField
                label="Paid Amount"
                value={`${getCurrencySymbol(
                  invoiceCurrency && invoiceCurrency.trim() !== '' ? invoiceCurrency : 'USD'
                )}${poTotalsFromJson.paid.toFixed(2)} `}
                readOnly
                disabled
              />
              {/* ✅ NEW: Add Button - Next to Paid Amount */}
              <TooltipHost content="Create New Invoice Request" id="create-main-invoice-tooltip" calloutProps={{ gapSpace: 0 }}>
                <span style={{ display: 'inline-block', marginTop: 25 }}> {/* Align with TextField labels */}
                  <PrimaryButton
                    iconProps={{ iconName: 'Add' }}
                    text="New Request"
                    onClick={() => {
                      // Create for entire PO (not specific child item)
                      const mainPOItem: ChildPOItem = {
                        Id: selectedItem?.Id || 0,
                        POID: selectedItem?.POID || '',
                        POAmount: selectedItem?.POAmount || '0',
                        ParentPOIndex: 0,
                        POIndex: 0
                      }
                      childPOSelection.setAllSelected(false)
                      childPOSelection.setKeySelected(mainPOItem.Id.toString(), true, false)
                      handleOpenInvoicePanel(mainPOItem)
                    }}
                    styles={{
                      root: {
                        height: 32,
                        minWidth: 120,
                        backgroundColor: primaryColor
                      }
                    }}
                    disabled={isReadOnlyInvoicePanel}
                  />
                </span>
              </TooltipHost>
            </div>

            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 7, color: "#626262" }}>PO Items:</div>
            <div>
              {fetchingChildPOs ? (
                <Spinner label="Loading child POs..." />
              ) : childPOItems.length > 0 ? (
                <div
                  style={{
                    maxHeight: 200,
                    border: '1px solid #edebe9',
                    borderRadius: 6,
                    overflow: 'auto',
                    backgroundColor: '#fafafa',
                  }}
                >
                  <DetailsList
                    items={childPOItems}
                    columns={adjustedChildColumns}
                    selection={childPOSelection}
                    selectionMode={SelectionMode.single}
                    setKey="childPOs"
                    onActiveItemChanged={handleChildPORowClick}
                    // getKey={item => item.key}
                    onRenderDetailsHeader={onRenderHeaderGeneric}
                    styles={{
                      root: {
                        background: "#fff",
                        border: "1px solid #eee",
                        borderRadius: 6,
                        overflowX: "auto",
                        width: '100%',
                        minWidth: 0,
                        // overflow: 'auto',
                      },
                    }}
                  />
                  {childColumnFilterMenu.visible && (
                    <ContextualMenu
                      items={simpleMenuItems}
                      target={childColumnFilterMenu.target}
                      shouldFocusOnContainer={true}
                      onDismiss={() => setChildColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                    />
                  )}
                </div>
              ) : (
                <div style={{ fontStyle: "italic", marginBottom: 10 }}>No PO items found.</div>
              )}
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginTop: 8,
                  marginBottom: 8,
                  justifyContent: "space-between",
                  fontSize: 15,
                  fontWeight: 600,
                  color: "#626262",
                }}>
                <div>
                  Invoice Requests of {activePOIDFilter ?? selectedItem?.POID ?? ""}
                </div>
                <div>
                  <PrimaryButton
                    text="Show all Invoice Requests"
                    onClick={() => { setActivePOIDFilter(selectedItem?.POID || null); setFilterMode('mainPO'); }}
                    styles={{ root: { marginLeft: 24, backgroundColor: primaryColor } }}
                    disabled={!activePOIDFilter}
                  />
                </div>
              </div>
            </div>
            <div>
              {fetchingInvoices ? (
                <Spinner label="Loading invoice requests..." />
              ) : showInvoices.length > 0 ? (
                <div>
                  <DetailsList
                    items={showInvoices}
                    columns={adjustedInvoiceColumns}
                    selectionMode={SelectionMode.single}
                    onActiveItemChanged={onInvoiceRequestClicked}
                    setKey="invoiceRequests"
                    isHeaderVisible={true}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    selectionPreservedOnEmptyClick={true}
                    onRenderDetailsHeader={onRenderHeaderGeneric}
                    styles={{ root: { overflowX: 'auto', background: "#fff", border: "1px solid #eee", borderRadius: 6 } }}
                  />
                  {columnFilterMenu.visible && (
                    <ContextualMenu
                      items={simpleMenuItems}
                      target={columnFilterMenu.target}
                      shouldFocusOnContainer={true}
                      // onDismiss={() => setColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                      onDismiss={() => {
                        console.log('Menu dismissed');  // DEBUG
                        setColumnFilterMenu({ visible: false, target: null, columnKey: null });
                      }}
                    />
                  )}

                </div>
              ) : (
                <div style={{ fontStyle: "italic" }}>No invoice requests found.</div>
              )}
            </div>
          </Stack>
        </Panel>
        <Panel
          isOpen={isInvoicePanelOpen}
          onDismiss={handleInvoicePanelDismiss}
          headerText="Create Invoice Request"
          closeButtonAriaLabel="Close"
          type={PanelType.custom}
          isLightDismiss={false}
          isFooterAtBottom={false}
          styles={{
            main: {
              right: 0,
              // left: "unset",
              margin: "auto",
              maxWidth: 900,
              minHeight: 450,
              borderRadius: "12px 0 0 12px",
              boxShadow: " -4px 0 16px rgba(0,0,0,0.1)",
              background: "#fafafa",
            },
            scrollableContent: {
              overflowY: "auto",
              paddingLeft: 24,
              paddingRight: 24,
            },
          }}
        >
          {invoicePanelLoading ? (
            <Stack styles={{ root: { minHeight: 180, alignItems: "center", justifyContent: "center" } }}>
              <Spinner label="Loading invoice form..." size={3} />
            </Stack>
          ) : (
            <Stack horizontal tokens={{ childrenGap: 18 }} styles={{ root: { marginTop: 6, marginBottom: 6, width: "100%" } }}>
              {/* LEFT HALF: Invoice Form */}
              <Stack styles={{ root: { flex: 1, minWidth: "400px", maxWidth: "60%" } }} tokens={{ childrenGap: 12 }}>
                <TextField label="PO ID" value={invoiceFormState.POID} readOnly disabled />
                <TextField label="Project Name" value={invoiceFormState.ProjectName} readOnly disabled />
                <TextField label="Cost Center" value={invoiceFormState.CostCenter} readOnly disabled />
                {invoicePanelPO && (
                  <>
                  </>
                )}
                <TextField
                  label="Customer Contact"
                  value={invoiceFormState.CustomerContact}
                  onChange={(_, val) => handleInvoiceFormChange("CustomerContact", val || "")}
                />
                <TextField
                  label="Comments"
                  value={invoiceFormState.Comments}
                  onChange={(_, val) => handleInvoiceFormChange("Comments", val || "")}
                  multiline
                />
                <div style={{ marginTop: 24 }}>
                  <table>
                    <thead>
                      <tr>
                        <th>PO Items Title</th>
                        <th>PO Item Value ({getCurrencySymbol(invoiceCurrency)})</th>
                        <th>Remaining Amount ({getCurrencySymbol(invoiceCurrency)})</th>
                        <th>Invoice Amount ({getCurrencySymbol(invoiceCurrency)})</th>
                      </tr>
                    </thead>
                    <tbody>
                      {lineAllocations.map((row, idx) => (
                        <tr key={row.poItemId}>
                          <td>{row.poItemId}</td>
                          <td>{row.poItemValue.toLocaleString()}</td>
                          <td>{row.remaining.toLocaleString()}</td>
                          <td>
                            <TextField
                              // value={row.invoiceAmount ? row.invoiceAmount.toString() : ""}
                              value={
                                row.invoiceAmount === undefined || row.invoiceAmount === null
                                  ? "0"
                                  : row.invoiceAmount.toString()
                              }
                              onChange={(_, v) => handleLineAmountChange(idx, v || "")}
                              type="number"
                              styles={{ root: { maxWidth: 140 } }}
                              errorMessage={row.error}
                            // required
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div style={{ marginTop: 12, textAlign: "right", fontWeight: 600 }}>
                    Total Invoice Amount:
                    {getCurrencySymbol(invoiceCurrency) + lineAllocations.reduce((sum, r) => sum + (r.invoiceAmount || 0), 0).toLocaleString()}
                  </div>
                </div>

                <PrimaryButton text={isSubmitting ? "Submitting..." : "Submit"} disabled={isSubmitting} onClick={handleInvoiceFormSubmit} styles={{ root: { marginTop: 12, minWidth: 110, backgroundColor: primaryColor } }} />
                {
                  submitDialogState !== 'idle' && (
                    <Dialog
                      hidden={false}
                      onDismiss={() => { }} // Non-dismissible
                      dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: submitDialogState === 'submitting' ? 'Submitting Request' : 'Request Submitted!',
                        // subText: submitDialogMessage || 'Please wait...',
                      }}
                      modalProps={{
                        isBlocking: true,
                        styles: {
                          main: { maxWidth: '400px', borderRadius: '12px' }
                        }
                      }}
                    >
                      <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: '20px' } }}>
                        {submitDialogState === 'submitting' ? (
                          <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
                            <Spinner size={2} />
                            {/* <Text variant="medium">This cannot be cancelled</Text> */}
                          </Stack>
                        ) : (
                          <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }}>
                            <Icon iconName="CheckMark" styles={{ root: { fontSize: 48, color: '#107C10' } }} />
                            <Text variant="xLarge" styles={{ root: { color: '#107C10', fontWeight: 600 } }}>
                              Success!
                            </Text>
                            <Text>Closing in 4 seconds...</Text>
                          </Stack>
                        )}
                      </Stack>
                    </Dialog>
                  )
                }
                {/* {isSubmitting && <Spinner label="Submitting invoice..." />} */}
              </Stack>

              {/* RIGHT HALF: Attachments & Preview */}
              <Stack styles={{ root: { flex: 1, minWidth: "300px", maxWidth: "40%" } }}>
                {previewFileIdx !== null && uploadedFiles[previewFileIdx] ? (
                  <Stack>
                    <PrimaryButton
                      text="Close Preview"
                      onClick={() => setPreviewFileIdx(null)}
                      styles={{ root: { marginBottom: 10, backgroundColor: primaryColor } }}
                    />
                    <iframe
                      src={URL.createObjectURL(uploadedFiles[previewFileIdx])}
                      style={{
                        width: "100%",
                        height: "380px",
                        border: "1px solid #eee",
                        borderRadius: 8,
                        marginBottom: 12,
                      }}
                      title={`Preview - ${uploadedFiles[previewFileIdx].name} `}
                    />
                    <div style={{ fontSize: 14, color: "#888", wordBreak: "break-all", marginBottom: 6 }}>
                      {uploadedFiles[previewFileIdx].name}
                    </div>
                  </Stack>
                ) : (
                  <Stack>
                    <div
                      style={{
                        margin: "10px 0 18px 0",
                        border: "2px dashed #d0d0d0",
                        borderRadius: 8,
                        padding: 24,
                        textAlign: "center",
                        background: isDragActive ? "#f6faff" : "#fafafa",
                        cursor: "pointer"
                      }}
                      onDragOver={e => { e.preventDefault(); setIsDragActive(true); }}
                      onDragLeave={e => { e.preventDefault(); setIsDragActive(false); }}
                      onDrop={handleDropMulti}
                      onClick={() => document.getElementById("multi-attachment-input")?.click()}
                    >
                      <input
                        id="multi-attachment-input"
                        type="file"
                        multiple
                        accept="*/*"
                        style={{ display: "none" }}
                        onChange={handleFilesChange}
                      />
                      <span style={{ fontSize: 36, color: "#bebebe" }}>
                        <i className="ms-Icon ms-Icon--Attach" aria-hidden="true"></i>
                      </span>
                      <div style={{ fontWeight: 500, fontSize: 18, marginBottom: 4 }}>Attachments</div>
                      <div style={{ color: "#888", fontSize: 15, marginTop: 4 }}>
                        Drop or select file(s).
                      </div>
                    </div>
                    <div style={{ marginBottom: 10 }}>
                      {uploadedFiles.length === 0 ? (
                        <div style={{ color: "#999", fontStyle: "italic" }}>No files added yet.</div>
                      ) : (
                        uploadedFiles.map((file, idx) => (
                          <Stack horizontal key={file.name + idx} verticalAlign="center" styles={{ root: { marginBottom: 6 } }}>
                            <span style={{ flex: 1, fontWeight: 520, fontSize: 15, overflow: 'hidden', textOverflow: 'ellipsis' }}>{file.name}</span>
                            <IconButton iconProps={{ iconName: 'Cancel' }} title="Remove" ariaLabel={`Remove ${file.name} `} onClick={() => removeAttachment(idx)} styles={{ root: { height: 28, minWidth: 28, color: '#ba0808' } }} />
                            <PrimaryButton text="Preview" onClick={() => openPreview(URL.createObjectURL(file), file.name)} styles={{ root: { marginLeft: 10, minWidth: 60, height: 28, backgroundColor: primaryColor } }} />
                          </Stack>
                        ))
                      )}
                    </div>
                    {isViewerOpen && (
                      <Panel
                        isOpen={isViewerOpen}
                        onDismiss={closePreview}
                        headerText={previewFileName || "Attachment Preview"}
                        closeButtonAriaLabel="Close"
                        type={PanelType.medium}
                      >
                        <DocumentViewer
                          url={previewUrl ?? ''}
                          fileName={previewFileName}
                          isOpen={isViewerOpen}
                          onDismiss={closePreview}
                        />
                      </Panel>
                    )}
                  </Stack>
                )}
              </Stack>
            </Stack>
          )}

          {invoicePanelPO && (
            <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
              <div style={{ marginTop: 30 }}>
                <Text variant="mediumPlus" block style={{ fontWeight: 600, marginBottom: 8 }}>
                  Existing Invoice Requests for "{invoiceFormState.POItemTitle}"
                </Text>
                <div>
                  <DetailsList
                    items={invoiceRequests.filter((inv) => inv.PurchaseOrderPO === invoiceFormState.PurchaseOrder)}
                    columns={invoiceColumnsView}
                    selectionMode={SelectionMode.single}
                    onItemInvoked={onInvoiceRequestClicked}
                    styles={{ root: { maxHeight: 200, overflowY: "auto", background: "#fafafa", border: "1px solid #eee", borderRadius: 4 } }}
                  />
                </div>
              </div>
            </Stack>
          )}
        </Panel>
        <Panel
          isOpen={isInvoiceRequestViewPanelOpen}
          onDismiss={() => {
            setIsInvoiceRequestViewPanelOpen(false);
            setSelectedInvoiceRequest(null);
          }}
          type={PanelType.medium}
          styles={{
            content: { padding: 20 },
            headerText: { fontWeight: 600, fontSize: 22, color: primaryColor }
          }}
        >
          {selectedInvoiceRequest && (
            <Stack tokens={{ childrenGap: 10 }}>

              {/* ===== HEADER ===== */}
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
                  Invoice Request Details
                </Text>
                <Text styles={{ root: { fontSize: 13, color: "#605e5c" } }}>
                  PO: {selectedInvoiceRequest.PurchaseOrderPO ?? "-"} · Project:{" "}
                  {renderValue(selectedInvoiceRequest.ProjectName)}
                </Text>
              </Stack>

              {/* ===== SUMMARY CARDS ===== */}
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                {[
                  {
                    label: amountLabel,
                    value: `${getCurrencySymbol(invoiceCurrency)}${renderValue(
                      selectedInvoiceRequest.Amount
                    )}`
                  },
                  {
                    label: "Current Status",
                    value: selectedInvoiceRequest.CurrentStatus
                  },
                  {
                    label: "Invoice Status",
                    value: selectedInvoiceRequest.Status
                  }
                ].map((item, idx) => (
                  <Stack
                    key={idx}
                    styles={{
                      root: {
                        background: "#fafafa",
                        padding: 16,
                        borderRadius: 8,
                        minWidth: 180,
                        boxShadow: "0 1px 4px rgba(0,0,0,0.08)"
                      }
                    }}
                  >
                    <Text variant="small" styles={{ root: { color: "#666" } }}>
                      {item.label}
                    </Text>
                    <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
                      {item.value || "—"}
                    </Text>
                  </Stack>
                ))}
              </Stack>

              {/* ===== PO ITEMS TABLE ===== */}
              {invoicedItems.length > 0 && (
                <div
                  style={{
                    background: "#fff",
                    borderRadius: 8,
                    boxShadow: "0 2px 6px rgba(0,0,0,0.08)",
                    padding: 16
                  }}
                >
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: 13
                    }}
                  >
                    <thead>
                      <tr style={{ background: "#f5f5f5" }}>
                        <th style={{ padding: 8, textAlign: "left" }}>PO Item</th>
                        <th style={{ padding: 8, textAlign: "right" }}>PO Amount</th>
                        <th style={{ padding: 8, textAlign: "right" }}>Invoiced</th>
                      </tr>
                    </thead>
                    <tbody>
                      {invoicedItems.map((row, idx) => (
                        <tr
                          key={idx}
                          style={{
                            borderBottom: "1px solid #eee",
                            background: idx % 2 === 1 ? "#fafafa" : "transparent"
                          }}
                        >
                          <td style={{ padding: 8 }}>{row.poItemTitle || "—"}</td>
                          <td style={{ padding: 8, textAlign: "right" }}>
                            {currencySymbolForDetails}
                            {row.poItemValue.toLocaleString()}
                          </td>
                          <td style={{ padding: 8, textAlign: "right" }}>
                            {currencySymbolForDetails}
                            {row.invoicedAmount.toLocaleString()}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {/* ===== COMMENTS (TIMELINE STYLE) ===== */}
              {formatCommentHistory(selectedInvoiceRequest.PMCommentsHistory)?.trim() && (
                <div
                  style={{
                    background: "#f9f9f9",
                    borderLeft: "4px solid primaryColor",
                    padding: 12,
                    borderRadius: 6
                  }}
                >
                  <Text styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
                    Requestor Comments
                  </Text>
                  <pre
                    style={{
                      margin: 0,
                      whiteSpace: "pre-wrap",
                      fontSize: 13,
                      fontFamily: "inherit"
                    }}
                  >
                    {formatCommentHistory(selectedInvoiceRequest.PMCommentsHistory)}
                  </pre>
                </div>
              )}

              {formatCommentHistory(selectedInvoiceRequest.FinanceCommentsHistory)?.trim() && (
                <div
                  style={{
                    background: "#f9f9f9",
                    borderLeft: "4px solid primaryColor",
                    padding: 12,
                    borderRadius: 6
                  }}
                >
                  <Text styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
                    Finance Comments
                  </Text>
                  <pre
                    style={{
                      margin: 0,
                      whiteSpace: "pre-wrap",
                      fontSize: 13,
                      fontFamily: "inherit"
                    }}
                  >
                    {formatCommentHistory(selectedInvoiceRequest.FinanceCommentsHistory)}
                  </pre>
                </div>
              )}

              {/* ===== METADATA ===== */}
              <Separator />

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(2, 1fr)",
                  gap: 12,
                  fontSize: 12,
                  color: "#666"
                }}
              >
                <div>
                  <b>Created:</b>{" "}
                  {new Date(selectedInvoiceRequest.Created).toLocaleDateString("en-GB")}
                </div>
                <div>
                  <b>Modified:</b>{" "}
                  {new Date(selectedInvoiceRequest.Modified).toLocaleDateString("en-GB")}
                </div>
                <div>
                  <b>Created By:</b> {renderValue(selectedInvoiceRequest.CreatedBy)}
                </div>
                <div>
                  <b>Modified By:</b> {renderValue(selectedInvoiceRequest.ModifiedBy)}
                </div>
              </div>
            </Stack>
          )}
        </Panel>
        {/* Column Management Panel */}
        <Panel
          isOpen={isColumnPanelOpen}
          onDismiss={() => setIsColumnPanelOpen(false)}
          headerText="Customize Columns"
          type={PanelType.medium}
        >
          <Stack tokens={{ childrenGap: 16 }}>
            <div style={{ height: 400, overflow: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 12 }}>
              {columns.map((col: any) => (
                <div key={col.key} style={{
                  display: 'flex', alignItems: 'center', padding: 12, marginBottom: 8,
                  borderRadius: 4, backgroundColor: visibleColumns.includes(col.key as string) ? '#f3f2f1' : '#faf9f8'
                }}>
                  <input
                    type="checkbox"
                    checked={visibleColumns.includes(col.key as string)}
                    onChange={() => toggleColumnVisibility(col.key as string)}
                    style={{ marginRight: 12 }}
                  />
                  <span style={{ flex: 1, fontWeight: 600 }}>{col.name}</span>
                  {visibleColumns.includes(col.key as string) && (
                    <div style={{ display: 'flex', gap: 4 }}>
                      <IconButton
                        iconProps={{ iconName: 'ChevronUp' }}
                        title="Move Up"
                        onClick={() => moveColumn(col.key as string, 'up')}
                        disabled={visibleColumns.indexOf(col.key as string) === 0}
                        styles={{ root: { height: 32, width: 32 } }}
                      />
                      <IconButton
                        iconProps={{ iconName: 'ChevronDown' }}
                        title="Move Down"
                        onClick={() => moveColumn(col.key as string, 'down')}
                        disabled={visibleColumns.indexOf(col.key as string) === visibleColumns.length - 1}
                        styles={{ root: { height: 32, width: 32 } }}
                      />
                    </div>
                  )}
                </div>
              ))}
            </div>
            <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
            </Stack>
          </Stack>
        </Panel>

        {/* Column Filter Panel */}
        <Panel
          isOpen={isFilterPanelOpen}
          onDismiss={() => setIsFilterPanelOpen(false)}
          headerText="Filter column"
          type={PanelType.smallFixedFar}
        >
          {currentFilterColumn && (
            <Stack tokens={{ childrenGap: 8 }}>
              {getColumnDistinctValues(currentFilterColumn).map(val => {
                const selected = columnFilters[currentFilterColumn]?.includes(val) ?? false;
                return (
                  <Checkbox
                    key={val}
                    label={val}
                    checked={selected}
                    onChange={(_, checked) => {
                      setColumnFilters(prev => {
                        const current = prev[currentFilterColumn] || [];
                        const next = checked
                          ? [...current, val]
                          : current.filter(v => v !== val);
                        return { ...prev, [currentFilterColumn]: next };
                      });
                    }}
                  />
                );
              })}
            </Stack>
          )}
        </Panel>
      </div>
    </section >
  );
};

export default CreateView;
