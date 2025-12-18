import * as React from "react";
import { useState, useEffect } from "react";
import {
  DetailsList,
  IColumn,
  Stack,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  Spinner,
  MessageBar,
  MessageBarType,
  Selection,
  SelectionMode,
  Panel,
  PanelType,
  Dialog,
  DialogType,
  DialogFooter,
  Label,
  DatePicker,
  IconButton,
  IDetailsHeaderProps,
  TooltipHost,
  IRenderFunction,
  Sticky,
  StickyPositionType,
  // ScrollablePane,
  ContextualMenu,
  ContextualMenuItemType,
  DetailsListLayoutMode
} from "@fluentui/react";
import * as XLSX from 'xlsx';
import { MSGraphClient } from '@microsoft/sp-http';
import { saveAs } from 'file-saver';
import { SPFI } from "@pnp/sp";
import DocumentViewer from "../DocumentViewer";
import styles from './FinanceView.module.scss';
interface FinanceViewProps {
  sp: SPFI;
  context: any;
  initialFilters?: {
    search?: string;
    requestedDate?: Date | null;
    customer?: string;
    Status?: string;
    FinanceStatus?: string;
    CurrentStatus?: string;
  };
  onNavigate: (pageKey: string, params?: any) => void;
  projectsp: SPFI;
}

// STATUS OPTIONS (STEP LABELS)
const InvstatusOptions: IDropdownOption[] = [
  { key: 'All', text: 'All' },
  { key: "Invoice Requested", text: "Invoice Requested" },
  { key: "Invoice Raised", text: "Invoice Raised" },
  { key: "Pending Payment", text: "Pending Payment" },
  { key: "Payment Received", text: "Payment Received" },
  { key: "Cancelled", text: "Cancelled" }
];
// Current Status Options
const CURRENT_STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'All', text: 'All' },
  { key: 'Request Submitted', text: 'Request Submitted' },
  { key: 'Pending Finance', text: 'Pending Finance' },
  { key: 'Finance asked Clarification', text: 'Finance asked Clarification' },
  { key: 'Completed', text: 'Completed' },
  { key: 'Cancelled Request', text: 'Cancelled Request' }
];
const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";

export default function FinanceView({ sp, projectsp, context, initialFilters, onNavigate }: FinanceViewProps) {
  const [items, setItems] = useState<any[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedItem, setSelectedItem] = useState<any>(null);
  // const [isClarificationOpen,] = React.useState(false);
  // const [isViewerOpen, setIsViewerOpen] = useState(false);
  const [viewerFileUrl, setViewerFileUrl] = useState<string | null>(null);
  const [viewerFileName, setViewerFileName] = useState<string | null>(null);
  const [originalStatus, setOriginalStatus] = useState<string | null>(null);
  const [invoiceNumberLoaded, setInvoiceNumberLoaded] = useState(false);
  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");
  const [dialogType, setDialogType] = useState<"success" | "error">("success");
  const [isDragActive, setIsDragActive] = useState(false);
  const [, setCustomerOptions] = useState<IDropdownOption[]>([]);
  // const [, setStatusOptions] = useState<IDropdownOption[]>([]);
  // Column management states
  const [visibleColumns, setVisibleColumns] = useState<string[]>([]);
  const [columnOrder,] = useState<Record<string, number>>({});
  const [isColumnPanelOpen, setIsColumnPanelOpen] = useState(false);
  // Column filters per column – now multi‑select values
  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState(false);
  const [currentFilterColumn, setCurrentFilterColumn] = useState<string>('');
  const [isClarificationPending, setIsClarificationPending] = React.useState(false);
  // const [isPreviewing, setIsPreviewing] = useState(false);
  // const [currentstatusOptions, setcurrentstatusOptions] = useState<IDropdownOption[]>([]);
  const [financeAttachments, setFinanceAttachments] = useState<{ name: string; url: string }[]>([]);
  const [columnFilterMenu, setColumnFilterMenu] = useState<{ visible: boolean; target: HTMLElement | null; columnKey: string | null }>({
    visible: false,
    target: null,
    columnKey: null,
  });
  const onColumnHeaderClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn) => {
    if (column && ev) {
      setColumnFilterMenu({ visible: true, target: ev.currentTarget, columnKey: column.key });
    }
  };
  // const invoiceStatusPriority = [
  //   "Invoice Requested",
  //   "Invoice Raised",
  //   "Pending Payment",
  //   "Payment Received"
  // ];
  // const [filters, setFilters] = useState({
  //   search: initialFilters?.search || "",
  //   requestedDate: initialFilters?.requestedDate || null,
  //   customer: initialFilters?.customer || "",
  //   status: initialFilters?.Status || "",
  //   financeStatus: initialFilters?.FinanceStatus || "",
  //   currentstatus: initialFilters?.CurrentStatus || "",
  // });

  const [filters, setFilters] = useState({
    search: initialFilters?.search || "",
    requestedDate: initialFilters?.requestedDate || null,
    customer: initialFilters?.customer || "",
    status: initialFilters?.Status ? [initialFilters.Status] : ["All"],
    financeStatus: initialFilters?.FinanceStatus || "",
    currentstatus: initialFilters?.CurrentStatus ? [initialFilters.CurrentStatus] : ["All"],
  });

  // Edit form fields and attachments
  const [editFields, setEditFields] = useState<any>({});
  const [attachments, setAttachments] = useState<File[]>([]);

  // PM Attachments loaded from SharePoint folder
  const [pmAttachments, setPmAttachments] = useState<{ name: string; url: string }[]>([]);

  // Selection for DetailsList
  const [selection] = useState(() =>
    new Selection({
      onSelectionChanged: () => {
        const selArr = selection.getSelection();
        const sel = selArr && selArr.length ? selArr[0] : null;
        console.log("Selection changed, selected item:", sel);
        if (sel) {
          setSelectedItem(sel);
          if (openEditForm) openEditForm(sel);
        } else {
          setSelectedItem(null);
          // if (closeEditForm) closeEditForm();
        }
      }
    })
  );

  const getColumnDistinctValues = (columnKey: string): string[] => {
    const col = columns.find(c => c.key === columnKey);
    if (!col || !col.fieldName) return [];
    const field = col.fieldName;

    const values = Array.from(
      new Set(
        items
          .map(i => i[field])
          .filter(v => v !== null && v !== undefined && v !== '')
          .map(v => v.toString())
      )
    );
    return values.sort((a, b) => a.localeCompare(b));
  };

  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [isDocPanelOpen, setIsDocPanelOpen] = useState(false);

  // const [sortState, setSortState] = useState<{ column: string; isSortedDescending: boolean } | null>(null);

  // const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
  //   const isSortedDescending = direction === 'desc';
  //   // setSortState({ column: columnKey, isSortedDescending });

  //   const sortedItems = [...items].sort((a: any, b: any) => {
  //     let aVal = a[columnKey];
  //     let bVal = b[columnKey];

  //     if (aVal === null && bVal === null) return 0;
  //     if (aVal === null) return 1;
  //     if (bVal === null) return -1;

  //     if (aVal instanceof Date) aVal = aVal.getTime();
  //     if (bVal instanceof Date) bVal = bVal.getTime();

  //     if (typeof aVal === 'number' && typeof bVal === 'number') {
  //       return isSortedDescending ? bVal - aVal : aVal - bVal;
  //     }

  //     const aStr = aVal?.toString() ?? '';
  //     const bStr = bVal?.toString() ?? '';
  //     return isSortedDescending
  //       ? bStr.localeCompare(aStr)
  //       : aStr.localeCompare(bStr);
  //   });

  //   setItems(sortedItems);
  //   setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  // };

  const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
    const isAmountField = ['POItemx0020Value', 'InvoiceAmount'].includes(columnKey)

    const sortedItems = [...items].sort((a: any, b: any) => {
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

    setItems(sortedItems)
    setColumnFilterMenu({ visible: false, target: null, columnKey: null })
  }

  const menuItems = [
    {
      key: 'sortAsc', text: 'Sort A→Z', iconProps: { iconName: 'SortUp' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc')
    },
    {
      key: 'sortDesc', text: 'Sort Z→A', iconProps: { iconName: 'SortDown' },
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
    }
  ];

  const getSelectedInvoiceIdFromUrl = (): number | null => {
    const hash = window.location.hash; // e.g. "#updaterequests?selectedInvoice=72"
    if (!hash.startsWith('#updaterequests')) return null;

    const queryString = hash.split('?')[1]; // gets "selectedInvoice=72"
    if (!queryString) return null;

    const params = new URLSearchParams(queryString);
    const selectedInvoice = params.get('selectedInvoice');
    return selectedInvoice ? parseInt(selectedInvoice, 10) : null;
  };

  const getVisibleColumns = (): IColumn[] => {
    return columns
      .filter(col => visibleColumns.includes(col.key as string))
      .sort((a, b) => {
        const aOrder = columnOrder[a.key as string] ?? visibleColumns.indexOf(a.key as string);
        const bOrder = columnOrder[b.key as string] ?? visibleColumns.indexOf(b.key as string);
        return aOrder - bOrder;
      });
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

  const toggleColumnVisibility = (columnKey: string) => {
    setVisibleColumns(prev =>
      prev.includes(columnKey)
        ? prev.filter(k => k !== columnKey)
        : [...prev, columnKey]
    );
  };

  const clearColumnFilter = (columnKey: string) => {
    setColumnFilters(prev => {
      const newFilters = { ...prev };
      delete newFilters[columnKey];
      return newFilters;
    });
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  };

  async function fetchData() {
    setLoading(true);
    setError(null);
    try {
      const fieldNames = [
        "Id",
        "PurchaseOrder",
        "ProjectName",
        "Status",
        "Comments",
        "POItem_x0020_Title",
        "POItem_x0020_Value",
        "InvoiceAmount",
        "Customer_x0020_Contact",
        "Modified",
        "Created",
        "FinanceStatus",
        "PMCommentsHistory",
        "FinanceCommentsHistory",
        "InvoiceNumber",
        "CurrentStatus",
        "Modified",
        "Created",
        "Author/Title",
        "Editor/Title",
        "DueDate",
        "Currency",
      ];
      const calculateWidth = (header: string) => Math.max(80, Math.min(header.length * 15, 300));
      const cols: IColumn[] = [
        { key: "PurchaseOrder", name: "Purchase Order", fieldName: "PurchaseOrder", minWidth: 100, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 120, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        {
          key: "CurrentStatus",
          name: "Current Status",
          fieldName: "CurrentStatus",
          minWidth: 130,
          isResizable: true,
          onRender: (item) => item.CurrentStatus || "-",
          isCollapsible: true,
        },
        { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 130, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        // { key: "Currency", name: "Currency", fieldName: "Currency", minWidth: 150, isResizable: true, isCollapsible: true, },
        { key: "DueDate", name: "DueDate", fieldName: "DueDate", minWidth: 90, isResizable: true, onRender: item => item.DueDate ? new Date(item.DueDate).toLocaleDateString() : "-", isCollapsible: true, onColumnClick: onColumnHeaderClick },
        // { key: "Comments", name: "Requestor Comments", fieldName: "Comments", minWidth: 160, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        { key: "POItem_x0020_Title", name: "PO Item Title", fieldName: "POItem_x0020_Title", minWidth: 120, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        {
          key: "POItem_x0020_Value", name: "PO Item Value", fieldName: "POItem_x0020_Value", minWidth: 100, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick, onRender: (item: any) => {
            if (item.POItem_x0020_Value != null && !isNaN(Number(item.POItem_x0020_Value))) {
              const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
              const value = item.POItem_x0020_Value ?? 0;
              return <span>{symbol} {Number(value).toLocaleString()}</span>;
              // return `${Number(item.POItem_x0020_Value).toLocaleString()}`.trim();
            }
            return '';
          }
        },
        {
          key: "InvoiceAmount", name: "Invoiced Amount", fieldName: "InvoiceAmount", minWidth: 100, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick, onRender: (item: any) => {
            if (item.InvoiceAmount != null && !isNaN(Number(item.InvoiceAmount))) {
              const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
              const value = item.InvoiceAmount ?? 0;
              return <span>{symbol} {Number(value).toLocaleString()}</span>;
              // return `${Number(item.InvoiceAmount).toLocaleString()} ${item.Currency ?? ''}`.trim();
            }
            return '';
          }
        },
        { key: "Customer_x0020_Contact", name: "Customer Contact", fieldName: "Customer_x0020_Contact", minWidth: 120, isResizable: true, isCollapsible: true, onColumnClick: onColumnHeaderClick },
        {
          key: "Created", name: "Created", fieldName: "Created", minWidth: calculateWidth("Created"), isResizable: true, onRender: item => new Date(item.Created).toLocaleDateString(), isCollapsible: true, onColumnClick: onColumnHeaderClick
        },
        {
          key: "CreatedBy", name: "Created By", fieldName: "Author", minWidth: calculateWidth("Created By"), isResizable: true,
          onRender: item => item.Author?.Title || "-",
          isCollapsible: true, onColumnClick: onColumnHeaderClick
        },
        {
          key: "Modified", name: "Modified", fieldName: "Modified", minWidth: calculateWidth("Modified"), isResizable: true,
          onRender: item => new Date(item.Modified).toLocaleDateString(),
          isCollapsible: true, onColumnClick: onColumnHeaderClick
        },
        {
          key: "ModifiedBy", name: "Modified By", fieldName: "Editor", minWidth: calculateWidth("Modified By"), isResizable: true,
          onRender: item => item.Editor?.Title || "-",
          isCollapsible: true, onColumnClick: onColumnHeaderClick
        },
      ];
      setColumns(cols);
      setVisibleColumns(cols.map(c => c.key as string));

      const list = sp.web.lists.getByTitle("Invoice Requests");
      const allItems: any[] = [];

      for await (const page of list.items
        .select(...fieldNames, "AttachmentFiles")
        .expand("AttachmentFiles", "Author", "Editor")
        .top(2000) // page size; adjust if needed
      ) {
        allItems.push(...page);
      }

      setItems(allItems);

      setCustomerOptions(
        Array.from(
          new Set(allItems.map(i => i.Customer).filter(Boolean))
        ).map(val => ({ key: val, text: val }))
      );

      // const listItems = await sp.web.lists
      //   .getByTitle("Invoice Requests")
      //   .items.select(...fieldNames, "AttachmentFiles")
      //   .expand("AttachmentFiles", "Author", "Editor")
      //   ();

      // setItems(listItems);

      // setCustomerOptions(Array.from(new Set(listItems.map(i => i.Customer).filter(Boolean))).map(val => ({ key: val, text: val })));
    } catch (e: any) {
      setError("Unable to load invoice requests: " + (e.message ?? e));
      setItems([]);
      setColumns([]);
    }
    setLoading(false);
  }

  useEffect(() => {
    fetchData();
  }, [sp]);

  useEffect(() => {
    if (initialFilters) {
      setFilters((f) => ({
        ...f,
        ...initialFilters,
      }));
    }
  }, [initialFilters]);

  // useEffect(() => {
  //   const initialVisible = columns.map(col => col.key as string);
  //   setVisibleColumns(initialVisible);
  //   setColumns(columns);
  // }, []);

  React.useEffect(() => {
    const invoiceId = getSelectedInvoiceIdFromUrl();
    if (invoiceId) {
      sp.web.lists
        .getByTitle('Invoice Requests')
        .items.getById(invoiceId)
        ()
        .then((item) => {
          setSelectedItem(item); // set invoice as selected
          setIsPanelOpen(true);  // open the panel
          loadPmAttachments(item); // load attachments if applicable
          // Set edit fields if you have like in openEditForm (optional)
          setEditFields({
            Status: item.Status?.trim(),
            FinanceComments: item.FinanceComments ?? '',
            InvoiceNumber: item.InvoiceNumber ?? '',
            DueDate: item.DueDate ?? null,
            // Set other fields if needed
          });
        })
        .catch((error: any) => {
          console.error('Failed to load invoice from URL ID:', error);
          // Optionally handle error or clear URL param here
        });
    }
  }, []);

  const showDialog = (message: string, type: "success" | "error") => {
    setDialogMessage(message);
    setDialogType(type);
    setDialogVisible(true);
  };

  const handleDialogClose = async () => {
    setDialogVisible(false);
    if (dialogType === 'success') {
      setIsPanelOpen(false);
      setDialogMessage("");
      setSelectedItem(null);
      setTimeout(() => {
        fetchData();
      }, 400);
    }
  };


  const handlePanelDismiss = () => {
    setSelectedItem(null);
    // setDialogVisible(false);
    setIsPanelOpen(false);
    setAttachments([]);  // clear attachments on close
    setEditFields({});   // optional: reset form fields too
    setPmAttachments([]);
    // setIsViewerOpen(false);
  };

  const handleDocPanelDismiss = () => {
    setIsDocPanelOpen(false);
  };

  const clearFilters = () => {
    setFilters({
      search: "",
      requestedDate: null,
      customer: "",
      status: ["All"],
      financeStatus: "",
      currentstatus: ["All"],
    });
  };

  const clearAllFilters = () => {
    clearFilters();
    setColumnFilters({});
  };

  const handleExportToExcel = () => {
    if (!filteredItems.length) {
      setDialogMessage('No available Data to export');
      setDialogType('error');
      setDialogVisible(true);
      return;
    }

    const exportData = filteredItems.map(item => {
      const obj: Record<string, any> = {};

      // columns.forEach(col => {
      //   const field = col.fieldName!;
      //   let value = item[field];

      //   // Special handling for nested or computed fields
      //   if (field === 'Author') value = item.Author?.Title || '-';
      //   else if (field === 'Editor') value = item.Editor?.Title || '-';
      //   else if (field === 'Created' && value) value = new Date(value).toLocaleDateString();
      //   else if (field === 'Modified' && value) value = new Date(value).toLocaleDateString();

      //   obj[col.name] = value ?? '-';
      // });
      columns.forEach(col => {
        const field = col.fieldName!;
        let value = item[field];

        if (field === 'Author') {
          value = item.Author?.Title || '-';
        } else if (field === 'Editor') {
          value = item.Editor?.Title || '-';
        } else if (field === 'Created' && value) {
          value = new Date(value).toLocaleDateString();
        } else if (field === 'Modified' && value) {
          value = new Date(value).toLocaleDateString();
        } else if (field === 'POItem_x0020_Value') {
          const symbol = getCurrencySymbol(item.Currency);
          const num = Number(item.POItem_x0020_Value || 0);
          value = `${symbol} ${num.toLocaleString()}`;
        } else if (field === 'InvoiceAmount') {
          const symbol = getCurrencySymbol(item.Currency);
          const num = Number(item.InvoiceAmount || 0);
          value = `${symbol} ${num.toLocaleString()}`;
        }

        obj[col.name] = value ?? '-';
      });


      return obj;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'InvoiceRequests');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), `InvoiceRequests_${new Date().toISOString()}.xlsx`);
  };


  // const filteredItems = React.useMemo(() => {
  //   const searchText = filters.search?.toLowerCase() || "";

  //   return items.filter(item => {
  //     const matchesSearch =
  //       !searchText ||
  //       columns.some(col => {
  //         const fieldValue = item[col.fieldName] ?? "";
  //         return fieldValue.toString().toLowerCase().includes(searchText);
  //       });

  //     return (
  //       matchesSearch &&
  //       (!filters.customer || item.Customer === filters.customer) &&
  //       // (!filters.status || item.Status === filters.status) &&
  //       (
  //         !filters.status.length ||
  //         filters.status.includes("All") ||
  //         filters.status.includes(item.Status)
  //       ) &&
  //       (
  //         !filters.currentstatus.length ||
  //         filters.currentstatus.includes("All") ||
  //         filters.currentstatus.includes(item.CurrentStatus)
  //       ) &&
  //       (!filters.financeStatus || item.FinanceStatus === filters.financeStatus) &&
  //       // (!filters.currentstatus || item.CurrentStatus === filters.currentstatus) &&
  //       (!filters.requestedDate || (item.RequestedDate && new Date(item.RequestedDate).toLocaleDateString() === filters.requestedDate.toLocaleDateString()))
  //     );
  //   });
  // }, [items, columns, filters]);

  const filteredItems = React.useMemo(() => {
    const searchText = filters.search?.toLowerCase() || '';

    return items.filter(item => {
      // global search
      const matchesSearch =
        !searchText ||
        columns.some(col => {
          const fieldValue = item[col.fieldName ?? ''] ?? '';
          return fieldValue.toString().toLowerCase().includes(searchText);
        });

      // per‑column checklist filters
      const matchesColumnFilters = Object.entries(columnFilters).every(([colKey, selectedVals]) => {
        if (!selectedVals || selectedVals.length === 0) return true;

        const col = columns.find(c => c.key === colKey);
        if (!col || !col.fieldName) return true;

        const value = item[col.fieldName];
        if (value === null || value === undefined || value === '') return false;

        const vStr = value.toString();
        return selectedVals.includes(vStr);
      });

      return (
        matchesSearch &&
        matchesColumnFilters &&
        (!filters.customer || item.Customer === filters.customer) &&
        (!filters.status.length ||
          filters.status.includes('All') ||
          filters.status.includes(item.Status)) &&
        (!filters.currentstatus.length ||
          filters.currentstatus.includes('All') ||
          filters.currentstatus.includes(item.CurrentStatus)) &&
        (!filters.financeStatus || item.FinanceStatus === filters.financeStatus) &&
        (!filters.requestedDate ||
          (item.RequestedDate &&
            new Date(item.RequestedDate).toLocaleDateString() ===
            filters.requestedDate.toLocaleDateString()))
      );
    });
  }, [items, columns, filters, columnFilters]);

  useEffect(() => {
    setCustomerOptions(getUniqueOptions(items, "Customer"));
    // setStatusOptions(getUniqueOptions(items, "Status"));
    // setcurrentstatusOptions(getUniqueOptions(items, "CurrentStatus"));
  }, [items]);

  // useEffect(() => {
  //   const raw = Array.from(
  //     new Set(items.map(i => i.Status).filter(Boolean))
  //   ) as string[];

  //   const ordered: string[] = [];

  //   // 1. push 4 required statuses in fixed order if present
  //   invoiceStatusPriority.forEach(s => {
  //     if (raw.includes(s)) {
  //       ordered.push(s);
  //     }
  //   });

  //   // 2. push remaining statuses sorted
  //   raw
  //     .filter(s => !invoiceStatusPriority.includes(s))
  //     .sort((a, b) => a.localeCompare(b))
  //     .forEach(s => ordered.push(s));

  //   // 3. set final options with All at top
  //   setStatusOptions([
  //     { key: "All", text: "All" },
  //     ...ordered.map(s => ({ key: s, text: s }))
  //   ]);
  // }, [items]);

  // useEffect(() => {
  //   const style = document.createElement('style');
  //   style.innerHTML = '[class*="contentContainer-"]';
  //   document.head.appendChild(style);
  //   return () => { document.head.removeChild(style); };
  // }, []);

  // useEffect(() => {
  //   const raw = Array.from(
  //     new Set(items.map(i => i.CurrentStatus).filter(Boolean))
  //   ) as string[];

  //   setcurrentstatusOptions([
  //     { key: "All", text: "All" },
  //     ...raw.map(s => ({ key: s, text: s }))
  //   ]);
  // }, [items]);

  const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) {
      return null;
    }
    return (
      <Sticky stickyPosition={StickyPositionType.Header}>
        {defaultRender!({ ...props })}
      </Sticky>
    );
  };

  function getUniqueOptions(data: any[], field: string): IDropdownOption[] {
    const uniqueVals = Array.from(
      new Set(
        data
          .map(item => item[field])
          .filter(v => v !== null && v !== undefined && v !== "")
          .map(v => v.toString())
      )
    );

    return uniqueVals.map(val => ({ key: val, text: val }));
  }

  // Handle fields change
  function handleFieldChange(field: string, value: any) {
    setEditFields((prev: any) => ({ ...prev, [field]: value }));
  }

  function decodeHtmlEntities(str: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = str;
    return txt.value;
  }

  function formatCommentHistory(jsonStr?: string): string {
    if (!jsonStr) return "";

    try {
      // Decode HTML entities before parsing JSON
      const decodedStr = decodeHtmlEntities(jsonStr);

      const arr = JSON.parse(decodedStr);
      if (!Array.isArray(arr)) return "";

      const formattedComments = arr.map((entry: any) => {
        const dateObj = entry.Date ? new Date(entry.Date) : null;
        const date = dateObj ? dateObj.toLocaleDateString() : "";
        const time = dateObj ? dateObj.toLocaleTimeString() : "";
        const title = entry.Title || entry.title || "";
        const user = entry.User || "";
        const role = entry.Role ? ` (${entry.Role})` : "";
        const data = entry.Data || entry.comment || "";
        return `[${date} ${time}]${user}${role} - ${title}: ${data}`;
      }).join("\n\n");

      console.log(formattedComments); // Log the output before returning

      return formattedComments;
    } catch (err) {
      console.error("Failed to format comment history", err, jsonStr);
      return "";
    }
  }

  // function getCurrencySymbol(currencyCode: string, locale = 'en-US'): string {
  //   return new Intl.NumberFormat(locale, {
  //     style: 'currency',
  //     currency: currencyCode,
  //     minimumFractionDigits: 0,
  //     maximumFractionDigits: 0
  //   })
  //     .formatToParts(1)
  //     .find(part => part.type === 'currency')?.value || currencyCode;
  // }

  function getCurrencySymbol(currencyCode?: string, locale: string = 'en-US'): string {
    if (!currencyCode) return '€'; // Default fallback

    // Trim whitespace and validate
    const trimmedCode = (currencyCode || '').trim();
    if (!trimmedCode || trimmedCode.length !== 3) return '€';

    try {
      return new Intl.NumberFormat(locale, {
        style: 'currency',
        currency: trimmedCode,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      }).formatToParts(1).find(part => part.type === 'currency')?.value || '€';
    } catch (error) {
      console.warn(`Invalid currency code: ${trimmedCode}`, error);
      return '€'; // Safe fallback
    }
  }

  async function sendMailUsingGraph(graphClient: MSGraphClient, toEmail: string, subject: string, body: string): Promise<void> {
    const mail = {
      message: {
        subject: subject,
        body: {
          contentType: 'HTML',
          content: body
        },
        toRecipients: [
          {
            emailAddress: {
              address: toEmail
            }
          }
        ]
      }
    };

    await graphClient.api('/me/sendMail').post(mail);
  }


  async function sendFinanceClarificationEmail(item: any) {
    if (!item) return;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const siteTitle = context.pageContext.web.title;
    // const toEmail = item.Author?.Email;
    const myRequestsUrl = `${siteUrl}#myrequests?selectedInvoice=${item.Id}`;
    const financeClarificationEmailBody = `
    <div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;background:#f9f9f9;border-radius:10px;padding:24px;">
      <div style="font-size:18px;font-weight:600;color:#1976d2;margin-bottom:16px;">
        Clarification Required: Invoice Request
      </div>
      <div style="font-size:16px;color:#444;margin-bottom:18px;">
        Please provide clarification by reviewing your invoice request.
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:15px;color:#333;margin-bottom:20px;">
        <tr>
          <td style="font-weight:600;padding:6px 0;">Purchase Order:</td>
          <td>${item.PurchaseOrder}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">Project Name:</td>
          <td>${item.ProjectName ?? "N/A"}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">PO Item Title:</td>
          <td>${item.POItem_x0020_Title ?? "N/A"}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">Finance Comments:</td>
          <td>${item.FinanceComments ?? "—"}</td>
        </tr>
      </table>
      <div style="margin-bottom:24px;">
        <a href="${myRequestsUrl}" style="font-size:15px;color:#0078d4;text-decoration:underline;">
          Click here to review and clarify
        </a>
      </div>
      <div style="border-top:1px solid #eee;margin-top:22px;padding-top:10px;font-size:13px;color:#999;">
        Invoice Tracker | SACHA Group
      </div>
    </div>
    `;

    try {
      const authorId = item?.AuthorId;
      const authorUser = await sp.web.getUserById(authorId)();
      const toEmail = authorUser.Email;
      const subject = `[${siteTitle}]Clarification Required on Invoice Request PO ${item.PurchaseOrder}`;

      context.msGraphClientFactory.getClient()
        .then(async (graphClient: MSGraphClient) => {
          await sendMailUsingGraph(graphClient, toEmail, subject, financeClarificationEmailBody);
        })
        .catch((error: any) => {
          console.error('Failed to send finance clarification email', error);
        });

    } catch (error) {
      console.error("Failed to send finance clarification email", error);
    }
  }

  async function sendPmStatusChangeEmail(item: any, oldStatus: string, newStatus: string) {
    if (!item) return;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const siteTitle = context.pageContext.web.title;
    const myRequestsUrl = `${siteUrl}#myrequests?selectedInvoice=${item.Id}`;
    const pmStatusChangeEmailBody = `
    <div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;background:#f9f9f9;border-radius:10px;padding:24px;">
      <div style="font-size:18px;font-weight:600;color:#1976d2;margin-bottom:16px;">
        Invoice Request Status Changed
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:15px;color:#333;margin-bottom:20px;">
        <tr>
          <td style="font-weight:600;padding:6px 0;">Purchase Order:</td>
          <td>${item.PurchaseOrder}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">Project Name:</td>
          <td>${item.ProjectName ?? "N/A"}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">PO Item Title:</td>
          <td>${item.POItem_x0020_Title ?? "N/A"}</td>
        </tr>
        <tr>
          <td style="font-weight:600;padding:6px 0;">New Status:</td>
          <td>${newStatus}</td>
        </tr>
      </table>
      <div style="margin-bottom:24px;">
        <a href="${myRequestsUrl}" style="font-size:15px;color:#1976d2;text-decoration:underline;">
          View Invoice Request
        </a>
      </div>
      <div style="border-top:1px solid #eee;margin-top:22px;padding-top:10px;font-size:13px;color:#999;">
        Invoice Tracker | SACHA Group
      </div>
    </div>
    `;

    try {
      const authorId = item?.AuthorId;
      const authorUser = await sp.web.getUserById(authorId)();
      const toEmail = authorUser.Email;
      const subject = `[${siteTitle}]Invoice Request Updated: PO ${item.PurchaseOrder}`;

      context.msGraphClientFactory.getClient()
        .then(async (graphClient: MSGraphClient) => {
          await sendMailUsingGraph(graphClient, toEmail, subject, pmStatusChangeEmailBody);
        })
        .catch((error: any) => {
          console.error('Failed to send finance clarification email', error);
        });

    } catch (error) {
      console.error("Failed to send Requestor status change email", error);
    }
  }

  async function loadPmAttachments(item: any) {
    if (!item) {
      setPmAttachments([]);
      return;
    }

    const attachments = item.AttachmentFiles || [];
    const pmAttachments = attachments
      .filter((att: any) => att.FileName.match(/Requestor(\.[^.]*)?$/i))
      .map((att: any) => ({ name: att.FileName, url: att.ServerRelativeUrl }));

    setPmAttachments(pmAttachments);
  }

  async function loadFinanceAttachments(item: any) {
    if (!item) {
      setFinanceAttachments([]);
      return;
    }
    const attachments = item.AttachmentFiles;
    const financeAttachments = attachments
      .filter((att: any) => att.FileName.match(/Finance/i))
      .map((att: any) => ({ name: att.FileName, url: att.ServerRelativeUrl }));
    setFinanceAttachments(financeAttachments);
  }

  // Open edit panel and load PM attachments
  function openEditForm(item: any) {
    if (!item) return;
    setInvoiceNumberLoaded(!!item.InvoiceNumber);

    // Determine the invoice status to use in the form:
    const normalizedStatus = (item.Status || "").trim();
    const defaultStatusForSubmitted = "Invoice Requested";
    const submittedStates = ["Request Submitted"];

    const statusToUse = submittedStates.includes(normalizedStatus)
      ? defaultStatusForSubmitted
      : normalizedStatus;

    setEditFields({
      Status: statusToUse,
      FinanceComments: item.FinanceComments ?? "",
      InvoiceNumber: item.InvoiceNumber || "",
      FinanceStatus: "Request Submitted",
      CurrentStatus: "",
      DueDate: item.DueDate || '',
    });

    console.log(item.DueDate);
    console.log(editFields.DueDate);
    setOriginalStatus(item.Status ?? null);
    setAttachments([]);
    loadPmAttachments(item);
    loadFinanceAttachments(item);

    const cs = (item.CurrentStatus || "").toLowerCase();
    const clarificationPending = cs.includes("finance asked clarification")
    setIsClarificationPending(clarificationPending);
    setIsPanelOpen(true);
  }

  async function handleClarification() {
    if (!selectedItem) return;

    if (!editFields.FinanceComments || editFields.FinanceComments.trim() === "") {
      setDialogMessage("Comments have to be entered in the Finance Comments field.");
      setDialogType("error");
      setDialogVisible(true);
      return;
    }
    // const userRole = await getCurrentUserRole(context, selectedItem);
    try {
      // Parse existing FinanceCommentsHistory, fallback to empty array
      let commentsArr = [];
      try {
        commentsArr = selectedItem.FinanceCommentsHistory ? JSON.parse(selectedItem.FinanceCommentsHistory) : [];
        if (!Array.isArray(commentsArr)) commentsArr = [];
      } catch {
        commentsArr = [];
      }
      let history = [];
      if (selectedItem.FinanceCommentsHistory) {
        try {
          const decodedJson = decodeHtmlEntities(selectedItem.FinanceCommentsHistory);
          history = JSON.parse(decodedJson);
          if (!Array.isArray(history)) history = [history];
        } catch {
          history = [];
        }
      }
      // Append new comment
      // const userRole = await getCurrentUserRole(context, selectedReq);
      history.push({
        Date: new Date().toISOString(),
        Title: 'Clarification',
        User: context.pageContext.user.displayName || 'Unknown User',
        // Role: userRole,
        Data: editFields.FinanceComments.trim(),
      });
      // Append the new clarification comment entry with FinanceComments content
      // commentsArr.push({
      //   Title: "Clarification",
      //   Date: new Date().toISOString(),
      //   User: context.pageContext.user.displayName,
      //   // Role: userRole,
      //   Data: editFields.FinanceComments.trim(),
      // });

      // Update SharePoint list item with updated JSON history and status fields
      await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id).update({
        FinanceCommentsHistory: JSON.stringify(history),
        FinanceStatus: "Clarification",
        PMStatus: "Pending",
        FinanceComments: editFields.FinanceComments.trim(),
        CurrentStatus: "Finance asked Clarification",
      });

      // Reload updated item data to refresh UI
      const updatedItem = await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id)();
      setSelectedItem(updatedItem);

      await sendFinanceClarificationEmail(updatedItem);
      clearAllFilters();
      showDialog("Clarification submitted successfully!", "success");
      // setTimeout(() => {
      //   setIsPanelOpen(false);
      //   setSelectedItem(null);
      // }, 500);
    } catch (error) {
      showDialog("Failed to submit clarification: " + (error as any)?.message, "error");
    }
  }
  // Handle file input change (Finance Attachments)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setAttachments(prev => [...prev, ...Array.from(e.target.files)]);
    }
  };

  async function handleSave() {
    if (!selectedItem) return;
    setLoading(true);
    setError(null);
    // const userRole = await getCurrentUserRole(context, selectedItem);
    try {

      // let historyArr = [];
      // try {
      //   historyArr = selectedItem.FinanceCommentsHistory ? JSON.parse(selectedItem.FinanceCommentsHistory) : [];
      //   if (!Array.isArray(historyArr)) historyArr = [];
      // } catch {
      //   historyArr = [];
      // }

      let history = [];
      if (selectedItem.FinanceCommentsHistory) {
        try {
          const decodedJson = decodeHtmlEntities(selectedItem.FinanceCommentsHistory);
          history = JSON.parse(decodedJson);
          if (!Array.isArray(history)) history = [history];
        } catch {
          history = [];
        }
      }
      // Append new comment entry if FinanceComments was updated
      if (editFields.FinanceComments && editFields.FinanceComments.trim()) {
        history.push({
          Date: new Date().toISOString(),
          Title: "Comment",
          User: context.pageContext.user.displayName,
          // Role: userRole,
          Data: editFields.FinanceComments.trim(),
        });
      }
      // let updatedFinanceStatus = editFields.FinanceStatus || selectedItem.FinanceStatus || "";
      // if ((editFields.Status || selectedItem.Status) === "Payment Received") {
      //   updatedFinanceStatus = "Paid";
      // } else {
      //   updatedFinanceStatus = "Pending";
      // }
      let newCurrentStatus: string;
      let newFinanceStatus: string;

      switch (editFields.Status) {
        case "Payment Received":
          newCurrentStatus = "Completed";
          newFinanceStatus = "Paid";
          break;
        case "Cancelled":
          newCurrentStatus = "Cancelled Request";
          newFinanceStatus = "Cancelled";
          break;
        case "Invoice Requested":
          newCurrentStatus = "Pending Finance Action";
          newFinanceStatus = "Pending";
          break;
        case "Invoice Raised":
          newCurrentStatus = "Pending Finance";
          newFinanceStatus = "Pending";
          break;
        default:
          newCurrentStatus = selectedItem.CurrentStatus ?? "Invoice Requested";
      }
      // Include updated FinanceCommentsHistory JSON string in update payload
      const updatePayload = {
        ...editFields,
        DueDate: editFields.DueDate
          ? (editFields.DueDate instanceof Date ? editFields.DueDate.toISOString() : editFields.DueDate)
          : null,
        FinanceCommentsHistory: JSON.stringify(history),
        FinanceComments: editFields.FinanceComments || "",
        FinanceStatus: newFinanceStatus,
        CurrentStatus: newCurrentStatus
      };

      const now = new Date().toISOString();
      let updatePayload2: any = { ...editFields };

      // Add dates when status is changed
      if (editFields.Status === 'Invoice Raised') {
        updatePayload2.InvoiceRaisedDate = now;
      }
      if (editFields.Status === 'Payment Completed' || editFields.Status === 'Payment Received') {
        updatePayload2.PaymentCompletedDate = now;
      }
      // Update the list item fields
      await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id).update(updatePayload);
      // await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id).update(editFields);
      if (attachments.length > 0) {
        for (const file of attachments) {
          const fileNameWithSuffix = `${file.name.replace(/\.[^/.]+$/, "")}_Finance${file.name.match(/\.[^/.]+$/)?.[0] || ""}`;
          const fileContent = await file.arrayBuffer();
          await sp.web.lists.getByTitle("Invoice Requests")
            .items.getById(selectedItem.Id)
            .attachmentFiles.add(fileNameWithSuffix, fileContent);
        }
      }
      const updatedItem = await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id)();
      if (originalStatus !== updatedItem.Status) {
        await sendPmStatusChangeEmail(updatedItem, updatedItem.Status ?? "", selectedItem.Status ?? "");
      }
      // await fetchData();
      setEditFields({});
      setAttachments([]);
      clearAllFilters();
      // Reload data to update UI
      // const fieldNames = [
      //   "Id",
      //   "PurchaseOrder",
      //   "ProjectName",
      //   "Status",
      //   "Comments",
      //   "POItem_x0020_Title",
      //   "POItem_x0020_Value",
      //   "InvoiceAmount",
      //   "Customer_x0020_Contact",
      //   "Modified",
      //   "Created",
      //   "FinanceStatus",
      // ];
      // const updatedItems = await sp.web.lists.getByTitle("Invoice Requests")
      //   .items.select(...fieldNames, "AttachmentFiles")
      //   .expand("AttachmentFiles")
      //   ();
      // setItems(updatedItems);
      // setIsPanelOpen(false);
      showDialog("Invoice request updated successfully!", "success");
    } catch (e: any) {
      setError("Update failed: " + (e.message ?? e));
      showDialog("Failed to update invoice request: " + (e.message ?? e), "error");
    }
    setLoading(false);
  }

  return (
    <section style={{ background: "#fff", borderRadius: 8, padding: 16 }}>
      <h2 style={{ fontWeight: 600, marginBottom: 16 }}>Update Invoice Request</h2>

      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" styles={{ root: { marginBottom: 20, fontSize: 10 } }} className={styles['compact-panel']}>
        <Stack.Item align="end"><Stack styles={{ root: { width: 140 } }}><Label>Search</Label>
          <TextField
            placeholder="Search"
            value={filters.search}
            onChange={(_, v) => setFilters(f => ({ ...f, search: v || "" }))}
          />
        </Stack></Stack.Item>
        {/* <Stack.Item align="end"><Stack styles={{ root: { width: 170 } }}><Label>Current Status</Label>
          <Dropdown
            multiSelect
            options={currentStatusOptions}              // includes { key: "All", text: "All" }
            selectedKeys={filterCurrentStatus}
            onChange={(e, option) => {
              if (!option) return;

              const key = option.key as string;

              if (key === "All") {
                // selecting All clears others
                setFilterCurrentStatus(["All"]);
              } else {
                setFilterCurrentStatus(prev => {
                  const withoutAll = prev.filter(k => k !== "All");
                  return option.selected
                    ? [...withoutAll, key]
                    : withoutAll.filter(k => k !== key);
                });
              }
            }}
            styles={dropdownStyles}
          />
        </Stack></Stack.Item>
        <Stack.Item align="end"><Stack styles={{ root: { width: 170 } }}><Label>Invoice Status</Label>
          <Dropdown
            multiSelect
            options={invoiceStatusOptions}              // All + ordered priority
            selectedKeys={filterInvoiceStatus}
            onChange={(e, option) => {
              if (!option) return;

              const key = option.key as string;

              if (key === "All") {
                setFilterInvoiceStatus(["All"]);
              } else {
                setFilterInvoiceStatus(prev => {
                  const withoutAll = prev.filter(k => k !== "All");
                  return option.selected
                    ? [...withoutAll, key]
                    : withoutAll.filter(k => k !== key);
                });
              }
            }}
            styles={dropdownStyles}
          /> */}
        <Stack.Item align="end">
          <Stack styles={{ root: { width: 170 } }}>
            <Label>Current Status</Label>
            <Dropdown
              multiSelect
              options={CURRENT_STATUS_OPTIONS}
              selectedKeys={filters.currentstatus}
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key.toString();

                setFilters(f => {
                  if (key === "All") {
                    return { ...f, currentstatus: ["All"] };
                  }

                  const prev = f.currentstatus.filter(k => k !== "All");
                  const next = option.selected
                    ? [...prev, key]
                    : prev.filter(k => k !== key);

                  return { ...f, currentstatus: next.length ? next : ["All"] };
                });
              }}
            />
          </Stack>
        </Stack.Item>

        <Stack.Item align="end">
          <Stack styles={{ root: { width: 170 } }}>
            <Label>Invoice Status</Label>
            <Dropdown
              multiSelect
              options={InvstatusOptions}
              selectedKeys={filters.status}
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key.toString();

                setFilters(f => {
                  if (key === "All") {
                    return { ...f, status: ["All"] };
                  }

                  const prev = f.status.filter(k => k !== "All");
                  const next = option.selected
                    ? [...prev, key]
                    : prev.filter(k => k !== key);

                  return { ...f, status: next.length ? next : ["All"] };
                });
              }}
            />
          </Stack>
        </Stack.Item>
        <Stack.Item align="end">
          <PrimaryButton
            text="Clear"
            onClick={clearFilters}
            disabled={
              !filters.search &&
              !filters.requestedDate &&
              !filters.customer &&
              (!filters.status || filters.status.length === 0 || filters.status.includes("All")) &&
              !filters.financeStatus &&
              (!filters.currentstatus || filters.currentstatus.length === 0 || filters.currentstatus.includes("All"))
            }
          // styles={{ root: { color: primaryColor } }}
          />
        </Stack.Item>
        <Stack.Item align="end" styles={{ root: { paddingLeft: 12 } }}>
          <IconButton
            iconProps={{ iconName: 'ExcelDocument' }}
            title="Export to Excel"
            ariaLabel="Export to Excel"
            onClick={handleExportToExcel}
            styles={{ root: { color: primaryColor } }}
          />
        </Stack.Item>
        <Stack.Item align="end" styles={{ root: { paddingLeft: 12 } }}>
          <IconButton
            iconProps={{ iconName: 'Columns' }}
            title="Manage Columns"
            ariaLabel="Manage Columns"
            onClick={() => setIsColumnPanelOpen(true)}
            styles={{ root: { color: primaryColor } }}
          />
        </Stack.Item>
      </Stack >


      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
      }
      {loading && <Spinner label="Loading Invoice Requests..." />}
      {
        !loading && (
          <>
            <div className={`ms-Grid-row ${styles.detailsListContainer}`}>
              <div style={{ height: 300, position: 'relative' }}>
                {/* <ScrollablePane> */}
                <div
                  className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.detailsList_Scrollablepane_Container}`}
                >
                  <div style={{ width: '100%', overflowX: 'auto' }}>
                    <DetailsList
                      items={filteredItems}
                      columns={getVisibleColumns()}
                      selection={selection}
                      selectionMode={SelectionMode.single}
                      setKey="financeViewList"
                      styles={{ root: { backgroundColor: "#fff", overflowX: 'auto' } }}
                      layoutMode={DetailsListLayoutMode.fixedColumns}
                      isHeaderVisible={true}
                      // onRenderRow={onRenderRow}
                      selectionPreservedOnEmptyClick={true}
                      onRenderDetailsHeader={onRenderDetailsHeader}
                    />
                  </div>
                </div>
                {columnFilterMenu.visible && (
                  <ContextualMenu
                    items={menuItems}
                    target={columnFilterMenu.target}
                    onDismiss={() => setColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                  />
                )}
                {/* </ScrollablePane> */}
              </div>
            </div>
          </>
        )
      }

      <Panel
        isOpen={isPanelOpen}
        onDismiss={handlePanelDismiss}
        headerText="Update Invoice Request"
        type={PanelType.custom}
        customWidth="1000px"
        isBlocking={true}
        isFooterAtBottom={false}
        styles={{
          main: {
            height: 'auto',
            margin: 'auto',
            borderRadius: 12,
          },
          scrollableContent: {
            overflowY: 'auto',
            padding: 5,
          }
        }}
      >
        {isClarificationPending && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={false} styles={{ root: { marginBottom: 12 } }}>
            Clarification has been requested. You cannot edit this request until it is Clarified.
          </MessageBar>
        )}
        {isPanelOpen && selectedItem && (
          <Stack
            horizontal
            styles={{ root: { height: 'calc(100vh - 150px)', overflow: 'hidden' } }}
            tokens={{ childrenGap: 20 }}
          >
            <Stack
              styles={{
                root: {
                  flexGrow: 1,
                  minWidth: 0,
                  maxWidth: '100%',
                  overflowY: 'auto',
                  padding: '24px',
                  background: '#fff',
                  borderRadius: 12,
                },
              }}
            >
              {/* Two-column form layout */}
              <Stack horizontal tokens={{ childrenGap: 36 }} styles={{ root: { width: '100%' } }}>
                {/* Left column */}
                <Stack tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 300, flex: 1 } }}>
                  <TextField label="Purchase Order" value={selectedItem?.PurchaseOrder || ''} disabled />
                  <TextField label="Project Name" value={selectedItem?.ProjectName || ''} disabled />
                  <TextField label="PO Item Title" value={selectedItem?.POItem_x0020_Title || ''} disabled />
                </Stack>
                {/* Right column */}
                <Stack tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 300, flex: 1 } }}>
                  <TextField label="Invoiced Amount" value={`${getCurrencySymbol(selectedItem.Currency)} ${selectedItem?.InvoiceAmount || ''}`} disabled />
                  <TextField label="Customer Contact" value={selectedItem?.Customer_x0020_Contact || ''} disabled />
                  <TextField label="PO Item Value" value={`${getCurrencySymbol(selectedItem.Currency)} ${selectedItem?.POItem_x0020_Value || ''}`} disabled />
                </Stack>
              </Stack>

              {/* New row for Invoice Due Date, Invoice Number and Invoice Status */}
              <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: 12, alignItems: 'flex-end' } }}>
                <DatePicker
                  label="Invoice Due Date"
                  value={editFields.DueDate ? new Date(editFields.DueDate) : undefined}
                  onSelectDate={date => handleFieldChange('DueDate', date ? date.toISOString() : '')}
                  styles={{ root: { flex: 1 } }}
                  disabled={isClarificationPending}
                />
                <TextField
                  label="Invoice Number"
                  value={editFields.InvoiceNumber || ''}
                  onChange={(e, val) => {
                    if (!invoiceNumberLoaded) handleFieldChange('InvoiceNumber', val || '');
                  }}
                  disabled={invoiceNumberLoaded || isClarificationPending}
                  styles={{ root: { flex: 1 } }}
                />
                <Dropdown
                  label="Invoice Status"
                  options={InvstatusOptions}
                  selectedKey={editFields.Status || selectedItem.Status || ''}
                  onChange={(_, option) => handleFieldChange('Status', option?.key)}
                  styles={{ root: { flex: 1 } }}
                  disabled={isClarificationPending}
                />
              </Stack>
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 16 } }}>
                {/* Comment fields in separate rows */}
                <TextField
                  label="Requestor Comments"
                  value={formatCommentHistory(selectedItem?.PMCommentsHistory) || ''}
                  multiline
                  rows={4}
                  disabled
                // styles={{ root: { backgroundColor: '#f3f2f1', marginTop: 24 } }}
                />
                <TextField
                  label="Finance Comments"
                  value={formatCommentHistory(selectedItem?.FinanceCommentsHistory) || ''}
                  multiline
                  rows={4}
                  disabled
                // styles={{ root: { backgroundColor: '#f3f2f1', marginTop: 12 } }}
                />
                <TextField
                  label="Finance Comments"
                  multiline
                  rows={5}
                  value={editFields.FinanceComments || ''}
                  onChange={(_, val) => handleFieldChange('FinanceComments', val || '')}
                  disabled={isClarificationPending}
                // styles={{ root: { marginTop: 12 } }}
                />
              </Stack>


              {/* Clarification button right below Finance Comments */}
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 12 }}>
                <PrimaryButton onClick={handleClarification} text="Ask Clarification" disabled={isClarificationPending} />
              </div>

              {/* Attachments - placed side by side in one row */}
              <Stack horizontal tokens={{ childrenGap: 36 }}>
                <Stack styles={{ root: { flex: 1 } }}>
                  <div style={{ fontWeight: '600', marginBottom: 8 }}>Requestor Attachments</div>
                  {pmAttachments.length ? (
                    <ul style={{ maxHeight: 140, overflowY: 'auto', paddingLeft: 20 }}>
                      {pmAttachments.map((att, i) => (
                        <li
                          key={i}
                          style={{ cursor: 'pointer', marginBottom: 6, display: 'flex', alignItems: 'center' }}
                          onClick={() => {
                            setViewerFileUrl(att.url);
                            setViewerFileName(att.name);
                            setIsDocPanelOpen(true);
                          }}
                        >
                          <span style={{ flexGrow: 1, color: '#0078d4', textDecoration: 'underline' }}>
                            {att.name}
                          </span>
                          <IconButton
                            iconProps={{ iconName: 'Download' }}
                            title={`Download ${att.name}`}
                            ariaLabel={`Download ${att.name}`}
                            onClick={(e) => {
                              e.stopPropagation();
                              const link = document.createElement('a');
                              link.href = att.url;
                              link.download = att.name;
                              link.click();
                            }}
                          // style={{ marginLeft: 12 }}
                          />
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <span style={{ color: '#888' }}>No Requestor attachments</span>
                  )}
                </Stack>
                <Stack styles={{ root: { flex: 1 } }}>
                  <div style={{ fontWeight: '600', marginBottom: 8 }}>Finance Attachments</div>
                  {financeAttachments.length ? (
                    <ul style={{ maxHeight: 140, overflowY: 'auto', paddingLeft: 20 }}>
                      {financeAttachments.map((att, i) => (
                        <li
                          key={i}
                          style={{ cursor: 'pointer', marginBottom: 6, display: 'flex', alignItems: 'center' }}
                          onClick={() => {
                            setViewerFileUrl(att.url);
                            setViewerFileName(att.name);
                            setIsDocPanelOpen(true);
                          }}
                        >
                          <span style={{ flexGrow: 1, color: '#0078d4', textDecoration: 'underline' }}>{att.name}</span>
                          <IconButton
                            iconProps={{ iconName: 'Download' }}
                            title={`Download ${att.name}`}
                            ariaLabel={`Download ${att.name}`}
                            onClick={(e) => {
                              e.stopPropagation();
                              const link = document.createElement('a');
                              link.href = att.url;
                              link.download = att.name;
                              link.click();
                            }}
                          // style={{ marginLeft: 12 }}
                          />
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <div style={{ color: '#888' }}>No Finance attachments</div>
                  )}
                </Stack>
              </Stack>

              {/* Drag and drop or click upload zone for new finance attachments */}
              <div
                onDrop={e => {
                  e.preventDefault();
                  const files = Array.from(e.dataTransfer.files);
                  if (files.length) setAttachments(files);
                  setIsDragActive(false);
                }}
                onDragOver={e => {
                  e.preventDefault();
                  setIsDragActive(true);
                }}
                onDragLeave={e => {
                  e.preventDefault();
                  setIsDragActive(false);
                }}
                onClick={() => document.getElementById('finance-attachment-input')?.click()}
                style={{
                  border: isDragActive ? '2px solid #0078d4' : '2px dashed #ccc',
                  borderRadius: 8,
                  padding: 20,
                  marginTop: 20,
                  cursor: 'pointer',
                  textAlign: 'center',
                  color: '#666',
                  userSelect: 'none',
                }}
              >
                <input
                  id="finance-attachment-input"
                  type="file"
                  multiple
                  accept="*/*"
                  style={{ display: 'none' }}
                  onChange={handleFileChange}
                />
                <i className='ms-Icon ms-Icon--Attach' style={{ fontSize: 46, color: '#aaa' }} aria-hidden="true"></i>
                <div style={{ marginTop: 12, fontWeight: 600 }}>Drop files here or click to upload</div>
                {attachments.length ? (
                  <div style={{ marginTop: 15, fontSize: 14, color: '#107c10' }}>
                    Selected: {attachments.map(f => f.name).join(', ')}
                  </div>
                ) : null}
              </div>

              {/* List of newly added finance attachments with preview and remove */}
              {attachments.length > 0 && (
                <ul>
                  {attachments.map((file, index) => (
                    <li key={index} className="attachmentRow">
                      <span className="attachmentFileName" style={{ flexGrow: 1, color: '#0078d4', textDecoration: 'underline', cursor: 'pointer' }}
                        onClick={() => { setViewerFileUrl(URL.createObjectURL(file)); setViewerFileName(file.name); }}>
                        {file.name}
                      </span>
                      <div className="attachmentButtons" style={{ display: 'flex', gap: '8px' }}>
                        <button onClick={e => { e.stopPropagation(); setViewerFileUrl(URL.createObjectURL(file)); setViewerFileName(file.name); setIsDocPanelOpen(true); }}>Preview</button>
                        <TooltipHost content="Remove attachment" id="remove-attachment-tooltip" calloutProps={{ gapSpace: 0 }} styles={{ root: { display: 'inline-block' } }}>
                          <button onClick={e => { e.stopPropagation(); setAttachments(prev => prev.filter((_, i) => i !== index)); }} style={{ background: 'transparent', border: 'none', color: '#a4262c', fontWeight: 'bold', cursor: 'pointer' }}>X</button>
                        </TooltipHost>
                      </div>
                    </li>
                  ))}
                </ul>
              )}

              <div style={{ height: 62 }}></div>

              {/* Submit button row */}
              <Stack horizontal tokens={{ childrenGap: 60 }} styles={{ root: { marginTop: 35, justifyContent: 'center' } }}>
                <PrimaryButton onClick={handleSave} text="Submit" disabled={loading || isClarificationPending} style={{ marginTop: 18 }} />
              </Stack>
            </Stack>
          </Stack>
        )}
        {/* Document viewer panel unchanged */}
        <Panel
          isOpen={isDocPanelOpen}
          onDismiss={handleDocPanelDismiss}
          type={PanelType.custom}
          customWidth="1000px"
          isBlocking={true}
          isFooterAtBottom={false}
          styles={{
            main: {
              height: 'auto',
              margin: 'auto',
              borderRadius: 12,
            },
            scrollableContent: {
              overflowY: 'auto',
              padding: 5,
            }
          }}
        >
          {isDocPanelOpen && (
            <Stack
              styles={{
                root: {
                  flexGrow: 1,
                  minWidth: 0,
                  maxWidth: '100%',
                  backgroundColor: '#f3f2f1',
                  borderTopRightRadius: 12,
                  borderBottomRightRadius: 12,
                  boxShadow: '-4px 0 8px rgba(0,0,0,0.1)',
                  position: 'relative',
                  display: 'flex',
                  flexDirection: 'column',
                  height: 'calc(100vh - 150px)',
                  overflow: 'hidden',
                  zIndex: 10,
                },
              }}
            >
              <div style={{ flexGrow: 1, overflow: 'auto' }}>
                <div style={{ height: "100%", width: "100%" }}>
                  <DocumentViewer
                    url={viewerFileUrl || ''}
                    isOpen={isDocPanelOpen}
                    onDismiss={() => {
                      setIsDocPanelOpen(false);
                      setViewerFileUrl(null);
                      setViewerFileName(null);
                    }}
                    fileName={viewerFileName || ''}
                  />
                </div>
              </div>
            </Stack>
          )}
        </Panel>
        <Dialog
          hidden={!dialogVisible}
          onDismiss={() => {
            setDialogVisible(false);
            setDialogMessage("");
          }}
          dialogContentProps={{
            type: dialogType === 'error' ? DialogType.largeHeader : DialogType.normal,
            title: dialogType === 'error' ? 'Error' : 'Success',
            subText: dialogMessage,
          }}
          modalProps={{ isBlocking: false }}
        >
          <DialogFooter>
            <div style={{ display: 'flex', justifyContent: 'center', width: '100%' }}>
              <PrimaryButton onClick={handleDialogClose} text="OK" />
            </div>
          </DialogFooter>
        </Dialog>
      </Panel>
      <Panel
        isOpen={isColumnPanelOpen}
        onDismiss={() => setIsColumnPanelOpen(false)}
        headerText="Customize Columns"
        type={PanelType.medium}
        isBlocking={true}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <div style={{ height: 400, overflow: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 12 }}>
            {columns.map((col: any) => (
              <div key={col.key} style={{
                display: 'flex',
                alignItems: 'center',
                padding: 12,
                marginBottom: 8,
                borderRadius: 4,
                backgroundColor: visibleColumns.includes(col.key as string) ? '#f3f2f1' : '#faf9f8'
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
            {/* <PrimaryButton
              text="Reset Default"
              onClick={() => {
                const defaults = ['PurchaseOrder', 'ProjectName', 'Status', 'InvoiceAmount', 'DueDate'];
                setVisibleColumns(defaults);
              }}
            /> */}
            {/* <PrimaryButton
              text="Apply Changes"
              onClick={() => setIsColumnPanelOpen(false)}
            /> */}
          </Stack>
        </Stack>
      </Panel>

      {/* Column Filter Panel */}
      <Panel
        isOpen={isFilterPanelOpen}
        onDismiss={() => setIsFilterPanelOpen(false)}
        headerText={
          currentFilterColumn
            ? `Filter: ${columns.find(c => c.key === currentFilterColumn)?.name || currentFilterColumn}`
            : 'Filter Column'
        }
        type={PanelType.smallFixedFar}
        isBlocking={true}
      >
        {currentFilterColumn && (
          <Stack tokens={{ childrenGap: 12 }}>
            <Label>Select values</Label>
            <div style={{ maxHeight: 300, overflowY: 'auto', border: '1px solid #edebe9', padding: 8, borderRadius: 4 }}>
              {getColumnDistinctValues(currentFilterColumn).map(val => {
                const selected = columnFilters[currentFilterColumn]?.includes(val) ?? false;
                return (
                  <div key={val} style={{ display: 'flex', alignItems: 'center', padding: '4px 0' }}>
                    <input
                      type="checkbox"
                      checked={selected}
                      onChange={e => {
                        setColumnFilters(prev => {
                          const prevForCol = prev[currentFilterColumn] || [];
                          let nextForCol: string[];
                          if (e.target.checked) {
                            nextForCol = [...prevForCol, val];
                          } else {
                            nextForCol = prevForCol.filter(v => v !== val);
                          }
                          return {
                            ...prev,
                            [currentFilterColumn]: nextForCol,
                          };
                        });
                      }}
                      style={{ marginRight: 8 }}
                    />
                    <span>{val}</span>
                  </div>
                );
              })}
              {getColumnDistinctValues(currentFilterColumn).length === 0 && (
                <span style={{ color: '#605e5c', fontStyle: 'italic' }}>No values available.</span>
              )}
            </div>

            <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="Clear"
                onClick={() => {
                  setColumnFilters(prev => {
                    const copy = { ...prev };
                    delete copy[currentFilterColumn];
                    return copy;
                  });
                }}
              />
              {/* <PrimaryButton
                text="Apply"
                onClick={() => setIsFilterPanelOpen(false)}
              /> */}
            </Stack>
          </Stack>
        )}
      </Panel>
    </section >

  );
}
