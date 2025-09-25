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
  DetailsListLayoutMode,
  IRenderFunction,
  Sticky,
  StickyPositionType,
  ScrollablePane,
  ContextualMenu,
  ContextualMenuItemType
} from "@fluentui/react";
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { SPFI } from "@pnp/sp";
import DocumentViewer from "../components/DocumentViewer";
import styles from './FinanceView.module.scss';
// import { columnSort, SortType } from './SortingUtils';
// import { update } from "@microsoft/sp-lodash-subset";
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
  // { key: "Request Draft", text: "Request Draft" },
  { key: "Not Generated", text: "Not Generated" },
  { key: "Invoice Raised", text: "Invoice Raised" },
  { key: "Pending Payment", text: "Pending Payment" },
  { key: "Payment Received", text: "Payment Received" },
  { key: "Cancelled", text: "Cancelled" }
];

export default function FinanceView({ sp, projectsp, context, initialFilters, onNavigate }: FinanceViewProps) {
  const [items, setItems] = useState<any[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedItem, setSelectedItem] = useState<any>(null);
  // const [isClarificationOpen,] = React.useState(false);
  const [isViewerOpen, setIsViewerOpen] = useState(false);
  const [viewerFileUrl, setViewerFileUrl] = useState<string | null>(null);
  const [viewerFileName, setViewerFileName] = useState<string | null>(null);
  const [originalStatus, setOriginalStatus] = useState<string | null>(null);
  const [invoiceNumberLoaded, setInvoiceNumberLoaded] = useState(false);
  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");
  const [dialogType, setDialogType] = useState<"success" | "error">("success");
  const [isDragActive, setIsDragActive] = useState(false);
  const [, setCustomerOptions] = useState<IDropdownOption[]>([]);
  const [statusOptions, setStatusOptions] = useState<IDropdownOption[]>([]);
  const [currentstatusOptions, setcurrentstatusOptions] = useState<IDropdownOption[]>([]);
  // const [filterCurrentStatus, setFilterCurrentStatus] = useState<string | undefined>(undefined);
  // const [isFilterPanelOpen, setIsFilterPanelOpen] = useState(false);
  // const [filteringColumnKey, setFilteringColumnKey] = useState<string | null>(null);

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
  const [filters, setFilters] = useState({
    search: initialFilters?.search || "",
    requestedDate: initialFilters?.requestedDate || null,
    customer: initialFilters?.customer || "",
    status: initialFilters?.Status || "",
    financeStatus: initialFilters?.FinanceStatus || "",
    currentstatus: initialFilters?.CurrentStatus || "",
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


  const [isPanelOpen, setIsPanelOpen] = useState(false);

  const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
    const sortedItems = [...items].sort((a, b) => {
      let aVal = (a as any)[columnKey];
      let bVal = (b as any)[columnKey];

      // Handle null/undefined
      if (aVal == null && bVal == null) return 0;
      if (aVal == null) return 1;
      if (bVal == null) return -1;

      // Handle Date objects
      if (aVal instanceof Date) aVal = aVal.getTime();
      if (bVal instanceof Date) bVal = bVal.getTime();

      // Number comparison (after Date conversion)
      if (typeof aVal === 'number' && typeof bVal === 'number') {
        return direction === 'asc' ? aVal - bVal : bVal - aVal;
      }

      // Try parsing as date strings if not a number
      const aAsDate = Date.parse(aVal);
      const bAsDate = Date.parse(bVal);
      if (!isNaN(aAsDate) && !isNaN(bAsDate)) {
        return direction === 'asc' ? aAsDate - bAsDate : bAsDate - aAsDate;
      }

      // Default to string comparison
      const aStr = aVal.toString();
      const bStr = bVal.toString();
      return direction === 'asc'
        ? aStr.localeCompare(bStr)
        : bStr.localeCompare(aStr);
    });

    setItems(sortedItems);
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  };
  const menuItems = [
    { key: 'asc', text: 'Sort Asc to Desc', iconProps: { iconName: 'SortUp' }, onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc') },
    { key: 'desc', text: 'Sort Desc to Asc', iconProps: { iconName: 'SortDown' }, onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc') },
    { key: 'divider', itemType: ContextualMenuItemType.Divider },
    // { key: 'filter', text: 'Filter...', iconProps: { iconName: 'Filter' }, onClick: () => openFilterPanelFromMenu() },
    // { key: 'clear', text: 'Clear Filter', iconProps: { iconName: 'ClearFilter' }, onClick: () => clearColumnFilter(columnFilterMenu.columnKey!) },
  ];

  // const clearColumnFilter = (columnKey: string) => {
  //   setFilters((prev) => ({ ...prev, [columnKey]: undefined }));
  //   setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  // };
  // const openFilterPanelFromMenu = () => {
  //   setIsFilterPanelOpen(true);
  //   setFilteringColumnKey(columnFilterMenu.columnKey);
  //   setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  // };

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
        { key: "PurchaseOrder", name: "Purchase Order", fieldName: "PurchaseOrder", minWidth: 80, maxWidth: 130, isResizable: true, onColumnClick: onColumnHeaderClick, },
        { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 120, maxWidth: 170, isResizable: true, onColumnClick: onColumnHeaderClick, },
        {
          key: "CurrentStatus",
          name: "Current Status",
          fieldName: "CurrentStatus",
          minWidth: 150,
          maxWidth: 200, isResizable: true,
          onRender: (item) => item.CurrentStatus || "-",
          onColumnClick: onColumnHeaderClick,
        },
        { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 150, maxWidth: 200, isResizable: true, onColumnClick: onColumnHeaderClick, },
        // { key: "Currency", name: "Currency", fieldName: "Currency", minWidth: 150, maxWidth: 200, isResizable: true, onColumnClick: onColumnHeaderClick, },
        { key: "DueDate", name: "DueDate", fieldName: "DueDate", minWidth: 150, maxWidth: 200, isResizable: true, onRender: item => item.DueDate ? new Date(item.DueDate).toLocaleDateString() : "-", onColumnClick: onColumnHeaderClick, },
        { key: "Comments", name: "PM Comments", fieldName: "Comments", minWidth: 160, maxWidth: 300, isResizable: true, onColumnClick: onColumnHeaderClick, },
        { key: "POItem_x0020_Title", name: "PO Item Title", fieldName: "POItem_x0020_Title", minWidth: 120, maxWidth: 170, isResizable: true, onColumnClick: onColumnHeaderClick, },
        {
          key: "POItem_x0020_Value", name: "PO Item Value", fieldName: "POItem_x0020_Value", minWidth: 100, maxWidth: 140, isResizable: true, onColumnClick: onColumnHeaderClick, onRender: (item: any) => {
            if (item.POItem_x0020_Value != null && !isNaN(Number(item.POItem_x0020_Value))) {
              return `${Number(item.POItem_x0020_Value).toLocaleString()} ${item.Currency ?? ''}`.trim();
            }
            return '';
          }
        },
        {
          key: "InvoiceAmount", name: "Invoiced Amount", fieldName: "InvoiceAmount", minWidth: 100, maxWidth: 140, isResizable: true, onColumnClick: onColumnHeaderClick, onRender: (item: any) => {
            if (item.InvoiceAmount != null && !isNaN(Number(item.InvoiceAmount))) {
              return `${Number(item.InvoiceAmount).toLocaleString()} ${item.Currency ?? ''}`.trim();
            }
            return '';
          }
        },
        { key: "Customer_x0020_Contact", name: "Customer Contact", fieldName: "Customer_x0020_Contact", minWidth: 120, maxWidth: 170, isResizable: true, onColumnClick: onColumnHeaderClick, },
        {
          key: "Created", name: "Created", fieldName: "Created", minWidth: calculateWidth("Created"), maxWidth: 300, isResizable: true, onRender: item => new Date(item.Created).toLocaleDateString(), onColumnClick: onColumnHeaderClick,
        },
        {
          key: "CreatedBy", name: "Created By", fieldName: "Author", minWidth: calculateWidth("Created By"), maxWidth: 300, isResizable: true,
          onRender: item => item.Author?.Title || "-",
          onColumnClick: onColumnHeaderClick,
        },
        {
          key: "Modified", name: "Modified", fieldName: "Modified", minWidth: calculateWidth("Modified"), maxWidth: 300, isResizable: true,
          onRender: item => new Date(item.Modified).toLocaleDateString(),
          onColumnClick: onColumnHeaderClick,
        },
        {
          key: "ModifiedBy", name: "Modified By", fieldName: "Editor", minWidth: calculateWidth("Modified By"), maxWidth: 300, isResizable: true,
          onRender: item => item.Editor?.Title || "-",
          onColumnClick: onColumnHeaderClick,
        },
      ];
      setColumns(cols);
      const listItems = await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.select(...fieldNames, "AttachmentFiles")
        .expand("AttachmentFiles", "Author", "Editor")
        .top(500)();

      setItems(listItems);

      setCustomerOptions(Array.from(new Set(listItems.map(i => i.Customer).filter(Boolean))).map(val => ({ key: val, text: val })));
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

  const showDialog = (message: string, type: "success" | "error") => {
    setDialogMessage(message);
    setDialogType(type);
    setDialogVisible(true);
  };

  const onDialogOk = async () => {
    setDialogVisible(false);
    setIsPanelOpen(false);
    if (dialogType === "success") {
      // Reload data after success dialog dismiss
      await fetchData();
    }
  };

  const handlePanelDismiss = () => {
    setIsPanelOpen(false);
    setAttachments([]);  // clear attachments on close
    setEditFields({});   // optional: reset form fields too
    setPmAttachments([]);
    setSelectedItem(null);
    // setIsViewerOpen(false);
  };

  const clearFilters = () => {
    setFilters({
      search: "",
      requestedDate: null,
      customer: "",
      status: "",
      financeStatus: "",
      currentstatus: "",
    });
  };

  const handleExportToExcel = () => {

    if (!editFields.FinanceComments || editFields.FinanceComments.trim() === "") {
      setDialogMessage("No available Data to export");
      setDialogType("error");
      setDialogVisible(true);
      return;
    }

    const exportData = filteredItems.map(item => ({
      PurchaseOrder: item.PurchaseOrder,
      ProjectName: item.ProjectName,
      CurrentStatus: item.CurrentStatus || "-",
      Status: item.Status,
      Comments: item.Comments,
      POItemTitle: item.POItem_x0020_Title,
      POItemValue: item.POItem_x0020_Value,
      InvoiceAmount: item.InvoiceAmount,
      CustomerContact: item.Customer_x0020_Contact,
      Created: item.Created ? new Date(item.Created).toLocaleDateString() : '',
      CreatedBy: item.Author?.Title || "-",
      Modified: item.Modified ? new Date(item.Modified).toLocaleDateString() : '',
      ModifiedBy: item.Editor?.Title || "-"
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'InvoiceRequests');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), `InvoiceRequests_${new Date().toISOString()}.xlsx`);
  };


  // Open the panel and select item
  // const closeDocViewer = (item: any) => {
  //   setSelectedItem(item);
  //   setIsViewerOpen(false);  // initially no viewer open
  //   setIsPanelOpen(true);
  // };

  const filteredItems = React.useMemo(() => {
    const searchText = filters.search?.toLowerCase() || "";

    return items.filter(item => {
      const matchesSearch =
        !searchText ||
        columns.some(col => {
          const fieldValue = item[col.fieldName] ?? "";
          return fieldValue.toString().toLowerCase().includes(searchText);
        });

      return (
        matchesSearch &&
        (!filters.customer || item.Customer === filters.customer) &&
        (!filters.status || item.Status === filters.status) &&
        (!filters.financeStatus || item.FinanceStatus === filters.financeStatus) &&
        (!filters.currentstatus || item.CurrentStatus === filters.currentstatus) &&
        (!filters.requestedDate || (item.RequestedDate && new Date(item.RequestedDate).toLocaleDateString() === filters.requestedDate.toLocaleDateString()))
      );
    });
  }, [items, columns, filters]);


  useEffect(() => {
    setCustomerOptions(getUniqueOptions(filteredItems, "Customer"));
    setStatusOptions(getUniqueOptions(filteredItems, "Status"));
    setcurrentstatusOptions(getUniqueOptions(filteredItems, "CurrentStatus"));
  }, [filteredItems]);

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

      return arr.map((entry: any) => {
        const dateObj = entry.Date ? new Date(entry.Date) : null;
        const date = dateObj ? dateObj.toLocaleDateString() : "";
        const time = dateObj ? dateObj.toLocaleTimeString() : "";
        const title = entry.Title || entry.title || "";
        const user = entry.User || "";
        const role = entry.Role ? ` (${entry.Role})` : "";
        const data = entry.Data || entry.comment || "";
        return `[${date} ${time}]${user}${role} \n${title}:${data}`;
      }).join("\n\n");

    } catch (err) {
      console.error("Failed to format comment history", err, jsonStr);
      return "";
    }
  }

  async function sendFinanceClarificationEmail(item: any) {
    if (!item) return;

    const siteUrl = context.pageContext.web.absoluteUrl;
    const listName = "Invoice Requests";

    const itemUrl = `${siteUrl}/Lists/${listName}/DispForm.aspx?ID=${item.Id}`;

    const emailProps = {
      To: ["srushti.hunalli@sacha.solutions"], // Replace with actual finance email
      Subject: `Clarification submitted for Invoice Request: PO ${item.PurchaseOrder}`,
      Body: `
      A clarification has been submitted on the following invoice request:<br/><br/>
      <b>Purchase Order:</b> ${item.PurchaseOrder}<br/>
      <b>Project Name:</b> ${item.ProjectName ?? "N/A"}<br/>
      <b>PO Item Title:</b> ${item.POItem_x0020_Title ?? "N/A"}<br/>
      <b>Finance Comments:</b> ${item.FinanceComments ?? "N/A"}<br/><br/>
      Please review the clarification <a href="${itemUrl}">here</a>.
    `,
      AdditionalHeaders: {
        "content-type": "text/html",
      },
    };

    try {
      // Use PnP to send email via SharePoint utility
      await sp.utility.sendEmail(emailProps);
    } catch (error) {
      console.error("Failed to send finance clarification email", error);
    }
  }

  async function sendPmStatusChangeEmail(item: any, oldStatus: string, newStatus: string) {
    if (!item) return;

    const siteUrl = context.pageContext.web.absoluteUrl;
    const listName = "Invoice Requests";
    const itemUrl = `${siteUrl}/Lists/${listName}/DispForm.aspx?ID=${item.Id}`;

    const emailProps = {
      To: ["Srushti.hunalli@sacha.solutions"], // Replace with actual PM email
      Subject: `Invoice Request Status Changed: PO ${item.PurchaseOrder}`,
      Body: `
      The status of the following invoice request has changed:<br/><br/>
      <b>Purchase Order:</b> ${item.PurchaseOrder}<br/>
      <b>Project Name:</b> ${item.ProjectName ?? "N/A"}<br/>
      <b>PO Item Title:</b> ${item.POItem_x0020_Title ?? "N/A"}<br/>
      <b>Previous Status:</b> ${oldStatus}<br/>
      <b>New Status:</b> ${newStatus}<br/><br/>
      You can view the invoice request <a href="${itemUrl}">here</a>.
    `,
      AdditionalHeaders: {
        "content-type": "text/html",
      },
    };

    try {
      await sp.utility.sendEmail(emailProps);
    } catch (error) {
      console.error("Failed to send PM status change email", error);
    }
  }

  async function loadPmAttachments(item: any) {
    if (!item) {
      setPmAttachments([]);
      return;
    }

    const attachments = item.AttachmentFiles || [];
    const pmAttachments = attachments
      .filter((att: any) => att.FileName.match(/_PM(\.[^.]*)?$/i))
      .map((att: any) => ({ name: att.FileName, url: att.ServerRelativeUrl }));

    setPmAttachments(pmAttachments);
  }

  // Open edit panel and load PM attachments
  function openEditForm(item: any) {
    if (!item) return;
    setInvoiceNumberLoaded(!!item.InvoiceNumber);

    // Determine the invoice status to use in the form:
    const normalizedStatus = (item.Status || "").trim();
    const defaultStatusForSubmitted = "Not Generated";
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

    setOriginalStatus(item.Status ?? null);
    setAttachments([]);
    setIsPanelOpen(true);
    loadPmAttachments(item);
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

      // Append the new clarification comment entry with FinanceComments content
      commentsArr.push({
        Title: "Clarification",
        Date: new Date().toISOString(),
        User: context.pageContext.user.displayName,
        // Role: userRole,
        Data: editFields.FinanceComments.trim(),
      });

      // Update SharePoint list item with updated JSON history and status fields
      await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id).update({
        FinanceCommentsHistory: JSON.stringify(commentsArr),
        FinanceStatus: "Clarification",
        PMStatus: "Pending",
        FinanceComments: editFields.FinanceComments.trim(),
        CurrentStatus: "Finance asked Clarification",
      });

      // Reload updated item data to refresh UI
      const updatedItem = await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id)();
      setSelectedItem(updatedItem);

      await sendFinanceClarificationEmail(updatedItem);

      showDialog("Clarification submitted successfully!", "success");
    } catch (error) {
      showDialog("Failed to submit clarification: " + (error as any)?.message, "error");
    }
  }
  // Handle file input change (Finance Attachments)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setAttachments(Array.from(e.target.files)); // attachments becomes File[]
    }
  };


  // Save updated invoice request and upload finance attachments

  async function handleSave() {
    if (!selectedItem) return;
    setLoading(true);
    setError(null);
    // const userRole = await getCurrentUserRole(context, selectedItem);
    try {

      let historyArr = [];
      try {
        historyArr = selectedItem.FinanceCommentsHistory ? JSON.parse(selectedItem.FinanceCommentsHistory) : [];
        if (!Array.isArray(historyArr)) historyArr = [];
      } catch {
        historyArr = [];
      }

      // Append new comment entry if FinanceComments was updated
      if (editFields.FinanceComments && editFields.FinanceComments.trim()) {
        historyArr.push({
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
        case "Not Generated":
          newCurrentStatus = "Pending Finance Action";
          newFinanceStatus = "Pending";
          break;
        case "Invoice Raised":
          newCurrentStatus = "Finance In Progress";
          newFinanceStatus = "Pending";
          break;
        default:
          newCurrentStatus = selectedItem.CurrentStatus ?? "Not Generated";
      }
      // Include updated FinanceCommentsHistory JSON string in update payload
      const updatePayload = {
        ...editFields,
        FinanceCommentsHistory: JSON.stringify(historyArr),
        FinanceComments: editFields.FinanceComments || "",
        FinanceStatus: newFinanceStatus,
        CurrentStatus: newCurrentStatus
      };

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

      if (originalStatus !== selectedItem.Status) {
        const updatedItem = await sp.web.lists.getByTitle("Invoice Requests").items.getById(selectedItem.Id)();
        await sendPmStatusChangeEmail(updatedItem, originalStatus ?? "", selectedItem.Status ?? "");
      }

      setIsPanelOpen(false);
      setEditFields({});
      setAttachments([]);
      showDialog("Invoice request updated successfully!", "success");
      // Reload data to update UI
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
      ];
      const updatedItems = await sp.web.lists.getByTitle("Invoice Requests")
        .items.select(...fieldNames, "AttachmentFiles")
        .expand("AttachmentFiles")
        .top(500)();
      setItems(updatedItems);

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
        <Stack.Item align="end"><Stack styles={{ root: { width: 140 } }}><Label>Current Status</Label>
          <Dropdown
            placeholder="Current Status"
            options={currentstatusOptions}
            selectedKey={filters.currentstatus}
            onChange={(_, option) => setFilters(f => ({ ...f, currentstatus: (option?.key ?? "").toString() }))}
          />
        </Stack></Stack.Item>
        <Stack.Item align="end"><Stack styles={{ root: { width: 140 } }}><Label>Invoice Status</Label>
          <Dropdown
            placeholder="Invoice Status"
            options={statusOptions}
            selectedKey={filters.status}
            onChange={(_, option) => setFilters(f => ({ ...f, status: (option?.key ?? "").toString() }))}
          />
        </Stack></Stack.Item>
        <Stack.Item align="end">
          <PrimaryButton
            text="Clear"
            onClick={clearFilters}
            disabled={
              !filters.search &&
              !filters.requestedDate &&
              !filters.customer &&
              !filters.status &&
              !filters.financeStatus &&
              !filters.currentstatus
            }
          />
        </Stack.Item>
        <Stack.Item align="end" styles={{ root: { paddingLeft: 12 } }}>
          <IconButton
            iconProps={{ iconName: 'ExcelDocument' }}
            title="Export to Excel"
            ariaLabel="Export to Excel"
            onClick={handleExportToExcel}
          />
        </Stack.Item>
      </Stack>


      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {loading && <Spinner label="Loading Invoice Requests..." />}
      {!loading && (
        <>
          <div className={`ms-Grid-row ${styles.detailsListContainer}`}>
            <div style={{ height: 300, position: 'relative' }}>
              <ScrollablePane>
                <div
                  className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.detailsList_Scrollablepane_Container}`}
                >
                  <DetailsList
                    items={filteredItems}
                    columns={columns}
                    selection={selection}
                    selectionMode={SelectionMode.single}
                    setKey="financeViewList"
                    styles={{ root: { backgroundColor: "#fff" } }}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    // onRenderRow={onRenderRow}
                    selectionPreservedOnEmptyClick={true}
                    onRenderDetailsHeader={onRenderDetailsHeader}
                  />
                </div>
                {columnFilterMenu.visible && (
                  <ContextualMenu
                    items={menuItems}
                    target={columnFilterMenu.target}
                    onDismiss={() => setColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                  />
                )}
              </ScrollablePane>
            </div>
          </div>
        </>
      )}

      <Panel
        isOpen={isPanelOpen}
        onDismiss={handlePanelDismiss}
        headerText="Update Invoice Request"
        type={PanelType.custom}
        customWidth="1000px"
        isBlocking={false}
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
        {isPanelOpen && selectedItem && (
          <Stack
            horizontal
            styles={{ root: { height: 'calc(100vh - 150px)', overflow: 'hidden' } }}
            tokens={{ childrenGap: 20 }}
          >
            {!isViewerOpen && (
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
                    <TextField
                      label="Invoice Number"
                      value={editFields.InvoiceNumber || ''}
                      onChange={(e, val) => {
                        if (!invoiceNumberLoaded) handleFieldChange('InvoiceNumber', val || '');
                      }}
                      disabled={invoiceNumberLoaded}
                    />
                    <DatePicker
                      label="Due Date"
                      value={editFields.DueDate ? new Date(editFields.DueDate) : undefined}
                      onSelectDate={date => handleFieldChange('DueDate', date ? date.toISOString() : '')}
                    />
                    <TextField
                      label="Previous PM Comments"
                      value={formatCommentHistory(selectedItem?.PMCommentsHistory) || ''}
                      multiline
                      rows={4}
                      disabled
                      styles={{ root: { backgroundColor: '#f3f2f1' } }}
                    />
                  </Stack>
                  {/* Right column */}
                  <Stack tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 300, flex: 1 } }}>
                    <TextField label="Invoiced Amount" value={selectedItem?.InvoiceAmount?.toString() || ''} disabled />
                    <TextField label="Customer Contact" value={selectedItem?.Customer_x0020_Contact || ''} disabled />
                    <TextField label="PO Item Value" value={selectedItem?.POItem_x0020_Value || ''} disabled />
                    <Dropdown
                      label="Invoice Status"
                      options={InvstatusOptions}
                      selectedKey={editFields.Status || selectedItem.Status || ''}
                      onChange={(_, option) => handleFieldChange('Status', option?.key)}
                    />
                    <div style={{ height: 62 }}></div>
                    <TextField
                      label="Previous Finance Comments"
                      value={formatCommentHistory(selectedItem?.FinanceCommentsHistory) || ''}
                      multiline
                      rows={4}
                      disabled
                      styles={{ root: { backgroundColor: '#f3f2f1' } }}
                    />
                  </Stack>
                </Stack>
                {/* Below columns */}
                <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 24 } }}>
                  <TextField
                    label="Finance Comments"
                    multiline
                    rows={5}
                    value={editFields.FinanceComments || ''}
                    onChange={(_, val) => handleFieldChange('FinanceComments', val || '')}
                  />

                  <div>PM Attachments</div>
                  {pmAttachments.length ? (
                    <ul style={{ maxHeight: 140, overflowY: 'auto', paddingLeft: 20 }}>
                      {pmAttachments.map((att, i) => (
                        <li
                          key={i}
                          style={{ cursor: 'pointer', marginBottom: 6, display: 'flex', alignItems: 'center' }}
                          onClick={() => {
                            setViewerFileUrl(att.url);
                            setViewerFileName(att.name);
                            setIsViewerOpen(true);
                          }}
                        >
                          <span style={{ flexGrow: 1, color: '#0078d4', textDecoration: 'underline' }}>
                            {att.name}
                          </span>
                          <a
                            href={att.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            style={{ marginLeft: 12, color: '#605e5c', fontSize: 12 }}
                          >
                            Download
                          </a>
                          <button
                            onClick={e => {
                              e.stopPropagation();
                              setIsViewerOpen(false);
                            }}
                            style={{
                              marginLeft: 8,
                              background: 'transparent',
                              border: 'none',
                              color: '#a4262c',
                              cursor: 'pointer',
                              fontWeight: 'bold',
                            }}
                            aria-label={`Clear preview of ${att.name}`}
                          >
                            ×
                          </button>
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <span style={{ color: '#888' }}>No PM attachments</span>
                  )}

                  <div style={{ marginTop: 20 }}>Finance Attachments</div>
                  <div
                    onDrop={e => {
                      e.preventDefault();
                      const files = Array.from(e.dataTransfer.files).filter(file =>
                        ['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(file.type)
                      );
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
                      marginTop: 8,
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
                      accept={'.pdf,.xls,.xlsx,application/pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel'}
                      style={{ display: 'none' }}
                      // onChange={e => {
                      //   if (e.target.files) setAttachments(Array.from(e.target.files));
                      // }}
                      onChange={handleFileChange}
                    />
                    <i className='ms-Icon ms-Icon--Attach' style={{ fontSize: 46, color: '#aaa' }} aria-hidden="true"></i>
                    <div style={{ marginTop: 12, fontWeight: 600 }}>Drop files here or click to upload (PDF/XLSX)</div>
                    {attachments.length ? (
                      <>
                        <div style={{ marginTop: 15, fontSize: 14, color: '#107c10' }}>
                          Selected: {attachments.map(f => f.name).join(', ')}
                        </div>
                      </>
                    ) : null}
                  </div>
                  {/* List of newly added finance attachments with preview and remove */}
                  {attachments.length > 0 && (
                    <ul style={{ maxHeight: 140, overflowY: 'auto', paddingLeft: 20 }}>
                      {attachments.map((file, index) => (
                        <li
                          key={index}
                          style={{ display: 'flex', alignItems: 'center', marginBottom: 6 }}
                        >
                          <span style={{ flexGrow: 1, color: '#0078d4', cursor: 'pointer', textDecoration: 'underline' }}
                            onClick={() => {
                              setViewerFileUrl(URL.createObjectURL(file));
                              setViewerFileName(file.name);
                              setIsViewerOpen(true);
                            }}>
                            {file.name}
                          </span>
                          <button onClick={(e) => {
                            e.stopPropagation();
                            const objectUrl = URL.createObjectURL(file); // 'file' is the variable holding an actual File or Blob
                            setViewerFileUrl(objectUrl);
                            setViewerFileName(file.name);
                            setIsViewerOpen(true);
                          }}>
                            Preview
                          </button>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              // Remove this file from attachments
                              setAttachments((prev) =>
                                prev.filter((_, i) => i !== index)
                              );
                              // If preview open and currently previewing this file, close viewer
                              if (viewerFileName === file.name) {
                                setIsViewerOpen(false);
                                setViewerFileUrl(null);
                                setViewerFileName(null);
                              }
                            }}
                            style={{
                              marginLeft: 8,
                              background: 'transparent',
                              border: 'none',
                              color: '#a4262c',
                              cursor: 'pointer',
                              fontWeight: 'bold',
                            }}
                            aria-label={`Remove ${file.name}`}
                          >
                            ×
                          </button>
                        </li>
                      ))}
                    </ul>
                  )}
                </Stack>
                <div style={{ height: 62 }}></div>
                {/* Buttons */}
                <div style={{ height: 28, marginTop: 25 }} />
                <Stack horizontal tokens={{ childrenGap: 60 }} styles={{ root: { marginTop: 25, justifyContent: 'center' } }}>
                  <PrimaryButton onClick={handleClarification} text="Ask Clarification" />
                  <PrimaryButton onClick={handleSave} text="Submit" disabled={loading} />
                </Stack>
              </Stack>
            )}

            {isViewerOpen && (
              // ---- Full-width Document Viewer ----
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
                <PrimaryButton
                  style={{ position: 'absolute', top: 8, right: 8, zIndex: 20 }}
                  iconProps={{ iconName: 'Cancel' }}
                  title="Close Viewer"
                  ariaLabel="Close Viewer"
                  onClick={() => {
                    setIsViewerOpen(false);
                    setViewerFileUrl(null);
                    setViewerFileName(null);
                    // selectedItem and panel remain unchanged
                  }}
                />
                <div style={{ flexGrow: 1, overflow: 'auto' }}>
                  <DocumentViewer
                    url={viewerFileUrl || ''}
                    isOpen={isViewerOpen}
                    onDismiss={() => {
                      setIsViewerOpen(false);
                      setViewerFileUrl(null);
                      setViewerFileName(null);
                    }}
                    fileName={viewerFileName || ''}
                  />
                </div>
              </Stack>
            )}
          </Stack>
        )}

      </Panel>
      <Dialog
        hidden={!dialogVisible}
        onDismiss={() => setDialogVisible(false)}
        dialogContentProps={{
          type: dialogType === 'error' ? DialogType.largeHeader : DialogType.normal,
          title: dialogType === 'error' ? 'Error' : 'Success',
          subText: dialogMessage,
        }}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter>
          <PrimaryButton onClick={onDialogOk} text="OK" />
        </DialogFooter>
      </Dialog>
    </section>

  );
}
