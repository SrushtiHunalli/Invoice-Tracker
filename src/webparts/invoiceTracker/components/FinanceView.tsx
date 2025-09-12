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
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import DocumentViewer from "../components/DocumentViewer";
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
  };
  onNavigate: (pageKey: string, params?: any) => void;
  projectsp: SPFI;
}

// STATUS OPTIONS (STEP LABELS)
const statusOptions: IDropdownOption[] = [
  // { key: "Request Draft", text: "Request Draft" },
  { key: "Not Generated", text: "Not Generated" },
  { key: "Pending Payment", text: "Pending Payment" },
  { key: "Payment Received", text: "Payment Received" },
];

// async function ensureFolder(sp: SPFI, parentFolderPath: string, folderName: string): Promise<string> {
//   const fullPath = `${parentFolderPath}/${folderName}`;
//   try {
//     const folderInfo = await sp.web.getFolderByServerRelativePath(fullPath).select("Exists", "ServerRelativeUrl")();
//     if (folderInfo.Exists) {
//       console.log(`Folder exists: ${fullPath}`);
//       return fullPath;
//     } else {
//       const newFolder = await sp.web.getFolderByServerRelativePath(parentFolderPath).addSubFolderUsingPath(folderName);
//       const newFolderInfo = await newFolder.select("ServerRelativeUrl")();
//       console.log(`Folder created: ${newFolderInfo.ServerRelativeUrl}`);
//       return newFolderInfo.ServerRelativeUrl;
//     }
//   } catch (error) {
//     console.error("Error checking/creating folder:", fullPath, error);
//     throw error;
//   }
// }

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

  const [filters, setFilters] = useState({
    search: initialFilters?.search || "",
    requestedDate: initialFilters?.requestedDate || null,
    customer: initialFilters?.customer || "",
    status: initialFilters?.Status || "",
    financeStatus: initialFilters?.FinanceStatus || "",
  });

  const financeStatusOptions: IDropdownOption[] = [
    { key: "Submitted", text: "Submitted" },
    { key: "Pending", text: "Pending" },
    { key: "Clarification", text: "Clarification" },
    { key: "Paid", text: "Paid" },
  ];

  const [, setCustomerOptions] = useState<IDropdownOption[]>([]);

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
          openEditForm(sel);
        }
      }
    })
  );

  const [isPanelOpen, setIsPanelOpen] = useState(false);

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
        "InvoiceNumber"
      ];
      const cols: IColumn[] = [
        { key: "PurchaseOrder", name: "Purchase Order", fieldName: "PurchaseOrder", minWidth: 80, maxWidth: 130 },
        { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 120, maxWidth: 170 },
        {
          key: "Finance Status",
          name: "Current Status",
          fieldName: "FinanceStatus",
          minWidth: 150,
          maxWidth: 200,
          onRender: (item) => item.FinanceStatus || "-"
        },
        { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 150, maxWidth: 200 },
        { key: "Comments", name: "PM Comments", fieldName: "Comments", minWidth: 160, maxWidth: 300 },
        { key: "POItem_x0020_Title", name: "PO Item Title", fieldName: "POItem_x0020_Title", minWidth: 120, maxWidth: 170 },
        { key: "POItem_x0020_Value", name: "PO Item Value", fieldName: "POItem_x0020_Value", minWidth: 100, maxWidth: 140 },
        { key: "InvoiceAmount", name: "Invoice Amount", fieldName: "InvoiceAmount", minWidth: 100, maxWidth: 140 },
        { key: "Customer_x0020_Contact", name: "Customer Contact", fieldName: "Customer_x0020_Contact", minWidth: 120, maxWidth: 170 },
      ]; // your columns setup
      setColumns(cols);

      const listItems = await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.select(...fieldNames, "AttachmentFiles")
        .expand("AttachmentFiles")
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
    });
  };

  // Open the panel and select item
  // const closeDocViewer = (item: any) => {
  //   setSelectedItem(item);
  //   setIsViewerOpen(false);  // initially no viewer open
  //   setIsPanelOpen(true);
  // };


  // const filteredItems = items.filter(item => {
  //   return (
  //     (!filters.search ||
  //       (item.ProjectName || "").toLowerCase().includes(filters.search.toLowerCase()) ||
  //       (item.PurchaseOrder || "").toLowerCase().includes(filters.search.toLowerCase())) ||
  //       (item.POItem_x0020_Title || "").toLowerCase().includes(filters.search.toLowerCase()) ||
  //       (item.POItem_x0020_Value || "").toString().toLowerCase().includes(filters.search.toLowerCase()) ||
  //       (item.InvoiceAmount || "").toString().toLowerCase().includes(filters.search.toLowerCase()) ||
  //       (item.Customer_x0020_Contact || "").toLowerCase().includes(filters.search.toLowerCase()) 
  //       &&
  //     (!filters.customer || item.Customer === filters.customer) &&
  //     (!filters.status || item.Status === filters.status) &&
  //     (!filters.financeStatus || item.FinanceStatus === filters.financeStatus) && // NEW filter
  //     (!filters.requestedDate || (item.RequestedDate && new Date(item.RequestedDate).toLocaleDateString() === filters.requestedDate.toLocaleDateString()))
  //   );
  // });

  const filteredItems = items.filter(item => {
    const searchText = filters.search?.toLowerCase() || "";

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
      (!filters.requestedDate || (item.RequestedDate && new Date(item.RequestedDate).toLocaleDateString() === filters.requestedDate.toLocaleDateString()))
    );
  });

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
    const submittedStates = ["Submitted"];

    const statusToUse = submittedStates.includes(normalizedStatus)
      ? defaultStatusForSubmitted
      : normalizedStatus;

    setEditFields({
      Status: statusToUse,
      FinanceComments: item.FinanceComments ?? "",
      InvoiceNumber: item.InvoiceNumber || "",
      FinanceStatus: "Submitted",
    });

    setOriginalStatus(item.Status ?? null);
    setAttachments([]);
    setIsPanelOpen(true);
    loadPmAttachments(item);
  }

  async function handleClarification() {
    if (!selectedItem) return;

    if (!editFields.FinanceComments || editFields.FinanceComments.trim() === "") {
      alert("Please enter finance comments before submitting clarification.");
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
  // const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
  //   if (e.target.files) {
  //     setAttachments(Array.from(e.target.files));
  //   }
  // };

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
      let updatedFinanceStatus = editFields.FinanceStatus || selectedItem.FinanceStatus || "";
      if ((editFields.Status || selectedItem.Status) === "Payment Received") {
        updatedFinanceStatus = "Paid";
      } else {
        updatedFinanceStatus = "Pending";
      }

      // Include updated FinanceCommentsHistory JSON string in update payload
      const updatePayload = {
        ...editFields,
        FinanceCommentsHistory: JSON.stringify(historyArr),
        FinanceComments: editFields.FinanceComments || "",
        FinanceStatus: updatedFinanceStatus,
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
  // async function getCurrentUserRole(context: any, poId: any): Promise<string> {
  //   try {
  //     const currentUserEmail = context.pageContext.user.email.toLowerCase();
  //     const projects = await projectsp
  //       .web.lists.getByTitle("Projects")
  //       .items
  //       .filter(`Title eq '${poId?.ProjectName?.replace(/'/g, "''")}'`)
  //       .select("PM/EMail", "DM/EMail", "DH/EMail")
  //       .expand("PM", "DM", "DH")
  //       .top(100)();

  //     const project = projects[0];
  //     if (!project) return "Unknown";

  //     if (project.PM?.EMail.toLowerCase() === currentUserEmail) return "PM";
  //     if (project.DM?.EMail.toLowerCase() === currentUserEmail) return "DM";
  //     if (project.DH?.EMail.toLowerCase() === currentUserEmail) return "DH";

  //     return "Unknown";
  //   } catch (error) {
  //     console.error("Error determining role", error);
  //     return "Unknown";
  //   }
  // }


  return (
    <section style={{ background: "#fff", borderRadius: 8, padding: 16 }}>
      <h2 style={{ fontWeight: 600, marginBottom: 16 }}>Update Invoice Request</h2>

      {/* FILTERS BAR */}
      {/* <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 18 }} styles={{ root: { marginBottom: 20 } }}>
        <div>
          <Label>Search</Label>
          <TextField
            placeholder="Search"
            value={filters.search}
            onChange={(_, v) => setFilters(f => ({ ...f, search: v || "" }))}
            styles={{ root: { width: 150 } }}
          />
        </div>
        <div>
          <Label>Current Status</Label>
          <Dropdown
            placeholder="Current Status"
            options={financeStatusOptions}
            selectedKey={filters.financeStatus}
            onChange={(_, option) => setFilters(f => ({ ...f, financeStatus: (option?.key ?? "").toString() }))}
            styles={{ root: { width: 150 } }}
          />
        </div>
        <div>
          <Label>Invoice Status</Label>
          <Dropdown
            placeholder="Invoice Status"
            options={statusOptions}
            selectedKey={filters.status}
            onChange={(_, option) => setFilters(f => ({ ...f, status: option?.key as string || "" }))}
            styles={{ root: { width: 150 } }}
          />
        </div>

        <div>
          <PrimaryButton
            text="Clear"
            onClick={clearFilters}
            disabled={
              !filters.search &&
              !filters.requestedDate &&
              !filters.customer &&
              !filters.status &&
              !filters.financeStatus
            }
            style={{ alignSelf: "center", marginLeft: 20 }}
          />
        </div>

      </Stack> */}
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" styles={{ root: { marginBottom: 20 } }}>
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
            options={financeStatusOptions}
            selectedKey={filters.financeStatus}
            onChange={(_, option) => setFilters(f => ({ ...f, financeStatus: (option?.key ?? "").toString() }))}
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
              !filters.financeStatus
            }
          />
        </Stack.Item>
      </Stack>


      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {loading && <Spinner label="Loading Invoice Requests..." />}
      {!loading && (
        <>
          <DetailsList
            items={filteredItems}
            columns={columns}
            selection={selection}
            selectionMode={SelectionMode.single}
            setKey="financeViewList"
            styles={{ root: { backgroundColor: "#fff" } }}
          />
        </>
      )}

      <Panel
        isOpen={isPanelOpen}
        onDismiss={handlePanelDismiss}
        headerText="Update Invoice Request"
        type={PanelType.large}
        customWidth="950px"
        isBlocking={false}
        isFooterAtBottom={false}
      >
        {isPanelOpen && selectedItem && (
          <Stack
            horizontal
            styles={{ root: { height: 'calc(100vh - 150px)', overflow: 'hidden' } }}
            tokens={{ childrenGap: 20 }}
          >
            {!isViewerOpen && (
              // ---- Form and Attachments Section ----
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
                    <TextField label="Invoice Amount" value={selectedItem?.InvoiceAmount?.toString() || ''} disabled />
                    <TextField label="Customer Contact" value={selectedItem?.Customer_x0020_Contact || ''} disabled />
                    <TextField label="PO Item Value" value={selectedItem?.POItem_x0020_Value || ''} disabled />
                    <Dropdown
                      label="Invoice Status"
                      options={statusOptions}
                      selectedKey={editFields.Status || selectedItem.Status || ''}
                      onChange={(_, option) => handleFieldChange('Status', option?.key)}
                    />
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
                      onChange={e => {
                        if (e.target.files) setAttachments(Array.from(e.target.files));
                      }}
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
                  <PrimaryButton
                    text="Remove All Attachments"
                    onClick={e => {
                      e.stopPropagation();
                      setAttachments([]);
                    }}
                    disabled={attachments.length === 0}
                    styles={{ root: { marginTop: 8, minWidth: 140 } }}
                  />
                </Stack>
                {/* Buttons */}
                <div style={{ height: 28 }} />
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
