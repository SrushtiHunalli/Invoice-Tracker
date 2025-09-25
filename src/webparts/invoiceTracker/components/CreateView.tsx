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
  DetailsListLayoutMode,
  IRenderFunction,
  Sticky,
  StickyPositionType,
  ContextualMenu,
  ContextualMenuItemType,
  Separator,
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
// import { columnSort, SortType } from "./SortingUtils"
import styles from "./CreateView.module.scss"
// import { SPFx } from "@pnp/sp/presets/all";
interface CreateViewProps {
  sp: SPFI;
  context: any;
  projectsp: SPFI;
}
interface InvoiceRequestCardProps {
  invoice: InvoiceRequest;
}
type PurchaseOrderItem = {
  Id: number;
  POID: string;
  ProjectName?: string;
  POAmount?: string;
  Currency?: string;
  POComments?: string;
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
};

const CreateView: React.FC<CreateViewProps> = ({ sp, projectsp, context }) => {
  // const [mergedItems, ] = useState<PurchaseOrderItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [filters, setFilters] = useState({ search: "", customer: "" });
  // const [customerOptions, ] = useState<IDropdownOption[]>([]);
  const [selectedItem, setSelectedItem] = useState<PurchaseOrderItem | null>(null);
  const [error, setError] = useState<string>("");
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [childPOItems, setChildPOItems] = useState<ChildPOItem[]>([]);
  const [fetchingChildPOs, setFetchingChildPOs] = useState(false);
  const [invoiceRequests, setInvoiceRequests] = useState<InvoiceRequest[]>([]);
  const [invoiceRequestsForPercent, setInvoiceRequestsForPercent] = useState<InvoiceRequest[]>([]);
  const [fetchingInvoices, setFetchingInvoices] = useState(false);
  const [activePOIDFilter, setActivePOIDFilter] = useState<string | null>(null);
  const [childPOSelection] = useState(new Selection());
  const [invoiceAmountError, setInvoiceAmountError] = useState<string | undefined>(undefined);
  const [isDragActive, setIsDragActive] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [previewFileIdx, setPreviewFileIdx] = useState<number | null>(null);
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
  const [columnFilterMenu, setColumnFilterMenu] = React.useState<{ visible: boolean; target: HTMLElement | null; columnKey: string | null }>({ visible: false, target: null, columnKey: null });

  const [, setSortedColumnKey] = React.useState<string | null>(null);
  const [, setIsSortedDescending] = React.useState<boolean>(false);
  const [isInvoiceRequestViewPanelOpen, setIsInvoiceRequestViewPanelOpen] = useState(false);
  const [selectedInvoiceRequest, setSelectedInvoiceRequest] = useState<InvoiceRequest | null>(null);
  const onInvoiceRequestClicked = (item: InvoiceRequest) => {
    setSelectedInvoiceRequest(item);
    setIsInvoiceRequestViewPanelOpen(true);
  };

  // const [selectedMainPO, setSelectedMainPO] = useState<PurchaseOrderItem | null>(null);
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
  });
  const onColumnHeaderClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn) => {
    if (column && ev) {
      setColumnFilterMenu({ visible: true, target: ev.currentTarget, columnKey: column.key });
    }
  };
  const columns: IColumn[] = [
    { key: "POID", name: "Purchase Order", fieldName: "POID", minWidth: 100, maxWidth: 150, isResizable: true, onColumnClick: onColumnHeaderClick },
    { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 150, maxWidth: 220, isResizable: true, onColumnClick: onColumnHeaderClick },
    // { key: "Customer", name: "Customer", fieldName: "Customer", minWidth: 120, maxWidth: 160, isResizable: true },
    {
      key: "POAmount", name: "PO Amount", fieldName: "POAmount", minWidth: 120, maxWidth: 160, isResizable: true, onColumnClick: onColumnHeaderClick, onRender: (item: PurchaseOrderItem) => {
        return `${item.POAmount} ${item.Currency ?? ''}`.trim();
      }
    },
    // { key: "Currency", name: "Currency", fieldName: "Currency", minWidth: 70, maxWidth: 90, isResizable: true, onColumnClick: onColumnHeaderClick },
    { key: "POComments", name: "PO Comments", fieldName: "POComments", minWidth: 70, maxWidth: 90, isResizable: true, onColumnClick: onColumnHeaderClick },
    {
      key: 'InvoicedPercent',
      name: 'Invoiced %',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onColumnClick: onColumnHeaderClick,
      onRender: item => {
        return `${calculateInvoicedPercent(item.POID, item.POAmount).toFixed(2)}%`;
      }
    }
  ];
  const invoiceColumnsView: IColumn[] = [
    { key: "POItemTitle", name: "PO Item Title", fieldName: "POItemTitle", minWidth: 130, maxWidth: 180, isResizable: true },
    { key: "POItemValue", name: `PO Item Value${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`, fieldName: "POItemValue", minWidth: 120, maxWidth: 140, isResizable: true },
    { key: "Amount", name: `Invoiced Amount${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`, fieldName: "Amount", minWidth: 120, maxWidth: 160, isResizable: true },
    { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 140, maxWidth: 170, isResizable: true },
    {
      key: "PMCommentsHistory",
      name: "PM Comments History",
      fieldName: "PMCommentsHistory",
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        if (!item.PMCommentsHistory) return "No PM Comments";
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
      name: "Finance Comments History",
      fieldName: "FinanceCommentsHistory",
      minWidth: 200,
      maxWidth: 300,
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
      maxWidth: 160,
      isResizable: true,
      onRender: (item: InvoiceRequest) => new Date(item.Created).toLocaleString()
    },
    {
      key: "CreatedBy",
      name: "Created By",
      fieldName: "CreatedBy",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: "Modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      onRender: (item: InvoiceRequest) => new Date(item.Modified).toLocaleString()
    },
    {
      key: "ModifiedBy",
      name: "Modified By",
      fieldName: "ModifiedBy",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    }
  ];
  const childPOColumns: IColumn[] = [
    {
      key: "POID",
      name: "PO Item Title",
      fieldName: "POID",
      minWidth: 150,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: ChildPOItem) => (
        <span style={{ color: "#0078d4", cursor: "pointer", fontWeight: 500 }}>{item.POID}</span>
      ),
    },
    {
      key: "POItemValue",
      name: `PO Item Value${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`,
      fieldName: "POItemValue",
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: ChildPOItem) => (
        <span>{item.POAmount}</span>
      ),
    },

    {
      key: "POAmount", name: `Remaining Item Value${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`, fieldName: "POAmount", minWidth: 120, maxWidth: 150, isResizable: true, onRender: (item: ChildPOItem) => {
        const remaining = getRemainingPOAmount(item, invoiceRequests);
        return <span>{remaining}</span>;
      },
    },
    {
      key: "action",
      name: "",
      fieldName: "action",
      minWidth: 34,
      maxWidth: 34,
      isResizable: false,
      onRender: (item: ChildPOItem) => {
        const remaining = getRemainingPOAmount(item, invoiceRequests);
        return remaining > 0 ? (
          <IconButton
            iconProps={{ iconName: "Add" }}
            ariaLabel="Create Invoice Request"
            onClick={() => handleOpenInvoicePanel(item)}
            styles={{ root: { marginLeft: 8 } }}
          />
        ) : null;
      },
    },

  ];
  const [invoicePanelLoading, setInvoicePanelLoading] = useState(false);

  // const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
  //   e.preventDefault();
  //   setIsDragActive(false);
  //   const file = e.dataTransfer.files[0];
  //   if (file && (file.type === "application/pdf" || file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || file.type === "application/vnd.ms-excel")) {
  //     setUploadedFiles(file);
  //     setInvoiceFormState(prev => ({ ...prev, Attachment: file }));
  //   }
  // };
  const menuItems = [
    { key: 'asc', text: 'Sort Asc to Desc', iconProps: { iconName: 'SortUp' }, onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc') },
    { key: 'desc', text: 'Sort Desc to Asc', iconProps: { iconName: 'SortDown' }, onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc') },
    { key: 'divider', itemType: ContextualMenuItemType.Divider },
    // { key: 'filter', text: 'Filter...', iconProps: { iconName: 'Filter' }, onClick: () => openFilterPanelFromMenu() },
    // { key: 'clear', text: 'Clear Filter', iconProps: { iconName: 'ClearFilter' }, onClick: () => clearColumnFilter(columnFilterMenu.columnKey!) },
  ];
  const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
    const sortedItems = [...filteredMainPOs].sort((a, b) => {
      let aVal = (a as any)[columnKey];
      let bVal = (b as any)[columnKey];

      if (aVal == null) return 1;
      if (bVal == null) return -1;

      // Handle numbers and strings
      if (typeof aVal === 'number' && typeof bVal === 'number') {
        return direction === 'asc' ? aVal - bVal : bVal - aVal;
      }
      return direction === 'asc'
        ? aVal.toString().localeCompare(bVal.toString())
        : bVal.toString().localeCompare(aVal.toString());
    });
    setMainPOs(sortedItems)

    // Close menu
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
    // Update sort state if needed
    setSortedColumnKey(columnKey);
    setIsSortedDescending(direction === 'desc');
  };

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

  useEffect(() => {
    (async () => {
      setLoading(true);
      setError("");
      try {
        const items = await sp.web.lists.getByTitle("InvoicePO")
          .items
          .select("ID", "POID", "ParentPOID", "POAmount", "LineItemsJSON", "ProjectName", "Currency", "POComments")();

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
        const email = context.pageContext.user.email.toLowerCase();
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

  const filteredMainPOs = mainPOs.filter(po => {
    if (!po.ProjectName) return false;

    const project = allProjects.find(p => p.Title === po.ProjectName);
    if (!project) return false;

    if (isAdminUser) return true;

    const userEmail = currentUserEmail.toLowerCase();

    const isUserPM = project.PM?.EMail?.toLowerCase() === userEmail;
    const isUserDM = project.DM?.EMail?.toLowerCase() === userEmail;
    const isUserDH = project.DH?.EMail?.toLowerCase() === userEmail;

    const isInPMGroup = userGroups.includes("pm");
    const isInDMGroup = userGroups.includes("dm");
    const isInDHGroup = userGroups.includes("dh");

    // User must be in the group and match the project role
    if ((isInPMGroup && isUserPM) ||
      (isInDMGroup && isUserDM) ||
      (isInDHGroup && isUserDH)) {
      return true;
    }

    return false;
  });

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
  const handleInvoiceAmountChange = (value: string) => {
    handleInvoiceFormChange("InvoiceAmount", value);

    const enteredAmount = parseFloat(value);
    if (!value) {
      setInvoiceAmountError("Invoiced Amount is required.");
    } else if (isNaN(enteredAmount) || enteredAmount <= 0) {
      setInvoiceAmountError("Please enter a valid positive number.");
    } else {
      // Determine if this is a single PO or child PO invoice form
      // For single PO, POItemTitle will be empty, so use POAmount directly
      const isSinglePO = !invoiceFormState.POItemTitle;

      const remainingAmount = getRemainingPOAmount(
        {
          POID: isSinglePO ? invoiceFormState.POID : (invoiceFormState.POItemTitle || ""),
          POAmount: isSinglePO ? (invoiceFormState.POAmount || "0") : (invoiceFormState.POItemValue || "0"),
          Id: 0,
          ParentPOIndex: 0,
          POIndex: 0,
        },
        invoiceRequests
      );
      if (enteredAmount > remainingAmount) {
        setInvoiceAmountError(`Invoiced Amount cannot exceed remaining amount: ${remainingAmount}`);
      } else {
        setInvoiceAmountError(undefined);
      }
    }
  };
  const handleOpenPanel = async () => {
    if (!selectedItem) return;
    setInvoiceCurrency(selectedItem.Currency || "");
    setFetchingChildPOs(true);
    setFetchingInvoices(true);
    setChildPOItems([]);
    setInvoiceRequests([]);
    setActivePOIDFilter(null);
    setIsPanelOpen(false);
    setIsInvoicePanelOpen(false);

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
  const fieldStyle = {
    root: { marginBottom: 8, display: 'flex' },
    label: { fontWeight: 600, width: 180, color: '#626262', fontSize: 15 },
    value: { fontWeight: 400, color: '#303030', fontSize: 15 }
  };

  const historyStyle = {
    root: { background: '#f6f8fa', padding: '7px 14px', borderRadius: 6, marginBottom: 8, fontSize: 14 }
  };

  // Helper to render text or fallback
  const renderValue = (value: any) => value ? value : <span style={{ color: '#999' }}>—</span>;

 const InvoiceRequestCard: React.FC<InvoiceRequestCardProps> = ({ invoice }) => (
    <Stack tokens={{ childrenGap: 18 }} styles={{ root: { background: '#fff', borderRadius: 10, boxShadow: '0 2px 16px #edf1f3', padding: 28, margin: '0 auto', width: '100%', maxWidth: 650 } }}>
      <Text variant="xLarge" styles={{ root: { marginBottom: 12, fontWeight: 600 } }}>Invoice Request Details</Text>
      <Separator />

      <div style={fieldStyle.root}><div style={fieldStyle.label}>PO Item Title</div><div style={fieldStyle.value}>{renderValue(invoice.POItemTitle)}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>PO Item Value</div><div style={fieldStyle.value}>{renderValue(invoice.POItemValue)}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>Invoiced Amount</div><div style={fieldStyle.value}>{renderValue(invoice.Amount)}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>Invoice Status</div><div style={fieldStyle.value}>{renderValue(invoice.Status)}</div></div>

      <Separator />

      <div style={fieldStyle.label}>PM Comments History</div>
      <div style={historyStyle.root}>{renderValue(formatCommentHistory(invoice.PMCommentsHistory))}</div>

      <div style={fieldStyle.label}>Finance Comments History</div>
      <div style={historyStyle.root}>{renderValue(formatCommentHistory(invoice.FinanceCommentsHistory))}</div>

      <Separator />

      <div style={fieldStyle.root}><div style={fieldStyle.label}>Created</div><div style={fieldStyle.value}>{renderValue(new Date(invoice.Created).toLocaleString())}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>Created By</div><div style={fieldStyle.value}>{renderValue(invoice.CreatedBy)}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>Modified</div><div style={fieldStyle.value}>{renderValue(new Date(invoice.Modified).toLocaleString())}</div></div>
      <div style={fieldStyle.root}><div style={fieldStyle.label}>Modified By</div><div style={fieldStyle.value}>{renderValue(invoice.ModifiedBy)}</div></div>
    </Stack>
  );
  // useEffect(() => {
  //   const fetchGroups = async () => {
  //     try {
  //       const groups = await sp.web.currentUser.groups();
  //       setUserGroups(groups.map((g: any) => g.Title.toLowerCase()));
  //     } catch (error) {
  //       setUserGroups([]);
  //     }
  //   };
  //   fetchGroups();
  // }, [sp]);

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
    const relevantPoIds = mainPOs.map(po => po.POID); // Adjust source of POIDs as needed
    loadInvoiceRequestsForPercent(relevantPoIds);
  }, [mainPOs]);

  const handleOpenInvoicePanelSinglePO = async (poItem: PurchaseOrderItem, poAmount: string) => {
    setInvoicePanelPO(null);
    setIsInvoicePanelOpen(true);
    // const projectName = await getProjectNameByPOID(context, poItem.Id, poItem);
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
    });
  };
  const handlePanelDismiss = () => {
    setIsPanelOpen(false);
    setChildPOItems([]);
    setInvoiceRequests([]);
    setActivePOIDFilter(null);
    childPOSelection.setAllSelected(false);
    window.history.replaceState(null, '', window.location.pathname);
  };
  const handleChildPORowClick = (item?: ChildPOItem) => {
    if (item) {
      setActivePOIDFilter(item.POID);
      childPOSelection.setKeySelected(item.Id.toString(), true, false);
    }
  };
  const showInvoices = activePOIDFilter
    ? invoiceRequests.filter((ir) => ir.POItemTitle === activePOIDFilter)
    : invoiceRequests;

  const handleInvoiceFormChange = (field: keyof InvoiceFormState, value: any) => {
    setInvoiceFormState((prev) => ({
      ...prev,
      [field]: value,
    }));
  };
  function getRemainingPOAmount(childPO: ChildPOItem, invoiceRequests: InvoiceRequest[]): number {
    // const childInvoices = invoiceRequests.filter(inv => inv.POItemTitle === childPO.POID);
    const childInvoices = invoiceRequests.filter(inv =>
      inv.POItemTitle === childPO.POID &&
      inv.Status?.toLowerCase() !== "cancelled"
    );
    const usedAmount = childInvoices.reduce((sum, inv) => sum + (inv.Amount || 0), 0);
    const originalAmount = parseFloat(childPO.POAmount) || 0;
    return originalAmount - usedAmount;
  }

  function calculateInvoicedPercent(rowPOID: string, mainPOAmount: number): number {
    if (!invoiceRequestsForPercent || invoiceRequestsForPercent.length === 0 || !mainPOAmount) {
      return 0;
    }
    const matchedInvoices = invoiceRequestsForPercent.filter(
      inv => inv.PurchaseOrderPO.trim() === rowPOID.trim()
    );
    const totalInvoicedAmount = matchedInvoices.reduce((sum, inv) => sum + (inv.Amount || 0), 0);
    return mainPOAmount > 0 ? (totalInvoicedAmount / mainPOAmount) * 100 : 0;
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

    setInvoiceFormState({
      POID: parentPOID,
      PurchaseOrder: parentPOID,
      ProjectName: projectName,
      POItemTitle: item.POID,
      POAmount: item.POAmount,
      POItemValue: item.POAmount,
      InvoiceAmount: String(getRemainingPOAmount(item, invoiceRequests)),
      CustomerContact: "",
      Comments: "",
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
  };

  const handleDropMulti = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragActive(false);
    const files = Array.from(e.dataTransfer.files);
    setUploadedFiles(prev => [...prev, ...files]);
  };

  // Remove specific file
  const removeAttachment = (idx: number) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== idx));
  };

  function decodeHtmlEntities(str: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = str;
    return txt.value;
  }

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

  async function fetchAllPORecords(): Promise<any[]> {
    // const remote = spfi(PROJECTS_SITE_URL).using(SPFx(context));
    const allItems: any[] = [];

    try {
      // Fetch all items, selecting only relevant InvoicePO fields
      const pagedItems = sp.web.lists.getByTitle("InvoicePO")
        .items
        .select("ID", "POID", "ParentPOID", "POAmount", "LineItemsJSON", "ProjectName")
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


  async function fetchInvoiceRequests(sp: SPFI, poids: string[]): Promise<InvoiceRequest[]> {
    if (poids.length === 0) return [];

    const filter = `(${poids.map(po => `PurchaseOrder eq '${po}'`).join(" or ")})`;


    console.log(filter);
    try {
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
          "Editor/Title"
        )
        .expand("Author", "Editor")();

      // console.log(items)
      return items.map(item => ({
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
      }));
    } catch {
      return [];
    }
  }

  async function getCurrentUserRole(context: any, poId: any): Promise<string> {
    try {
      // const sp = spfi(PROJECTS_SITE_URL).using(SPFx(context));
      const currentUserEmail = context.pageContext.user.email.toLowerCase();

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
        .top(100)();

      // const projectNameFromInvoice = poId?.ProjectName;
      // const matchedProject = projects.find((p: any) => {
      //   const projectTitle = (p.Title ?? "").toString().trim().toLowerCase();
      //   const invoiceProjectName = (projectNameFromInvoice ?? "").toString().trim().toLowerCase();

      //   return projectTitle === invoiceProjectName;
      // });

      const matchedProject = projects[0];
      if (!matchedProject) return "Unknown Role";

      if (matchedProject.PM?.EMail.toLowerCase() === currentUserEmail) return "PM";
      if (matchedProject.DM?.EMail.toLowerCase() === currentUserEmail) return "DM";
      if (matchedProject.DH?.EMail.toLowerCase() === currentUserEmail) return "DH";

      return "Unknown Role";
    } catch (error) {
      console.error("Error determining user role:", error);
      return "Unknown Role";
    }
  }

  async function getProjectNameByPOID(context: any, poId: number, poItem: any): Promise<string> {
    try {
      console.log("Fetching project name for POID:", poId, poItem);
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
  async function loadInvoiceRequestsForPercent(poIds: string[]) {
    if (poIds.length === 0) {
      setInvoiceRequestsForPercent([]);
      return;
    }

    try {
      const fetchedInvoices = await fetchInvoiceRequests(sp, poIds); // your existing fetch func
      setInvoiceRequestsForPercent(fetchedInvoices);
    } catch (error) {
      console.error("Failed to load invoice requests for percent calculation", error);
      setInvoiceRequestsForPercent([]);
    }
  }

  const handleInvoiceFormSubmit = async () => {
    let addedItemId: number | null = null;
    try {
      if (invoiceAmountError || !invoiceFormState.InvoiceAmount) {
        alert(invoiceAmountError || "Invoiced Amount is required.");
        return;
      }
      const userRole = await getCurrentUserRole(context, selectedItem);

      const financeStatusValue = "Not Generated";

      const newCommentEntry = {
        Date: new Date().toISOString(),
        Title: "Comment",
        User: context.pageContext.user.displayName,
        Role: userRole,
        Data: invoiceFormState.Comments || ""
      };

      const added = await sp.web.lists.getByTitle("Invoice Requests").items.add({
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
        CurrentStatus: `Request Submitted by ${userRole}`
      });

      addedItemId = added.Id;

      const currentItem = await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).select("PMCommentsHistory")();
      let history = [];

      try {
        history = currentItem.PMCommentsHistory ? JSON.parse(currentItem.PMCommentsHistory) : [];
      } catch { history = []; }

      history.push(newCommentEntry);

      await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).update({
        PMCommentsHistory: JSON.stringify(history)
      });

      if (invoicePanelPO === null && invoiceFormState.POItemValue) {
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).update({
          POAmount: Number(invoiceFormState.POItemValue)
        });
      }
      if (invoiceFormState.Attachment) {
        const file = invoiceFormState.Attachment;
        const fileNameWithSuffix = `${file.name.replace(/\.[^/.]+$/, "")}_PM${file.name.match(/\.[^/.]+$/)?.[0] || ""}`;
        const fileContent = await file.arrayBuffer();
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId)
          .attachmentFiles.add(fileNameWithSuffix, fileContent);
      }

      // const siteUrl = context.pageContext.web.absoluteUrl;
      // const listName = "Invoice Requests";

      // const itemUrl = `${siteUrl}/Lists/${listName}/DispForm.aspx?ID=${addedItemId}`;
      const creatorEmail = context.pageContext.user.email;

      const siteUrl = context.pageContext.web.absoluteUrl;
      const pageName = context.pageContext.site.serverRequestPath.split('/').pop() || 'InvoiceTracker.aspx';
      const appPageUrl = `${siteUrl}/SitePages/${pageName}`;

      const itemLink = `${appPageUrl}#myrequests?selectedInvoice=${addedItemId}`;


      const sendNotificationEmail = async () => {
        try {
          await sp.utility.sendEmail({
            To: [creatorEmail],
            Subject: `New Invoice Request: ${invoiceFormState.InvoiceAmount} for ${invoiceFormState.PurchaseOrder}`,
            Body: `
        A new invoice request has been created.<br/><br/>
        <b>PO ID:</b> ${invoiceFormState.POID}<br/>
        <b>Project Name:</b> ${invoiceFormState.ProjectName}<br/>
        <b>PO Item Title:</b> ${invoiceFormState.POItemTitle}<br/>
        <b>Invoiced Amount:</b> ${invoiceFormState.InvoiceAmount}<br/>
        <b>Comments:</b> ${invoiceFormState.Comments}<br/><br/>
        <a href="${itemLink}">Click here to view the invoice request.</a>
      `,
          });
          setDialogType("success");
          setDialogMessage("Invoice request submitted successfully!");
          setDialogVisible(true);

          // Refresh invoiceRequests data after update
          // const lookupPOIDs = [invoiceFormState.POID, ...childPOItems.map(c => c.POID)];
          const allPOIDs = mainPOs.map(po => po.POID);
          const updatedInvoices = await fetchInvoiceRequests(sp, allPOIDs);
          setInvoiceRequests(updatedInvoices);
          setInvoiceRequestsForPercent(updatedInvoices);
          setIsInvoicePanelOpen(false);
          setInvoicePanelPO(null);


        } catch (error) {
          if (addedItemId !== null) {
            await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).delete();
          }
          setDialogType("error");
          setDialogMessage("Error submitting invoice request: " + (error as any)?.message);
          setDialogVisible(true);
        }
      };

      await sendNotificationEmail();


      // const lookupPOIDs = [invoiceFormState.POID, ...childPOItems.map(c => c.POID)];
      const allPOIDs = mainPOs.map(po => po.POID);
      const updatedInvoices = await fetchInvoiceRequests(sp, allPOIDs);
      setInvoiceRequests(updatedInvoices);
      setInvoiceRequestsForPercent(updatedInvoices);

      setIsInvoicePanelOpen(false);
      setInvoicePanelPO(null);

    } catch (error) {
      if (addedItemId !== null) {
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).delete();
      }
      alert("Error submitting invoice request: " + (error as any)?.message);
    }
  };
  // Inside the CreateView component function, before return:

  const adjustedChildColumns = computeColumnWidths(childPOItems, childPOColumns);

  // For invoice requests, you have invoiceRequests as state; apply filter if needed:
  const filteredInvoiceRequests = activePOIDFilter
    ? invoiceRequests.filter((ir) => ir.POItemTitle === activePOIDFilter)
    : invoiceRequests;

  const adjustedInvoiceColumns = computeColumnWidths(filteredInvoiceRequests, invoiceColumnsView);

  return (
    <section style={{ background: "#fff", borderRadius: 8, padding: 16 }}>
      <div style={{ flexGrow: 1, overflowY: 'auto' }}>
        <h2 style={{ marginBottom: 20 }}>Create Invoice Request</h2>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 16 }} styles={{ root: { flexWrap: "nowrap", overflowX: "auto", paddingBottom: 8 } }}>
          <SearchBox
            placeholder="Search"
            value={filters.search}
            onChange={(ev, newVal) => setFilters((f) => ({ ...f, search: newVal || "" }))}
            styles={{ root: { width: 250, minWidth: 250 } }}
          />
          <PrimaryButton text="Create Invoice Request" disabled={!selectedItem} onClick={handleOpenPanel} />
        </Stack>
        {loading && <Spinner label="Loading data..." />}
        {error && <div style={{ color: "red" }}>{error}</div>}
        {!loading && !error && (
          <div className={`ms-Grid-row ${styles.detailsListContainer}`}>
            <div style={{ height: 300, position: 'relative' }}>
              <ScrollablePane>
                <div
                  className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.detailsList_Scrollablepane_Container}`}
                >
                  <DetailsList
                    items={filteredMainPOs}
                    columns={columns}
                    selection={selection}
                    selectionMode={SelectionMode.single}
                    setKey="mainPOsList"
                    isHeaderVisible={true}
                    layoutMode={DetailsListLayoutMode.justified}
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
        )}
        <Panel
          isOpen={isPanelOpen}
          onDismiss={handlePanelDismiss}
          // headerText="Purchase Order"
          closeButtonAriaLabel="Close"
          type={PanelType.custom}
          isLightDismiss={false}
          isBlocking={false}
          isFooterAtBottom={true}
          styles={{
            main: {
              maxWidth: 1000,
              width: '100%',
              margin: 'auto',
              borderRadius: 8,
            },
            scrollableContent: {
              padding: 16,
            },
          }}
        >
          <Stack tokens={{ childrenGap: 18 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
            {/* <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
            <TextField
              label={"Purchase Order:"}
              value={selectedItem?.POID || ""}
              readOnly
              disabled
              styles={{ root: { maxWidth: 280, marginBottom: 0 } }}
            />
            <TextField
              label={`PO Amount${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`}
              value={selectedItem?.POAmount || ""}
              readOnly
              disabled
              styles={{ root: { maxWidth: 280, marginTop: 10, marginBottom: 0 } }}
            />
            </div> */}
            <div style={{ display: "flex", flexDirection: "row", gap: 24, alignItems: "flex-start", marginTop: 0, marginBottom: 0 }}>
              <TextField
                label="Purchase Order"
                value={selectedItem?.POID}
                readOnly
                disabled
                styles={{ root: { maxWidth: 220, marginTop: 0, marginBottom: 0, fontSize: 15, fontWeight: 600 } }}
              />
              <TextField
                label={`PO Amount${selectedItem?.Currency ? ` (${selectedItem?.Currency})` : ''}`}
                value={selectedItem?.POAmount}
                readOnly
                disabled
                styles={{ root: { maxWidth: 220, marginTop: 0, marginBottom: 0, fontSize: 15, fontWeight: 600 } }}
              />
            </div>

            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 7, color: "#626262" }}>PO Items:</div>
            <div>
              {fetchingChildPOs ? (
                <Spinner label="Loading child POs..." />
              ) : childPOItems.length > 0 ? (
                <DetailsList
                  items={childPOItems}
                  columns={adjustedChildColumns}
                  selection={childPOSelection}
                  selectionMode={SelectionMode.single}
                  setKey="childPOs"
                  onActiveItemChanged={handleChildPORowClick}
                  styles={{
                    root: {
                      background: "#fff",
                      border: "1px solid #eee",
                      borderRadius: 6,
                      overflowX: "hidden",
                      width: '100%',
                      minWidth: 0,
                    },
                  }}
                />
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
                    onClick={() => setActivePOIDFilter(null)}
                    styles={{ root: { marginLeft: 24 } }}
                    disabled={!activePOIDFilter}
                  />
                </div>
              </div>
            </div>
            <div>
              {fetchingInvoices ? (
                <Spinner label="Loading invoice requests..." />
              ) : showInvoices.length > 0 ? (
                <DetailsList
                  items={showInvoices}
                  columns={adjustedInvoiceColumns}
                  selectionMode={SelectionMode.single}
                  onActiveItemChanged={onInvoiceRequestClicked}
                  setKey="invoiceRequests"
                  styles={{ root: { background: "#fff", border: "1px solid #eee", borderRadius: 6 } }}
                />
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
              left: "unset",
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
              <Stack styles={{ root: { flex: 1, minWidth: "340px", maxWidth: "48%" } }} tokens={{ childrenGap: 12 }}>
                <TextField label="PO ID" value={invoiceFormState.POID} readOnly disabled />
                <TextField label="Project Name" value={invoiceFormState.ProjectName} readOnly disabled />
                {invoicePanelPO && (
                  <>
                    <TextField label="PO Item Title" value={invoiceFormState.POItemTitle} readOnly disabled />
                    <TextField label={`PO Item Value${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`} value={invoiceFormState.POItemValue} readOnly disabled />
                    <TextField
                      label={`Amount remaining${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`}
                      value={String(getRemainingPOAmount(
                        { POID: invoiceFormState.POItemTitle || "", POAmount: invoiceFormState.POItemValue || "0", Id: 0, ParentPOIndex: 0, POIndex: 0 },
                        invoiceRequests
                      ))}
                      readOnly disabled
                    />
                  </>
                )}
                <TextField
                  label={`Invoiced Amount${invoiceCurrency ? ` (${invoiceCurrency})` : ""}`}
                  value={invoiceFormState.InvoiceAmount}
                  onChange={(_, val) => handleInvoiceAmountChange(val || "")}
                  type="number"
                  required
                  errorMessage={invoiceAmountError}
                />
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
                <PrimaryButton text="Submit" onClick={handleInvoiceFormSubmit} styles={{ root: { marginTop: 12, minWidth: 110 } }} />
              </Stack>

              {/* RIGHT HALF: Attachments & Preview */}
              <Stack styles={{ root: { flex: 1, minWidth: "340px", maxWidth: "48%" } }}>
                {previewFileIdx !== null && uploadedFiles[previewFileIdx] ? (
                  <Stack>
                    <PrimaryButton
                      text="Close Preview"
                      onClick={() => setPreviewFileIdx(null)}
                      styles={{ root: { marginBottom: 10 } }}
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
                      title={`Preview-${uploadedFiles[previewFileIdx].name}`}
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
                            <span style={{ flex: 1, fontWeight: 520, fontSize: 15, overflow: "hidden", textOverflow: "ellipsis" }}>{file.name}</span>
                            <IconButton
                              iconProps={{ iconName: "Remove" }}
                              title="Remove"
                              ariaLabel={`Remove ${file.name}`}
                              onClick={() => removeAttachment(idx)}
                              styles={{ root: { height: 28, minWidth: 28, color: "#ba0808" } }}
                            />
                            <PrimaryButton
                              text="Preview"
                              onClick={() => setPreviewFileIdx(idx)}
                              styles={{ root: { marginLeft: 10, minWidth: 60, height: 28 } }}
                            />
                          </Stack>
                        ))
                      )}
                    </div>
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
                <DetailsList
                  items={invoiceRequests.filter((inv) => inv.POItemTitle === invoiceFormState.POItemTitle)}
                  columns={invoiceColumnsView}
                  selectionMode={SelectionMode.single}
                  onItemInvoked={onInvoiceRequestClicked}
                  styles={{ root: { maxHeight: 200, overflowY: "auto", background: "#fafafa", border: "1px solid #eee", borderRadius: 4 } }}
                />
              </div>
            </Stack>
          )}
        </Panel>
        <Panel
          isOpen={isInvoiceRequestViewPanelOpen}
          onDismiss={() => { setIsInvoiceRequestViewPanelOpen(false); setSelectedInvoiceRequest(null); }}
          headerText=""
          type={PanelType.medium}
        >
          {selectedInvoiceRequest && <InvoiceRequestCard invoice={selectedInvoiceRequest} />}
        </Panel>

        <Dialog
          hidden={!dialogVisible}
          onDismiss={() => setDialogVisible(false)}
          dialogContentProps={{
            type: dialogType === "error" ? DialogType.largeHeader : DialogType.normal,
            title: dialogType === "error" ? "Error" : "Success",
            subText: dialogMessage,
          }}
          modalProps={{
            isBlocking: false,
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={() => setDialogVisible(false)} text="OK" />
          </DialogFooter>
        </Dialog>
      </div>
    </section >
  );
};

export default CreateView;
