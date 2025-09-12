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
  Text
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
// import { SPFx } from "@pnp/sp/presets/all";
interface CreateViewProps {
  sp: SPFI;
  context: any;
  projectsp: SPFI;
}
type PurchaseOrderItem = {
  Id: number;
  POID: string;
  ProjectName?: string;
  POAmount?: string;
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
  const [fetchingInvoices, setFetchingInvoices] = useState(false);
  const [activePOIDFilter, setActivePOIDFilter] = useState<string | null>(null);
  const [childPOSelection] = useState(new Selection());
  const [invoiceAmountError, setInvoiceAmountError] = useState<string | undefined>(undefined);
  const [isDragActive, setIsDragActive] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
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
  const [mainPOs, setMainPOs] = useState<PurchaseOrderItem[]>([]);
  const [, setChildPOMap] = useState<Map<string, ChildPOItem[]>>(new Map());
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
  const columns: IColumn[] = [
    { key: "POID", name: "Purchase Order", fieldName: "POID", minWidth: 100, maxWidth: 150, isResizable: true },
    { key: "ProjectName", name: "Project Name", fieldName: "ProjectName", minWidth: 150, maxWidth: 220, isResizable: true },
    // { key: "Customer", name: "Customer", fieldName: "Customer", minWidth: 120, maxWidth: 160, isResizable: true },
    { key: "POAmount", name: "PO Amount", fieldName: "POAmount", minWidth: 120, maxWidth: 160, isResizable: true }
  ];
  const invoiceColumnsView: IColumn[] = [
    { key: "POItemTitle", name: "PO Item Title", fieldName: "POItemTitle", minWidth: 130, maxWidth: 180, isResizable: true },
    { key: "POItemValue", name: "PO Item Value", fieldName: "POItemValue", minWidth: 120, maxWidth: 140, isResizable: true },
    { key: "Amount", name: "Invoice Amount", fieldName: "Amount", minWidth: 120, maxWidth: 160, isResizable: true },
    { key: "Status", name: "Status", fieldName: "Status", minWidth: 140, maxWidth: 170, isResizable: true },
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
      name: "PO Item Value",
      fieldName: "POItemValue",
      minWidth: 120,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: ChildPOItem) => (
        <span>{item.POAmount}</span>
      ),
    },

    {
      key: "POAmount", name: "Remaining Item Value", fieldName: "POAmount", minWidth: 120, maxWidth: 150, isResizable: true, onRender: (item: ChildPOItem) => {
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
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragActive(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.type === "application/pdf" || file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || file.type === "application/vnd.ms-excel")) {
      setUploadedFile(file);
      setInvoiceFormState(prev => ({ ...prev, Attachment: file }));
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setUploadedFile(file);
      setInvoiceFormState(prev => ({ ...prev, Attachment: file }));
    }
  };

  const handleRemoveAttachment = () => {
    setInvoiceFormState(prev => ({ ...prev, Attachment: null }));
    setUploadedFile(null);
  };

  const handleInvoicePanelDismiss = () => {
    setIsInvoicePanelOpen(false);
    setInvoicePanelPO(null);
    // Clear uploaded attachments on panel close
    setInvoiceFormState(prev => ({
      ...prev,
      Attachment: null,
    }));
    setUploadedFile(null);
  };

  useEffect(() => {
    (async () => {
      setLoading(true);
      setError("");
      try {
        const items = await sp.web.lists.getByTitle("InvoicePO")
          .items
          .select("ID", "POID", "ParentPOID", "POAmount", "LineItemsJSON", "ProjectName")();

        setAllInvoicePOs(items);

        // Build a POID-to-item map
        const poidMap = new Map(items.map(item => [item.POID, item]));

        // Filter Main POs (ParentPOID empty)
        const mains = items.filter(item => !item.ParentPOID || item.ParentPOID.trim() === "")
          .map(item => ({
            Id: item.ID,
            POID: item.POID,
            ProjectName: item.ProjectName || "",
            POAmount: item.POAmount || ""
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
        setChildPOMap(childrenByMainPO);
      } catch (e: any) {
        setError("Error loading POs: " + (e.message || e));
        setMainPOs([]);
        setChildPOMap(new Map());
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

  const handleInvoiceAmountChange = (value: string) => {
    handleInvoiceFormChange("InvoiceAmount", value);

    const enteredAmount = parseFloat(value);
    if (!value) {
      setInvoiceAmountError("Invoice Amount is required.");
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
        setInvoiceAmountError(`Invoice Amount cannot exceed remaining amount: ${remainingAmount}`);
      } else {
        setInvoiceAmountError(undefined);
      }
    }
  };
  const handleOpenPanel = async () => {
    if (!selectedItem) return;

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
  useEffect(() => {
    const fetchGroups = async () => {
      try {
        const groups = await sp.web.currentUser.groups();
        setUserGroups(groups.map((g: any) => g.Title.toLowerCase()));
      } catch (error) {
        setUserGroups([]);
      }
    };
    fetchGroups();
  }, [sp]);

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

  // const isUserInPMGroup = userGroups.includes("pm");
  // const isUserInDMGroup = userGroups.includes("dm");
  // const isUserInDHGroup = userGroups.includes("dh");

  const handleOpenInvoicePanelSinglePO = async (poItem: PurchaseOrderItem, poAmount: string) => {
    setInvoicePanelPO(null);
    setIsInvoicePanelOpen(true);
    // const projectName = await getProjectNameByPOID(context, poItem.Id, poItem);

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
  const handleAttachmentChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] || null;
    setInvoiceFormState((prev) => ({
      ...prev,
      Attachment: file,
    }));
  };
  function getRemainingPOAmount(childPO: ChildPOItem, invoiceRequests: InvoiceRequest[]): number {
    const childInvoices = invoiceRequests.filter(inv => inv.POItemTitle === childPO.POID);
    const usedAmount = childInvoices.reduce((sum, inv) => sum + (inv.Amount || 0), 0);
    const originalAmount = parseFloat(childPO.POAmount) || 0;
    return originalAmount - usedAmount;
  }

  const handleOpenInvoicePanel = async (item: ChildPOItem) => {
    if (!selectedItem) return;
    setInvoicePanelLoading(true);
    setInvoicePanelPO(item);
    setIsInvoicePanelOpen(true);
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
    const filter = poids.map((po) => `PurchaseOrder eq '${po}'`).join(" or ");
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
          "Status"
        )();
      return items.map((item) => ({
        Id: item.Id,
        PurchaseOrderPO: item.PurchaseOrder,
        Amount: item.InvoiceAmount,
        Status: item.Status,
        ProjectName: item.ProjectName,
        POItemTitle: item.POItem_x0020_Title,
        POItemValue: item.POItem_x0020_Value,
        CustomerContact: item.Customer_x0020_Contact,
        Comments: item.Comments,
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
      // const sp = spfi(PROJECTS_SITE_URL).using(SPFx(context));
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
  const handleInvoiceFormSubmit = async () => {
    let addedItemId: number | null = null;
    try {
      if (invoiceAmountError || !invoiceFormState.InvoiceAmount) {
        alert(invoiceAmountError || "Invoice Amount is required.");
        return;
      }
      const userRole = await getCurrentUserRole(context, selectedItem);

      const financeStatusValue = "Submitted";

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

      const siteUrl = context.pageContext.web.absoluteUrl;
      const listName = "Invoice Requests";

      const itemUrl = `${siteUrl}/Lists/${listName}/DispForm.aspx?ID=${addedItemId}`;
      const creatorEmail = context.pageContext.user.email;
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
        <b>Invoice Amount:</b> ${invoiceFormState.InvoiceAmount}<br/>
        <b>Comments:</b> ${invoiceFormState.Comments}<br/><br/>
        <a href="${itemUrl}">Click here to view the invoice request.</a>
      `,
          });
          setDialogType("success");
          setDialogMessage("Invoice request submitted successfully!");
          setDialogVisible(true);

          // Refresh invoiceRequests data after update
          const lookupPOIDs = [invoiceFormState.POID, ...childPOItems.map(c => c.POID)];
          const updatedInvoices = await fetchInvoiceRequests(sp, lookupPOIDs);
          setInvoiceRequests(updatedInvoices);

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


      const lookupPOIDs = [invoiceFormState.POID, ...childPOItems.map(c => c.POID)];
      const updatedInvoices = await fetchInvoiceRequests(sp, lookupPOIDs);
      setInvoiceRequests(updatedInvoices);

      setIsInvoicePanelOpen(false);
      setInvoicePanelPO(null);

    } catch (error) {
      if (addedItemId !== null) {
        await sp.web.lists.getByTitle("Invoice Requests").items.getById(addedItemId).delete();
      }
      alert("Error submitting invoice request: " + (error as any)?.message);
    }
  };
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
          <div style={{ maxHeight: 500, overflowY: "auto", overflowX: "auto", border: "1px solid #eee", borderRadius: 4, background: "#fff" }}>
            <DetailsList
              items={filteredMainPOs}
              columns={columns}
              selection={selection}
              selectionMode={SelectionMode.single}
              setKey="mainPOsList"
            />
          </div>
        )}
        <Panel
          isOpen={isPanelOpen}
          onDismiss={handlePanelDismiss}
          headerText="Purchase Order"
          closeButtonAriaLabel="Close"
          type={PanelType.medium}
          isLightDismiss={false}
          isBlocking={false}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 18 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
            <TextField
              value={selectedItem?.POID || ""}
              readOnly
              disabled
              styles={{ root: { maxWidth: 280, marginBottom: 0 } }}
            />
            <TextField
              label="PO Amount"
              value={selectedItem?.POAmount || ""}
              readOnly
              disabled
              styles={{ root: { maxWidth: 280, marginTop: 10, marginBottom: 0 } }}
            />
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 7, color: "#626262" }}>PO Items:</div>
            <div>
              {fetchingChildPOs ? (
                <Spinner label="Loading child POs..." />
              ) : childPOItems.length > 0 ? (
                <DetailsList
                  items={childPOItems}
                  columns={childPOColumns}
                  selection={childPOSelection}
                  selectionMode={SelectionMode.single}
                  setKey="childPOs"
                  onActiveItemChanged={handleChildPORowClick}
                  styles={{
                    root: { background: "#fff", border: "1px solid #eee", borderRadius: 6 },
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
                <span>
                  Invoice Requests of {activePOIDFilter ?? selectedItem?.POID ?? ""}
                </span>
                {activePOIDFilter && (
                  <PrimaryButton
                    text="Show all Invoice Requests"
                    onClick={() => setActivePOIDFilter(null)}
                    styles={{ root: { marginLeft: 24 } }}
                  />
                )}

              </div>

            </div>
            <div>
              {fetchingInvoices ? (
                <Spinner label="Loading invoice requests..." />
              ) : showInvoices.length > 0 ? (
                <DetailsList
                  items={showInvoices}
                  columns={invoiceColumnsView}
                  selectionMode={SelectionMode.single}
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
          onDismiss={() => {
            handleInvoicePanelDismiss();
          }}
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
              maxWidth: 600,
              minHeight: 350,
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
          ) : invoicePanelPO === null ? (
            // SINGLE PO invoice form rendering
            <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
              <TextField label="PO ID" value={invoiceFormState.POID} readOnly disabled />
              <TextField label="Project Name" value={invoiceFormState.ProjectName} readOnly disabled />

              {/* PO Item Title and Value hidden for single PO */}

              <TextField
                label="PO Value"
                value={invoiceFormState.POAmount}
                type="number"
                disabled
              />

              <TextField
                label="Invoice Amount"
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
              <label style={{ marginTop: 8, fontWeight: 500 }}>
                Attachment <span style={{ marginLeft: 5, marginRight: 5, display: "inline-block" }}>🔗</span>
                <input type="file" style={{ display: "block", marginTop: 6 }} onChange={handleAttachmentChange} />
              </label>
              <PrimaryButton text="Submit" onClick={handleInvoiceFormSubmit} />
            </Stack>
          ) : (
            // CHILD PO invoice form rendering
            invoicePanelPO && (
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 6, marginBottom: 6 } }}>
                <TextField label="PO ID" value={invoiceFormState.POID} readOnly disabled />
                <TextField label="Project Name" value={invoiceFormState.ProjectName} readOnly disabled />
                <TextField label="PO Item Title" value={invoiceFormState.POItemTitle} readOnly disabled />
                <TextField label="PO Item Value" value={invoiceFormState.POItemValue} readOnly disabled />
                <TextField
                  label="Amount remaining"
                  value={String(
                    getRemainingPOAmount(
                      { POID: invoiceFormState.POItemTitle || "", POAmount: invoiceFormState.POItemValue || "0", Id: 0, ParentPOIndex: 0, POIndex: 0 },
                      invoiceRequests
                    )
                  )}
                  readOnly
                  disabled
                />
                <TextField
                  label="Invoice Amount"
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
                {/* <label style={{ marginTop: 8, fontWeight: 500 }}>
                  Attachment <span style={{ marginLeft: 5, marginRight: 5, display: "inline-block" }}>🔗</span>
                  <input type="file" style={{ display: "block", marginTop: 6 }} onChange={handleAttachmentChange} />
                </label> */}

                <div
                  style={{
                    margin: "24px 0 12px 0",
                    border: "2px dashed #d0d0d0",
                    borderRadius: 8,
                    padding: 28,
                    textAlign: "center",
                    background: isDragActive ? "#f6faff" : "#fafafa",
                    transition: "background 0.2s,border 0.2s",
                    position: "relative",
                    cursor: "pointer",
                    outline: isDragActive ? "2px solid #0078d4" : "none",
                  }}
                  onDragOver={e => { e.preventDefault(); setIsDragActive(true); }}
                  onDragLeave={e => { e.preventDefault(); setIsDragActive(false); }}
                  onDrop={handleDrop}
                  onClick={() => document.getElementById("custom-attachment-input")?.click()}
                >
                  <input
                    id="custom-attachment-input"
                    type="file"
                    accept=".pdf,.xls,.xlsx,application/pdf,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
                    style={{ display: "none" }}
                    onChange={handleFileChange}
                  />
                  <span style={{ fontSize: 36, color: "#bebebe" }}>
                    <i className="ms-Icon ms-Icon--Attach" aria-hidden="true"></i>
                  </span>
                  <div>Attachments</div>
                  <div style={{ marginTop: 14, color: "#888", fontSize: 16 }}>
                    Drop your file(s) here or click to upload.<br />
                    <span style={{ color: "#555", fontSize: 14 }}>Only PDF and Excel files are accepted.</span>
                  </div>
                  {uploadedFile && (
                    <div style={{ marginTop: 12, fontSize: 15, color: "#327800" }}>
                      <i className="ms-Icon ms-Icon--DocumentSet" aria-hidden="true" style={{ marginRight: 8 }} />
                      Selected: <strong>{uploadedFile.name}</strong>
                      {/* <PrimaryButton text="Remove" onClick={handleRemoveAttachment} styles={{ root: { marginLeft: 12, height: 28, minWidth: 70 } }} /> */}
                    </div>
                  )}
                  <PrimaryButton text="Remove"
                    // onClick={handleRemoveAttachment}
                    onClick={e => {
                      e.stopPropagation();
                      handleRemoveAttachment();
                    }}
                    disabled={!uploadedFile}
                    styles={{ root: { marginLeft: 12, height: 28, minWidth: 70 } }} />
                </div>
                <PrimaryButton text="Submit" onClick={handleInvoiceFormSubmit} />
              </Stack>
            )
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
                  styles={{ root: { maxHeight: 200, overflowY: "auto", background: "#fafafa", border: "1px solid #eee", borderRadius: 4 } }}
                />
              </div>
            </Stack>
          )}
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
