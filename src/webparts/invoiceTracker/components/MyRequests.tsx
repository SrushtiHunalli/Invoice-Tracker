import * as React from "react";
import { useState, useEffect } from "react";
import {
  DetailsList,
  SelectionMode,
  IColumn,
  Spinner,
  TextField,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Selection,
  Dropdown,
  IDropdownOption,
  Stack,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  Label,
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import DocumentViewer from "../components/DocumentViewer";

interface MyProps {
  sp: SPFI;
  context: any;
  initialFilters?: {
    searchText?: string;
    projectName?: string;
    Status?: string;
    FinanceStatus?: string;
  };
  onNavigate?: (pageKey: string, params?: any) => void;
  projectsp: SPFI;
}

interface InvoiceRequest {
  Id: number;
  Title: string;
  Status: string;
  PurchaseOrderId: number;
  PurchaseOrder: string;
  ProjectName?: string;
  InvoiceAmount?: number;
  "POItem_x0020_Title"?: string;
  "POItem_x0020_Value"?: number;
  AttachmentFiles?: any[];
  PMCommentsHistory?: string;
  FinanceCommentsHistory?: string;
  PMStatus?: string;
  POAmount?: number;
  Customer_x0020_Contact?: string;
  FinanceStatus?: string;
}

interface InvoicePO {
  Id: number;
  Title: string;
  POID: string;
  ParentPOID?: string;
  LineItems?: string;
  POAmount?: string;
  LineItemsJSON?: string;
}

interface POHierarchy {
  mainPO: InvoicePO;
  lineItemGroups: { poItem: any; requests: InvoiceRequest[] }[];
  childPOGroups: { childPO: InvoicePO; requests: InvoiceRequest[] }[];
  mainPORequests: InvoiceRequest[];
}
const steps = ["Submitted", "Not Generated", "Pending Payment", "Payment Received"];

function StatusStepper({ currentStatus, steps }: { currentStatus: string; steps: string[] }) {
  const currentStep = steps.indexOf(currentStatus);
  return (
    <div style={{ margin: "40px 0 16px 0" }}>
      <div style={{ display: "flex", alignItems: "center" }}>
        {steps.map((step, idx) => {
          let circleBorder = "#E5AF5";
          let circleBg = "#fff";
          let dotColor = "#166BDD";
          let connectorBg = "#E5AF5";
          let dot = null;
          if (idx === steps.length - 1 && currentStep === idx) {
            circleBorder = "#20bb55";
            circleBg = "#20bb55";
            dot = <span style={{ fontWeight: "bold", fontSize: 18, color: "#fff" }}>✓</span>;
          } else if (idx === currentStep) {
            dot = <span style={{ width: 10, height: 10, borderRadius: "50%", background: dotColor, display: "block" }} />;
            circleBorder = "#166BDD";
          } else if (idx < currentStep) {
            circleBorder = "#166BDD";
            circleBg = "#166BDD";
            dot = <span style={{ fontWeight: "bold", fontSize: 18, color: "#fff" }}>✓</span>;
            connectorBg = "#166BDD";
          }
          return (
            <React.Fragment key={`step-${step}`}>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                <div
                  style={{
                    width: 28,
                    height: 28,
                    borderRadius: "50%",
                    border: `2px solid ${circleBorder}`,
                    background: circleBg,
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    marginBottom: 6,
                    fontWeight: 600,
                  }}
                >
                  {dot}
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: idx <= currentStep ? (idx === steps.length - 1 && currentStep >= idx ? "#20bb55" : "#166BDD") : "#A0A5AF",
                    fontWeight: idx === currentStep ? 600 : 400,
                    textAlign: "center",
                    minWidth: 72,
                    userSelect: "none",
                  }}
                >
                  {step}
                </div>
              </div>
              {idx < steps.length - 1 && <div style={{ flex: 1, height: 2, background: connectorBg, margin: "0 4px" }} />}
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
}

// function InvoiceDetailsCard({ item, onShowAttachment }: { item: InvoiceRequest; onShowAttachment: (url: string, name: string) => void }) {
//   if (!item) return null;
//   const hidePOItem = !item["POItem_x0020_Title"] && !item["POItem_x0020_Value"];

//   const rows = [
//     { label: "Purchase Order", value: item.PurchaseOrder || "-" },
//     { label: "Project Name", value: item.ProjectName || "-" },
//     {
//       label: "PO Amount",
//       value: item.POAmount != null && !isNaN(Number(item.POAmount))
//         ? Number(item.POAmount).toLocaleString()
//         : "-"
//     },
//     !hidePOItem && { label: "PO Item Title", value: item["POItem_x0020_Title"] || "-" },
//     !hidePOItem && {
//       label: "PO Item Value",
//       value: item["POItem_x0020_Value"] != null && !isNaN(Number(item["POItem_x0020_Value"]))
//         ? Number(item["POItem_x0020_Value"]).toLocaleString()
//         : "-"
//     },
//     {
//       label: "Invoice Amount",
//       value: item.InvoiceAmount != null && !isNaN(Number(item.InvoiceAmount))
//         ? Number(item.InvoiceAmount).toLocaleString()
//         : "-"
//     },
//     { label: "Invoice Status", value: item.Status || "-" },
//   ].filter(Boolean) as { label: string; value: any }[];


//   return (
//     <div style={{ padding: 20, maxWidth: 700, margin: "auto", backgroundColor: "white", borderRadius: 8, boxShadow: "0 0 10px rgba(0,0,0,0.1)" }}>
//       {/* <Icon iconName="Info" style={{ fontSize: 24, color: "#1875f0" }} /> */}
//       <h2 style={{ marginTop: 0, marginBottom: 16 }}>All Invoice Requests for {item.PurchaseOrder}</h2>
//       <strong>All Invoice Requests for {item.PurchaseOrder}</strong>
//       <div>
//         {rows.map((row) => (
//           <div key={row.label} style={{ marginBottom: 8 }}>
//             <strong>{row.label}:</strong> {row.value}
//           </div>
//         ))}
//       </div>
//       <StatusStepper currentStatus={item.Status ?? ""} steps={steps} />
//       {item.AttachmentFiles && item.AttachmentFiles.length > 0 && (
//         <div style={{ marginTop: 24 }}>
//           <h3>Attachments</h3>
//           <ul style={{ paddingInlineStart: 20 }}>
//             {item.AttachmentFiles.map((file) => (
//               <li key={file.UniqueId}>
//                 <a href="#" onClick={(e) => {
//                   e.preventDefault();
//                   onShowAttachment(file.ServerRelativeUrl, file.FileName);
//                 }}>
//                   {file.FileName}
//                 </a>
//               </li>
//             ))}
//           </ul>
//         </div>
//       )}
//     </div>
//   );
// };

function InvoiceDetailsCard({
  item,
  onShowAttachment,
}: {
  item: InvoiceRequest;
  onShowAttachment: (url: string, name: string) => void;
}) {
  if (!item) return null;
  const hidePOItem = !item["POItem_x0020_Title"] && !item["POItem_x0020_Value"];

  const detailRows = [
    { label: "Purchase Order", value: item.PurchaseOrder || "-" },
    { label: "Project Name", value: item.ProjectName || "-" },
    {
      label: "PO Amount",
      value: item.POAmount != null && !isNaN(Number(item.POAmount))
        ? Number(item.POAmount).toLocaleString()
        : "-",
    },
    !hidePOItem && { label: "PO Item Title", value: item["POItem_x0020_Title"] || "-" },
    !hidePOItem && {
      label: "PO Item Value",
      value:
        item["POItem_x0020_Value"] != null && !isNaN(Number(item["POItem_x0020_Value"]))
          ? Number(item["POItem_x0020_Value"]).toLocaleString()
          : "-",
    },
    {
      label: "Invoice Amount",
      value: item.InvoiceAmount != null && !isNaN(Number(item.InvoiceAmount))
        ? Number(item.InvoiceAmount).toLocaleString()
        : "-",
    },
    { label: "Invoice Status", value: item.Status || "-" },
  ].filter(Boolean) as { label: string; value: any }[];

  return (
    <div
      style={{
        width: "100%",
        maxWidth: 1100,
        borderRadius: 14,
        background: "white",
        boxShadow: "0 2px 16px rgba(45,55,72,0.06)",
        margin: "20px auto",
        padding: 28,
        boxSizing: "border-box",
        display: "flex",
        flexDirection: "column",
        gap: 0,
      }}
    >
      <h2 style={{ margin: "0 0 6px 0", fontWeight: 700, fontSize: 22 }}>
        Invoice Details <span style={{ color: "#166BDD" }}>{item.PurchaseOrder}</span>
      </h2>
      <div style={{ paddingBottom: 10, fontWeight: 500 }}>
        {detailRows.map(row => (
          <div key={row.label} style={{ margin: "0 0 2px 0", fontSize: 16 }}>
            <span style={{ fontWeight: 600 }}>{row.label}:</span>{" "}
            <span>{row.value}</span>
          </div>
        ))}
      </div>
      <div style={{ margin: "18px 0" }}>
        <StatusStepper currentStatus={item.Status ?? ""} steps={steps} />
      </div>
      {item.AttachmentFiles && item.AttachmentFiles.length > 0 && (
        <div style={{ marginTop: 14 }}>
          <strong>Attachments</strong>
          <ul style={{ paddingInlineStart: 18, margin: "7px 0" }}>
            {item.AttachmentFiles.map(file => (
              <li key={file.UniqueId} style={{ margin: "3px 0" }}>
                <a
                  href="#"
                  onClick={e => {
                    e.preventDefault();
                    onShowAttachment(file.ServerRelativeUrl, file.FileName);
                  }}
                  style={{ color: "#166BDD", textDecoration: "underline" }}
                >
                  {file.FileName}
                </a>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

export default function MyRequests({ sp, projectsp, context, initialFilters }: MyProps) {
  const [invoicePOs, setInvoicePOs] = useState<InvoicePO[]>([]);
  const [invoiceRequests, setInvoiceRequests] = useState<InvoiceRequest[]>([]);
  const [poHierarchy, setPOHierarchy] = useState<null | {
    mainPO: InvoicePO;
    lineItemGroups: { poItem: any; requests: InvoiceRequest[] }[];
    childPOGroups: { childPO: InvoicePO; requests: InvoiceRequest[] }[];
    mainPORequests: InvoiceRequest[];
  }>(null);
  const [selectedReq, setSelectedReq] = useState<InvoiceRequest | null>(null);
  const [selectedPOItem, setSelectedPOItem] = useState<{ POID: string; POAmount: string } | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [showHierPanel, setShowHierPanel] = useState(false);
  const [viewerUrl, setViewerUrl] = useState<string | null>(null);
  const [viewerName, setViewerName] = useState<string | null>(null);
  const [showClarifyPanel, setShowClarifyPanel] = useState(false);
  const [clarifyInvoiceAmount, setClarifyInvoiceAmount] = useState<number | undefined>();
  const [clarifyCustomerContact, setClarifyCustomerContact] = useState<string | undefined>();
  const [clarifyComment, setClarifyComment] = useState("");
  const [clarifyLoading, setClarifyLoading] = useState(false);
  const [searchText, setSearchText] = useState("");
  const [filterProjectName, setFilterProjectName] = useState<string | undefined>(undefined);
  const [filterStatus, setFilterStatus] = useState<string | undefined>(undefined);
  const [filterFinanceStatus, setFilterFinanceStatus] = useState<string | undefined>(undefined);
  const [projectOptions, setProjectOptions] = useState<string[]>([]);
  const [statusOptions, setStatusOptions] = useState<string[]>([]);
  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");
  const [dialogType, setDialogType] = useState<"success" | "error">("success");
  const [selectedProject, setSelectedProject] = useState<any | null>(null);
  const toDropdownOptions = (items: string[]): IDropdownOption[] => [
    { key: "All", text: "All" },
    ...items.map(item => ({ key: item, text: item }))
  ];

  // Columns for Invoice requests list:
  const invoiceColumns: IColumn[] = [
    // { key: "Title", name: "Title", fieldName: "Title", minWidth: 170, maxWidth: 270, isResizable: true },
    { key: "PurchaseOrder", name: "POID", fieldName: "PurchaseOrder", minWidth: 100, maxWidth: 160, isResizable: true },
    { key: "ProjectName", name: "Project", fieldName: "ProjectName", minWidth: 150, maxWidth: 220, isResizable: true },
    {
      key: "Finance Status",
      name: "Current Status",
      fieldName: "FinanceStatus",
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      onRender: (item) => item.FinanceStatus || "-"
    },
    { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 120, maxWidth: 160, isResizable: true },
    {
      key: "POItem_x0020_Title",
      name: "PO Item Title",
      fieldName: "POItem_x0020_Title",
      minWidth: 150,
      maxWidth: 220,
      isResizable: true,
      onRender: item => item["POItem_x0020_Title"] || "-"
    },
    {
      key: "POItem_x0020_Value",
      name: "PO Item Value",
      fieldName: "POItem_x0020_Value",
      minWidth: 140,
      maxWidth: 160,
      isResizable: true,
      onRender: (item) =>
        item["POItem_x0020_Value"] != null && !isNaN(Number(item["POItem_x0020_Value"]))
          ? Number(item["POItem_x0020_Value"]).toLocaleString()
          : "-"
    },
    {
      key: "InvoiceAmount",
      name: "Invoice Amount",
      fieldName: "InvoiceAmount",
      minWidth: 150,
      maxWidth: 160,
      isResizable: true,
      onRender: (item) => item.InvoiceAmount ? item.InvoiceAmount.toLocaleString() : "",
    },
  ];

  // Columns for PO items:
  const poColumns: IColumn[] = [
    { key: "POID", name: "POItem Title", fieldName: "POID", minWidth: 150, maxWidth: 220, isResizable: true },
    { key: "POAmount", name: "POItem Amount", fieldName: "POAmount", minWidth: 140, maxWidth: 160, isResizable: true },
  ];

  const poColumnsLine: IColumn[] = [
    { key: "POItem_x0020_Title", name: "POItem Title", fieldName: "POItem_x0020_Title", minWidth: 150, maxWidth: 220, isResizable: true },
    { key: "POItem_x0020_Value", name: "POItem Amount", fieldName: "POItem_x0020_Value", minWidth: 140, maxWidth: 160, isResizable: true },
    // { key: "Comments", name: "Description", fieldName: "Comments", minWidth: 170, maxWidth: 270, isResizable: true }, // Optional
  ];


  // Columns for invoice requests grouped by PO:
  const groupedInvColumns: IColumn[] = [
    { key: "POItem_x0020_Title", name: "PO Item Title", fieldName: "POItem_x0020_Title", minWidth: 150, maxWidth: 220, isResizable: true },
    { key: "POItem_x0020_Value", name: "PO Item Value", fieldName: "POItem_x0020_Value", minWidth: 140, maxWidth: 160, isResizable: true },
    {
      key: "InvoiceAmount",
      name: "Invoice Amount",
      fieldName: "InvoiceAmount",
      minWidth: 150,
      maxWidth: 160,
      isResizable: true,
      onRender: (item) => item.InvoiceAmount ? item.InvoiceAmount.toLocaleString() : "",
    },
    { key: "Invoice Status", name: "Status", fieldName: "Status", minWidth: 120, maxWidth: 160, isResizable: true },
  ];

  const [selection] = useState(
    new Selection({
      onSelectionChanged: () => {
        const selected = selection.getSelection()[0] as InvoiceRequest | undefined;
        onInvoiceRequestSelect(selected);
      }
    })
  );
  const [clearCounter, setClearCounter] = useState(0);

  const clearAllFilters = () => {
    setSearchText("");
    setFilterProjectName("All");
    setFilterStatus("All");
    setFilterFinanceStatus("All");
    setClearCounter(clearCounter + 1);
  };


  useEffect(() => {
    async function loadData() {
      setLoading(true);
      try {
        const [pos, reqs] = await Promise.all([
          sp.web.lists.getByTitle("InvoicePO").items(),
          sp.web.lists.getByTitle("Invoice Requests").items.select("*").expand("AttachmentFiles")(),
        ]);
        setInvoicePOs(pos);
        setInvoiceRequests(reqs);
        setProjectOptions(Array.from(new Set(reqs.map(r => r.ProjectName).filter(Boolean))));
        setStatusOptions(Array.from(new Set(reqs.map(r => r.Status).filter(Boolean))));
        setError(null);
      } catch (err: any) {
        setError(`Error loading data: ${err.message || err}`);
      } finally {
        setLoading(false);
      }
    }
    loadData();
  }, [sp]);

  useEffect(() => {
    if (initialFilters) {
      if (initialFilters.searchText !== undefined) setSearchText(initialFilters.searchText);
      if (initialFilters.projectName !== undefined) setFilterProjectName(initialFilters.projectName);
      if (initialFilters.Status !== undefined) setFilterStatus(initialFilters.Status);
      if (initialFilters.FinanceStatus !== undefined) setFilterFinanceStatus(initialFilters.FinanceStatus);
    }
    // empty dep array if initialFilters won't change after mount
  }, [initialFilters]);

  useEffect(() => {
    console.log("Filters set from initialFilters:", {
      searchText,
      filterProjectName,
      filterStatus,
      filterFinanceStatus,
    });
  }, [searchText, filterProjectName, filterStatus, filterFinanceStatus]);

  useEffect(() => {
    if (selectedReq?.ProjectName) {
      loadProject(selectedReq.ProjectName);
    } else {
      setSelectedProject(null);
    }
  }, [selectedReq]);

  async function loadProject(projectName?: string) {
    if (!projectName) {
      setSelectedProject(null);
      return;
    }
    try {
      const project = await projectsp.web.lists
        .getByTitle("Projects")
        .items
        .filter(`Title eq '${projectName}'`)
        .select("*")() // or select/expand if needed
        .then(items => items[0]);

      setSelectedProject(project || null);
      console.log("Loaded project:", project);
    } catch (error) {
      console.error("Failed to load project", error);
      setSelectedProject(null);
    }
  }

  function findMainPO(request: InvoiceRequest, allPOs: InvoicePO[]): InvoicePO | undefined {
    let po = allPOs.find((p) => p.POID === request.PurchaseOrder);
    while (po && po.ParentPOID) {
      po = allPOs.find((p) => p.POID === po.ParentPOID);
    }
    return po;
  }

  function getLineItemsList(h: POHierarchy | null) {
    if (!h) return [];
    // Adapt line items to POItems table structure for the chilPO table UI
    return h.lineItemGroups.map(g => ({
      // Map 'Title' to POItem Title, 'Value' to POItem Value
      POItem_x0020_Title: g.poItem.Title,              // Displayed as POItem Title
      POItem_x0020_Value: g.poItem.Value,              // Displayed as POItem Value
      Comments: g.poItem.Comments || "",               // Optional, use as Description or Comments
    }));
  }


  function decodeHtmlEntities(str: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = str;
    return txt.value;
  }

  function formatCommentsHistory(historyJson?: string) {
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


  async function handleClarifySubmit() {
    setClarifyLoading(true);

    try {

      let history: any[] = [];
      if (selectedReq.PMCommentsHistory) {
        try {
          history = JSON.parse(selectedReq.PMCommentsHistory);
          if (!Array.isArray(history)) history = [];
        } catch {
          history = [];
        }
      }

      const userRole = await getCurrentUserRole(context, selectedReq);
      history.push({
        Date: new Date().toISOString(),
        Title: "Clarification",
        User: context.pageContext.user.displayName || "Unknown User",
        Role: userRole,
        Data: clarifyComment,
      });

      await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.getById(selectedReq.Id)
        .update({
          InvoiceAmount: clarifyInvoiceAmount,
          PMCommentsHistory: JSON.stringify(history),
          PMStatus: "Submitted",
          FinanceStatus: "Pending",
          Customer_x0020_Contact: clarifyCustomerContact,
        });

      setShowClarifyPanel(false);

      setDialogType("success");
      setDialogMessage("Clarification submitted successfully!");
      setDialogVisible(true);

      // Refresh invoiceRequests data after update
      const refreshedReqs = await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.select("*")
        .expand("AttachmentFiles")();
      setInvoiceRequests(refreshedReqs);
    } catch (err) {
      setDialogType("error");
      setDialogMessage("Error updating invoice: " + (err.message || err.toString()));
      setDialogVisible(true);
    } finally {
      setClarifyLoading(false);
    }
  }

  function decodeHtml(html: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = html;
    return txt.value;
  }


  function getHierarchyForPO(
    mainPO: InvoicePO,
    allPOs: InvoicePO[],
    allRequests: InvoiceRequest[]
  ): POHierarchy | null {
    const childPOs = allPOs.filter((p) => p.ParentPOID === mainPO.POID);

    // let lineItems: any[] = [];
    // if (mainPO.LineItems) {
    //   try {
    //     lineItems = JSON.parse(mainPO.LineItems);
    //   } catch {
    //     lineItems = [];
    //   }
    // }
    let lineItems: any[] = [];
    if (mainPO.LineItemsJSON) {
      try {
        // Decode HTML entities from rich text column first
        const decoded = decodeHtml(mainPO.LineItemsJSON);
        lineItems = JSON.parse(decoded);
      } catch (err) {
        console.error("Failed to parse LineItemsJSON", err, mainPO.LineItemsJSON);
        lineItems = [];
      }
    }


    // If no lineItems but LineItemsJSON exists, parse it after stripping HTML
    if ((!lineItems || lineItems.length === 0) && mainPO.LineItemsJSON) {
      try {
        const raw = mainPO.LineItemsJSON.replace(/<\/?[^>]+(>|$)/g, "");
        lineItems = JSON.parse(raw);
      } catch {
        lineItems = [];
      }
    }

    // Case 3: No children & no line items → return null (no hierarchy)
    if (childPOs.length === 0 && lineItems.length === 0) {
      return null;
    }

    // const lineItemGroups = lineItems.map((li) => ({
    //   poItem: li,
    //   requests: allRequests.filter(
    //     (req) =>
    //       req.PurchaseOrder === mainPO.POID &&
    //       req["POItem_x0020_Title"] === li.POID
    //   ),
    // }));

    const lineItemGroups = lineItems.map((li) => ({
      poItem: li,
      requests: allRequests.filter(
        (req) =>
          req.PurchaseOrder === mainPO.POID &&
          req["POItem_x0020_Title"] === li.POID
      ),
    }));

    const childPOGroups = childPOs.map((child) => ({
      childPO: child,
      requests: allRequests.filter((req) => req.PurchaseOrder === child.POID),
    }));

    const mainPORequests = allRequests.filter(
      (req) =>
        req.PurchaseOrder === mainPO.POID &&
        (!req["POItem_x0020_Title"] ||
          !lineItems.some((li) => li.POID === req["POItem_x0020_Title"]))
    );

    return { mainPO, lineItemGroups, childPOGroups, mainPORequests };
  }

  function flattenRequests(groups: any[], keyProperty: string, keyValueProp: string) {
    return groups.reduce((acc: any[], group) => {
      const mapped = group.requests.map((r: any) => ({
        ...r,
        [keyProperty]: group[keyValueProp].POID,
        POItem_x0020_Value: group[keyValueProp].POAmount,
      }));
      return acc.concat(mapped);
    }, []);
  }

  function flattenRequestsLine(groups: any[], keyProperty: string, keyValueProp: string) {
    return groups.reduce((acc: any[], group) => {
      const mapped = group.requests.map((r: any) => ({
        ...r,
        [keyProperty]: group[keyValueProp].Title,      // for "POItem_x0020_Title"
        POItem_x0020_Value: group[keyValueProp].Value, // for "POItem_x0020_Value"
        // add Description, Comments, etc. if needed
      }));
      return acc.concat(mapped);
    }, []);
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
  function getAllPOItemsForList(h: POHierarchy | null) {
    if (!h) {
      console.log("getAllPOItemsForList: Received null or undefined hierarchy");
      return [];
    }
    console.log("getAllPOItemsForList: input hierarchy:", h);

    const lineItems = h.lineItemGroups.map((g) => {
      console.log("Processing line item group:", g);
      return {
        POID: g.poItem.POID,
        POAmount: g.poItem.POAmount,
      };
    });

    const childPOs = h.childPOGroups.map((g) => {
      console.log("Processing child PO group:", g);
      return {
        POID: g.childPO.POID,
        POAmount: g.childPO.POAmount,
      };
    });

    const combined = [...lineItems, ...childPOs];
    console.log("getAllPOItemsForList: combined PO items:", combined);
    return combined;
  }


  function getFilteredRequests(): InvoiceRequest[] {
    if (!poHierarchy) return [];

    if (!selectedPOItem) {
      // Show all requests for PO hierarchy
      const lineItemRequests = flattenRequestsLine(poHierarchy.lineItemGroups, "Title", "Value");
      const childPORequests = flattenRequests(poHierarchy.childPOGroups, "POItem_x0020_Title", "childPO");
      return [
        ...poHierarchy.mainPORequests,
        ...lineItemRequests,
        ...childPORequests,
      ];
    } else {
      // Filter by selected PO Item Title (for line items or child PO)
      const allRequests = [
        ...poHierarchy.mainPORequests,
        ...flattenRequests(poHierarchy.lineItemGroups, "POItem_x0020_Title", "poItem"),
        ...flattenRequests(poHierarchy.childPOGroups, "POItem_x0020_Title", "childPO"),
      ];
      return allRequests.filter(
        req => req["POItem_x0020_Title"] === selectedPOItem.POID
      );
    }
  }

  const filteredInvoiceRequests = invoiceRequests.filter(item => {
    const searchLower = searchText.toLowerCase().trim();
    const matchesSearch = !searchLower || Object.values(item).some(val =>
      val !== undefined &&
      val !== null &&
      val.toString().toLowerCase().includes(searchLower)
    );
    const matchesProject = !filterProjectName || filterProjectName === "All" || item.ProjectName === filterProjectName;
    const matchesStatus = !filterStatus || filterStatus === "All" || item.Status === filterStatus;
    const matchesFinanceStatus = !filterFinanceStatus || filterFinanceStatus === "All" || item.FinanceStatus === filterFinanceStatus;

    return matchesSearch && matchesProject && matchesStatus && matchesFinanceStatus;
  });


  async function onInvoiceRequestSelect(item?: InvoiceRequest) {
    if (item) {
      setSelectedReq(item);

      // Load project details here directly
      if (item.ProjectName) {
        try {
          const project = await projectsp.web.lists
            .getByTitle("Projects")
            .items
            .filter(`Title eq '${item.ProjectName}'`)
            .select("Title", "PM/Title", "PM/EMail", "DM/Title", "DM/EMail", "DH/Title", "DH/EMail") // Select sub-fields explicitly
            .expand("PM", "DM", "DH")
            .top(1)()
            .then(items => items[0]);

          setSelectedProject(project || null);
        } catch (error) {
          console.error("Failed to load project", error);
          setSelectedProject(null);
        }
      } else {
        setSelectedProject(null);
      }

      // Existing logic for PO hierarchy etc.
      const mainPO = findMainPO(item, invoicePOs);
      if (!mainPO) {
        setPOHierarchy(null);
        return;
      }
      const hierarchy = getHierarchyForPO(mainPO, invoicePOs, invoiceRequests);
      setPOHierarchy(hierarchy);
      setShowHierPanel(true);

      if (hierarchy) {
        const selectedPO =
          hierarchy.lineItemGroups.find(
            (g) => g.poItem.POID === item["POItem_x0020_Title"]
          ) || hierarchy.childPOGroups.find(
            (g) => g.childPO.POID === item["POItem_x0020_Title"]
          );

        setSelectedPOItem(
          selectedPO
            ? 'poItem' in selectedPO
              ? { POID: selectedPO.poItem.Title, POAmount: selectedPO.poItem.Value }
              : { POID: selectedPO.childPO.POID, POAmount: selectedPO.childPO.POAmount }
            : null
        );
      } else {
        setSelectedPOItem(null);
      }
    } else {
      // Clear everything if no item selected
      setSelectedReq(null);
      setSelectedProject(null);
      setPOHierarchy(null);
      setSelectedPOItem(null);
      setShowHierPanel(false);
    }
  }

  function normalizeSelectedPOItem(item: any): { POID: string, POAmount: string } | null {
    if (!item) return null;
    return {
      POID: item.POID ?? item.POItem_x0020_Title,
      POAmount: item.POAmount ?? item.POItem_x0020_Value
    };
  }

  return (
    <div style={{ padding: 16 }}>
      <h2>My Invoice Requests</h2>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" styles={{ root: { marginBottom: 12 } }}>
        <div>
          <Label>Search</Label>
          <TextField
            placeholder="Search"
            value={searchText}
            onChange={(e, val) => setSearchText(val || "")}
            styles={{ root: { width: 320 } }}  // Double length
          />
        </div>
        <div>
          <Label>Project Name</Label>
          <Dropdown
            placeholder="Project Name"
            options={toDropdownOptions(projectOptions)}
            selectedKey={filterProjectName || undefined}
            onChange={(e, option) => setFilterProjectName(option?.key as string || "All")}
            styles={{ root: { width: 160 } }}
            ariaLabel="Project Name"
          />
        </div>
        <div>
          <Label>Current Status</Label>
          <Dropdown
            placeholder="Current Status"
            options={toDropdownOptions(
              Array.from(new Set(invoiceRequests.map(r => r.FinanceStatus).filter(Boolean)))
            )}
            selectedKey={filterFinanceStatus || undefined}
            onChange={(e, option) => setFilterFinanceStatus(option?.key as string || "All")}
            styles={{ root: { width: 160 } }}
            ariaLabel="Current Status"
          />
        </div>
        <div>
          <Label>Invoice Status</Label>
          <Dropdown
            placeholder="Invoice Status"
            options={toDropdownOptions(statusOptions)}
            selectedKey={filterStatus || undefined}
            onChange={(e, option) => setFilterStatus(option?.key as string || "All")}
            styles={{ root: { width: 160 } }}
            ariaLabel="Invoice Status"
          />
        </div>
        <div>
          <PrimaryButton
            text="Clear"
            onClick={clearAllFilters}
            style={{ alignSelf: "center", marginLeft: 20 }}
            disabled={searchText === "" && 
              (filterProjectName === "All" || !filterProjectName) && 
              (filterStatus === "All" || !filterStatus) && 
              (filterFinanceStatus === "All" || !filterFinanceStatus)
            }
          />
        </div>
      </Stack>


      {loading ? (
        <Spinner label="Loading..." />
      ) : (
        <>
          <DetailsList
            items={filteredInvoiceRequests}
            columns={invoiceColumns}
            selectionMode={SelectionMode.single}
            onActiveItemChanged={onInvoiceRequestSelect}
            setKey="invoiceRequestList"
          />
          <Panel
            isOpen={showHierPanel}
            onDismiss={() => {
              if (showClarifyPanel) return;
              setShowHierPanel(false);
              setSelectedReq(null);
              setSelectedPOItem(null);
              setPOHierarchy(null);
              // setViewerUrl(null);
            }}

            // headerText={`Invoice Details: ${poHierarchy.mainPO.POID}`}
            type={PanelType.largeFixed}
            // isLightDismiss
            closeButtonAriaLabel="Close"
          >
            {selectedReq && selectedProject && (
              <>
                <InvoiceDetailsCard
                  item={selectedReq}
                  onShowAttachment={(url, name) => {
                    setViewerUrl(url);
                    setViewerName(name);
                  }}
                />
                {selectedReq.PMStatus === "Pending" && (
                  <div style={{ margin: "16px 0" }}>
                    <button
                      onClick={() => {
                        setClarifyInvoiceAmount(selectedReq.InvoiceAmount);
                        setClarifyCustomerContact(selectedReq.Customer_x0020_Contact);
                        setClarifyComment("");
                        setShowClarifyPanel(true);
                      }}
                      style={{ padding: '8px 24px', background: '#166BDD', color: '#fff', borderRadius: 4, border: 'none' }}
                    >
                      Clarify
                    </button>
                  </div>
                )}
              </>
            )}
            {poHierarchy && poHierarchy.lineItemGroups.length > 0 && (
              <div style={{ marginTop: 16 }}>
                <strong>All PO Items for {poHierarchy.mainPO.POID}</strong>
                <DetailsList
                  items={getLineItemsList(poHierarchy)}
                  columns={poColumnsLine}
                  selectionMode={SelectionMode.single}
                  onActiveItemChanged={(item) => setSelectedPOItem(normalizeSelectedPOItem(item))}
                  setKey="lineItemsList"
                  styles={{ root: { marginBottom: 16 } }}
                />
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <strong>
                    Invoice Requests for {selectedPOItem ? selectedPOItem.POID : poHierarchy.mainPO.POID}
                  </strong>
                  {selectedPOItem && (
                    <PrimaryButton
                      text={`Show all Invoice Requests`}
                      onClick={() => setSelectedPOItem(null)}
                      style={{
                        marginLeft: 12,
                        color: "white",
                        background: "#166BDD",
                        fontWeight: 600,
                        borderRadius: 4,
                        padding: "4px 16px"
                      }}
                      // title={`Show all Invoice Requests for ${poHierarchy.mainPO.POID}`}
                    />
                  )}

                </div>

                <DetailsList
                  items={getFilteredRequests()}
                  columns={groupedInvColumns}
                  selectionMode={SelectionMode.single}
                  // selection={selection}
                  setKey="invoiceRequestsListByPO"
                />

              </div>
            )}

            {poHierarchy && poHierarchy.lineItemGroups.length < 1 && (
              <div style={{ marginTop: 20 }}>
                {/* <h3>PO Hierarchy</h3> */}
                <div style={{ marginBottom: 12 }}>
                  {/* <strong>{poHierarchy.mainPO.POID}</strong> */}
                  <strong>All PO Items for {poHierarchy.mainPO.POID}</strong>
                </div>
                <DetailsList
                  items={getAllPOItemsForList(poHierarchy)}
                  columns={poColumns}
                  selectionMode={SelectionMode.single}
                  onActiveItemChanged={(item) => setSelectedPOItem(item)}
                  setKey="poItemsList"
                  styles={{ root: { marginBottom: 20 } }}
                />
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <strong>
                    Invoice Requests for {selectedPOItem ? selectedPOItem.POID : poHierarchy.mainPO.POID}
                  </strong>
                  {selectedPOItem && (
                    <PrimaryButton
                      text="Show all Invoice Requests"
                      onClick={() => setSelectedPOItem(null)}
                      style={{
                        marginLeft: 12,
                        color: "white",
                        background: "#166BDD",
                        fontWeight: 600,
                        borderRadius: 4,
                        padding: "4px 16px"
                      }}
                      title={`Show all Invoice Requests for ${poHierarchy.mainPO.POID}`}
                    />
                  )}

                </div>

                <DetailsList
                  items={getFilteredRequests()}
                  columns={groupedInvColumns}
                  selectionMode={SelectionMode.single}
                  // selection={selection}
                  setKey="invoiceRequestsListByPO"
                />
              </div>
            )}
            {/* {!poHierarchy && selectedReq && (
              <div style={{ marginTop: 20 }}>
                <h3>Invoice Requests</h3>
                <DetailsList
                  items={invoiceRequests.filter(
                    req => req.PurchaseOrder === selectedReq.PurchaseOrder
                  )}
                  columns={invoiceColumns}
                  selectionMode={SelectionMode.single}
                  // Optionally, activate selection:
                  // onActiveItemChanged={onInvoiceRequestSelect}
                  setKey="singlePOInvoiceRequestList"
                />
              </div>
            )} */}

          </Panel>
          <Panel
            isOpen={showClarifyPanel}
            onDismiss={() => setShowClarifyPanel(false)}
            headerText={`Clarify Invoice: ${selectedReq?.Title ?? ''}`}
            type={PanelType.medium}
            isLightDismiss={false}
            hasCloseButton={true}
          >
            {selectedReq && (
              <>
                <TextField
                  label="POID"
                  value={selectedReq.PurchaseOrder || ''}
                  disabled
                />
                <TextField
                  label="PO Amount"
                  value={selectedReq?.POAmount?.toString() || ''}
                  disabled
                />
                <TextField
                  label="PO Item Title"
                  value={selectedReq["POItem_x0020_Title"] || ''}
                  disabled
                />
                <TextField
                  label="PO Item Value"
                  value={selectedReq["POItem_x0020_Value"]?.toString() || ''}
                  disabled
                />
                <TextField
                  label="Customer Contact"
                  value={clarifyCustomerContact?.toString() || ''}
                  onChange={(_, val) => setClarifyCustomerContact(val || '')}
                  required
                />
                <TextField
                  label="Invoice Amount"
                  value={clarifyInvoiceAmount?.toString() || ''}
                  onChange={(_, val) => setClarifyInvoiceAmount(val ? Number(val) : undefined)}
                  required
                />
                <TextField
                  label="Add Comment"
                  multiline
                  value={clarifyComment}
                  onChange={(_, val) => setClarifyComment(val || '')}
                />
                <TextField
                  label="PM Comment History"
                  value={formatCommentsHistory(selectedReq.PMCommentsHistory)}
                  multiline
                  disabled
                />
                <TextField
                  label="Finance Comments History"
                  value={formatCommentsHistory(selectedReq.FinanceCommentsHistory)}
                  multiline
                  disabled
                />
                <div style={{ marginTop: 12 }}>
                  <button
                    type="button"
                    disabled={clarifyLoading || clarifyInvoiceAmount === undefined}
                    onClick={handleClarifySubmit}
                    style={{
                      background: '#20bb55',
                      color: '#fff',
                      padding: '8px 24px',
                      border: 'none',
                      borderRadius: 4,
                      cursor: 'pointer'
                    }}
                  >
                    Submit
                  </button>
                </div>
              </>
            )}
          </Panel>
          <Panel
            isOpen={!!viewerUrl}
            onDismiss={() => {
              setViewerUrl(null);
            }}
            headerText={viewerName ?? "Document Viewer"}
            type={PanelType.large}
            // isLightDismiss
            closeButtonAriaLabel="Close"
          >
            {viewerUrl && viewerName && (
              <DocumentViewer url={viewerUrl} isOpen onDismiss={() => setViewerUrl(null)} fileName={viewerName} />
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

        </>
      )}
    </div>
  );
}
