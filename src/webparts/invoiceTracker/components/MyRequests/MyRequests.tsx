
import * as React from "react";
import { useState, useEffect } from "react";
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { MSGraphClient } from '@microsoft/sp-http';
import {
  DetailsList,
  Checkbox,
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
  IDropdownStyles,
  Stack,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  Label,
  Text,
  Separator,
  Icon,
  IconButton,
  IDetailsHeaderProps,
  // ComboBox,
  IComboBoxOption,
  // DetailsListLayoutMode,
  IRenderFunction,
  Sticky,
  StickyPositionType,
  ScrollablePane,
  IDropdownOption,
  ContextualMenu,
  ContextualMenuItemType,
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import DocumentViewer from "../DocumentViewer";
import SearchableDropdown from "../../SearchableDropdown/SearchableDropdown";
import styles from "./MyRequests.module.scss"
interface MyProps {
  sp: SPFI;
  context: any;
  initialFilters?: {
    searchText?: string;
    projectName?: string;
    Status?: string;
    FinanceStatus?: string;
    selectedInvoice?: number | string;
    [key: string]: any;
  };

  onNavigate?: (pageKey: string, params?: any) => void;
  projectsp: SPFI;
  getCurrentPageUrl?: () => string;
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
  CurrentStatus?: string;
  DueDate?: Date;
  Currency?: string;
  Created?: string;
  Modified?: string;
  Author?: {
    Title?: string;
    EMail?: string;
  };
  Editor?: {
    Title?: string;
    EMail?: string;
  };
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
// interface StatusStepperProps {
//   currentStatus: string;
//   steps: string[];
// }
const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";

const steps = ["Invoice Requested", "Invoice Raised", "Pending Payment", "Payment Received", "Cancelled"];
// const spTheme = (window as any).__themeState__?.theme;
// const primaryColor = spTheme?.themePrimary || "#0078d4";
// function StatusStepper({ currentStatus, steps }: { currentStatus: string; steps: string[] }) {
//   let visibleSteps: string[] = [];
//   let onlyCancelledStep = false;

//   // Logic to determine which steps to show based on currentStatus
//   if (currentStatus === "Cancelled") {

//     visibleSteps = ["Cancelled"];
//     onlyCancelledStep = true;

//   } else {
//     visibleSteps = ["Request Submitted", "Not Generated", "Invoice Raised", "Pending Payment", "Payment Received"];
//   }
//   const currentStep = visibleSteps.indexOf(currentStatus);
//   return (
//     <div style={{ margin: "40px 0 16px 0" }}>
//       <div style={{ display: "flex", alignItems: "center" }}>
//         {visibleSteps.map((visibleSteps, idx) => {
//           let circleBorder = "#E5AF5";
//           let circleBg = "#fff";
//           let dotColor = "#166BDD";
//           let connectorBg = "#E5AF5";
//           let dot = null;
//           if (onlyCancelledStep) {
//             circleBorder = "FF0000"; // red border
//             circleBg = "#fff";
//             dot = <span style={{ color: "red", fontWeight: "bold", fontSize: 18 }}>X</span>;
//           }
//           else if (idx === visibleSteps.length - 1 && currentStep === idx) {
//             circleBorder = "#20bb55";
//             circleBg = "#20bb55";
//             dot = <span style={{ fontWeight: "bold", fontSize: 18, color: "#fff" }}>✓</span>;
//           } else if (idx === currentStep) {
//             dot = <span style={{ width: 10, height: 10, borderRadius: "50%", background: dotColor, display: "block" }} />;
//             circleBorder = "#166BDD";
//           } else if (idx < currentStep) {
//             circleBorder = "#166BDD";
//             circleBg = "#166BDD";
//             dot = <span style={{ fontWeight: "bold", fontSize: 18, color: "#fff" }}>✓</span>;
//             connectorBg = "#166BDD";
//           }
//           return (
//             <React.Fragment key={`step-${visibleSteps}`}>
//               <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
//                 <div
//                   style={{
//                     width: 28,
//                     height: 28,
//                     borderRadius: "50%",
//                     border: `2px solid ${circleBorder}`,
//                     background: circleBg,
//                     display: "flex",
//                     justifyContent: "center",
//                     alignItems: "center",
//                     marginBottom: 6,
//                     fontWeight: 600,
//                   }}
//                 >
//                   {dot}
//                 </div>
//                 <div
//                   style={{
//                     fontSize: 12,
//                     color: idx <= currentStep ? (idx === visibleSteps.length - 1 && currentStep >= idx ? "#20bb55" : "#166BDD") : "#A0A5AF",
//                     fontWeight: idx === currentStep ? 600 : 400,
//                     textAlign: "center",
//                     minWidth: 72,
//                     userSelect: "none",
//                   }}
//                 >
//                   {visibleSteps}
//                 </div>
//               </div>
//               {idx < visibleSteps.length - 1 && <div style={{ flex: 1, height: 2, background: connectorBg, margin: "0 4px" }} />}
//             </React.Fragment>
//           );
//         })}
//       </div>
//     </div>
//   );
// }

function StatusStepper({ currentStatus, steps }: { currentStatus: string; steps: string[] }) {
  let visibleSteps: string[] = [];
  let onlyCancelledStep = false;

  // Logic to determine which steps to show based on currentStatus
  if (currentStatus === "Cancelled") {
    visibleSteps = ["Cancelled"];
    onlyCancelledStep = true;
  } else {
    visibleSteps = ["Invoice Requested", "Invoice Raised", "Pending Payment", "Payment Received"];
  }

  const currentStep = visibleSteps.indexOf(currentStatus);

  return (
    <div style={{ margin: "40px 0 16px 0" }}>
      <div style={{ display: "flex", alignItems: "center" }}>
        {visibleSteps.map((step, idx) => {
          let circleBorder = "#E5AF5";
          let circleBg = "#fff";
          let dotColor = "#166BDD";
          let connectorBg = "#E5AF5";
          let dot: JSX.Element | null = null;

          if (onlyCancelledStep) {
            circleBorder = "#FF0000"; // red border for Cancelled
            circleBg = "#fff";
            dot = <span style={{ color: "red", fontWeight: "bold", fontSize: 18 }}>X</span>;
          } else if (step === "Payment Received" && currentStep === idx) {
            // Current step is Payment Received
            circleBorder = "#20bb55";
            circleBg = "#1ae962ff";
            dot = <span style={{ fontWeight: "bold", fontSize: 18, color: "#fff" }}>✓</span>;
          } else if (idx === currentStep) {
            // Current step (not Payment Received)
            dot = <span style={{ width: 10, height: 10, borderRadius: "50%", background: dotColor, display: "block" }} />;
            circleBorder = "#166BDD";
          } else if (idx < currentStep) {
            // Steps before current (completed)
            circleBorder = "#1469daff";
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
                    color: idx <= currentStep ? (step === "Payment Received" && currentStep >= idx ? "#20bb55" : "#166BDD") : "#A0A5AF",
                    fontWeight: idx === currentStep ? 600 : 400,
                    textAlign: "center",
                    minWidth: 72,
                    userSelect: "none",
                  }}
                >
                  {step}
                </div>
              </div>
              {idx < visibleSteps.length - 1 && <div style={{ flex: 1, height: 2, background: connectorBg, margin: "0 4px" }} />}
            </React.Fragment>
          );
        })}
      </div>
    </div>
  );
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

function InvoiceDetailsCard({
  item,
  onShowAttachment,
}: {
  item: InvoiceRequest;
  onShowAttachment: (url: string, name: string) => void;
}) {
  if (!item) return null;
  // const hideItem = !item["POItem_x0020_Title"] && !item["POItem_x0020_Value"];
  const itemCurrency = getCurrencySymbol(item.Currency);

  function getCurrencySymbol(currencyCode: string, locale = 'en-US'): string {
    if (!currencyCode) return '';
    try {
      return new Intl.NumberFormat(locale, {
        style: 'currency',
        currency: currencyCode,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
      })
        .formatToParts(1)
        .find(part => part.type === 'currency')?.value || currencyCode;
    } catch {
      return currencyCode;
    }
  }

  return (
    <Stack
      tokens={{ childrenGap: 20 }}
      styles={{
        root: {
          // maxWidth: 900,
          // margin: "auto",
          padding: 10,
          backgroundColor: "#fff",
          borderRadius: 10,
          boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
        },
      }}
    >
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
        <div style={{ width: 6, height: 48, backgroundColor: primaryColor, borderRadius: 2 }} />
        <Icon iconName="PageDetails" styles={{ root: { fontSize: 36, color: primaryColor } }} />
        <Text variant="xxLarge" styles={{ root: { fontWeight: "600", color: primaryColor } }}>
          Invoice Details
        </Text>
      </Stack>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
          gap: "16px 20px",
          marginTop: 10,
        }}
      >

        {/* Reusable info cell */}
        {[
          { label: "Purchase Order:", value: item.PurchaseOrder || "-" },
          { label: "Project Name:", value: item.ProjectName || "-" },
          {
            label: "PO Item Title:",
            value: item["POItem_x0020_Title"] || "-",
          },
          {
            label: "PO Item Value:",
            value: item["POItem_x0020_Value"]
              ? `${itemCurrency} ${Number(item["POItem_x0020_Value"]).toLocaleString()}`
              : "-",
          },
          {
            label: "Invoiced Amount:",
            value: item.InvoiceAmount
              ? `${itemCurrency} ${item.InvoiceAmount.toLocaleString()}`
              : "-",
          },
          {
            label: "Invoice Status:",
            value: (
              <span
                style={{
                  fontWeight: 600,
                }}
              >
                {item.Status || "-"}
              </span>
            ),
          },
          {
            label: "Current Status:",
            value: item.CurrentStatus || "-",
          },
          {
            label: "Due Date:",
            value: item.DueDate ? new Date(item.DueDate).toLocaleDateString() : "N/A",
          },
        ].map((field, idx) => (
          <div
            key={idx}
            style={{
              display: "grid",
              gridTemplateColumns: "110px 1fr",
              padding: "10px 14px",
              background: "#fafafa",
              border: "1px solid #eee",
              borderRadius: 8,
              boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
              alignItems: "center",
            }}
          >
            <span style={{ fontSize: 13, fontWeight: 600, color: primaryColor }}>{field.label}</span>
            <span style={{ fontSize: 14, wordBreak: "break-word" }}>{field.value}</span>
          </div>
        ))}
      </div>

      <Separator styles={{ root: { marginTop: 5, marginBottom: 5 } }} />

      {
        item.PMCommentsHistory && formatCommentsHistory(item.PMCommentsHistory).trim() !== "" && (
          <Stack>
            <Text variant="mediumPlus" styles={{ root: { fontSize: 13, fontWeight: 600, color: primaryColor } }}>Requestor Comments</Text>
            <div style={{
              maxHeight: 180,
              overflowY: "auto",
              backgroundColor: "#f3f2f1",
              borderRadius: 6,
              padding: 12,
            }}>
              <pre style={{
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
                margin: 0,
                fontSize: 14,
                fontFamily: "Segoe UI",
                color: "#333",
              }}>
                {formatCommentsHistory(item.PMCommentsHistory)}
              </pre>
            </div>
          </Stack>
        )
      }

      {
        item.FinanceCommentsHistory && formatCommentsHistory(item.FinanceCommentsHistory).trim() !== "" && (
          <Stack>
            <Text variant="mediumPlus" styles={{ root: { fontSize: 13, fontWeight: 600, color: primaryColor } }}>Finance Comments</Text>
            <div style={{
              maxHeight: 180,
              overflowY: "auto",
              backgroundColor: "#f3f2f1",
              borderRadius: 6,
              padding: 12,
            }}>
              <pre style={{
                whiteSpace: "pre-wrap",
                wordBreak: "break-word",
                margin: 0,
                fontSize: 14,
                fontFamily: "Segoe UI",
                color: "#333",
              }}>
                {formatCommentsHistory(item.FinanceCommentsHistory)}
              </pre>
            </div>
          </Stack>
        )
      }

      {/* Attachments */}
      {item.AttachmentFiles && item.AttachmentFiles.length > 0 && (
        <Stack>
          <Text variant="mediumPlus" block styles={{ root: { marginTop: 16, fontWeight: '600' } }}>
            Attachments
          </Text>
          <ul style={{ paddingLeft: 20, marginTop: 8 }}>
            {item.AttachmentFiles.map((file) => (
              <li key={file.UniqueId} style={{ marginBottom: 6 }}>
                <a
                  href="#"
                  onClick={(e) => {
                    e.preventDefault();
                    onShowAttachment(file.ServerRelativeUrl, file.Name || file.FileName);
                  }}
                  style={{ color: "#0078d4", textDecoration: "underline" }}
                >
                  {file.Name || file.FileName}
                </a>
                <IconButton
                  iconProps={{ iconName: "Download" }}
                  title="Download attachment"
                  ariaLabel="Download attachment"
                  onClick={() => {
                    const absoluteUrl = file.ServerRelativeUrl.startsWith("http")
                      ? file.ServerRelativeUrl
                      : `${window.location.origin}${file.ServerRelativeUrl}`;
                    const link = document.createElement("a");
                    link.href = absoluteUrl;
                    link.download = file.Name || file.FileName || "attachment";
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                  }}
                  styles={{
                    root: { marginLeft: 8, color: "#0078d4" },
                    rootHovered: { background: "#f3f2f1" },
                  }}
                />
              </li>
            ))}
          </ul>
        </Stack>
      )}

      <div style={{ marginTop: 24 }}>
        <StatusStepper currentStatus={item.Status ?? ""} steps={steps} />
      </div>
    </Stack>
  );
}

export default function MyRequests({ sp, projectsp, context, initialFilters, getCurrentPageUrl }: MyProps) {
  const [invoicePOs, setInvoicePOs] = useState<InvoicePO[]>([]);
  const [invoiceRequests, setInvoiceRequests] = useState<InvoiceRequest[]>([]);
  const [poHierarchy, setPOHierarchy] = useState<null | {
    mainPO: InvoicePO;
    lineItemGroups: { poItem: any; requests: InvoiceRequest[] }[];
    childPOGroups: { childPO: InvoicePO; requests: InvoiceRequest[] }[];
    mainPORequests: InvoiceRequest[];
  }>(null);
  const [selectedReq, setSelectedReq] = useState<InvoiceRequest | null>(null);
  const [selectedPOItem, setSelectedPOItem] = useState<{ POID: string; POAmount: string; Currency: string } | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [showClarifyPanel, setShowClarifyPanel] = useState(false);
  const [clarifyInvoiceAmount, setClarifyInvoiceAmount] = useState<number | undefined>();
  const [clarifyCustomerContact, setClarifyCustomerContact] = useState<string | undefined>();
  const [clarifyComment, setClarifyComment] = useState("");
  const [clarifyLoading, setClarifyLoading] = useState(false);
  const [searchText, setSearchText] = useState("");
  // const [filterProjectName, setFilterProjectName] = useState<string | undefined>(undefined);
  const [filterProjectName, setFilterProjectName] = useState<string[]>(["All"]);
  const [, setProjectOptions] = useState<string[]>([]);
  const [, setStatusOptions] = useState<string[]>([]);
  const [dialogVisible, setDialogVisible] = useState(false);
  const [dialogMessage, setDialogMessage] = useState("");
  const [dialogType, setDialogType] = useState<"success" | "error">("success");
  const [selectedProject, setSelectedProject] = useState<any | null>(null);
  const [showHierPanel, setShowHierPanel] = useState(false); // main panel
  const [viewerUrl, setViewerUrl] = useState<string | null>(null); // viewer panel
  const [viewerName, setViewerName] = useState<string | null>(null);
  const [sortedFilteredItems, setSortedFilteredItems] = React.useState<any[]>([]);
  // const [filterCurrentStatus, setFilterCurrentStatus] = useState<string | undefined>(undefined); // uses CurrentStatus field
  // const [filterInvoiceStatus, setFilterInvoiceStatus] = useState<string | undefined>(undefined);   // uses Status field
  // const [filterFinanceStatus, setFilterFinanceStatus] = useState<string | undefined>(undefined);
  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
  const [isColumnPanelOpen, setIsColumnPanelOpen] = useState(false);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState(false);
  const [currentFilterColumn, setCurrentFilterColumn] = useState<string>('');
  const [columnFilterMenu, setColumnFilterMenu] = useState<{ visible: boolean; target: HTMLElement | null; columnKey: string | null }>({
    visible: false,
    target: null,
    columnKey: null,
  });
  const [isAdmin, setIsAdmin] = useState<boolean>(false);
  const [sortedColumnKey,] = React.useState<string | null>(null);
  const [isSortedDescending,] = React.useState<boolean>(false);
  const [isInvoiceRequestViewPanelOpen, setIsInvoiceRequestViewPanelOpen] = useState(false);
  const [selectedInvoiceRequest, setSelectedInvoiceRequest] = useState<InvoiceRequest | null>(null);
  const [projectSearch,] = React.useState<string>('')
  // const [invoicePercentStatusFilter, setInvoicePercentStatusFilter] = React.useState<string | null>(null);
  const invoicePercentStatusOptions: IDropdownOption[] = [
    { key: "All", text: "All" },
    { key: "NotPaid", text: "Not Paid" },
    { key: "PartiallyInvoiced", text: "Partially Invoiced" },
    { key: "CompletelyInvoiced", text: "Completely Invoiced" },
  ];

  const INVOICE_STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'All', text: 'All' },
    { key: 'Invoice Requested', text: 'Invoice Requested' },
    { key: 'Invoice Raised', text: 'Invoice Raised' },
    { key: 'Pending Payment', text: 'Pending Payment' },
    { key: 'Payment Received', text: 'Payment Received' },
    { key: 'Cancelled', text: 'Cancelled' }
  ];

  const CURRENT_STATUS_OPTIONS: IDropdownOption[] = [
    { key: 'All', text: 'All' },
    { key: 'Request Submitted', text: 'Request Submitted' },
    { key: 'Pending Finance', text: 'Pending Finance' },
    { key: 'Finance asked Clarification', text: 'Finance asked Clarification' },
    { key: 'Completed', text: 'Completed' },
    { key: 'Cancelled Request', text: 'Cancelled Request' }
  ];

  const [filterCurrentStatus, setFilterCurrentStatus] = useState<string[]>(["All"]);
  const [filterInvoiceStatus, setFilterInvoiceStatus] = useState<string[]>(["All"]);
  const [filterFinanceStatus, setFilterFinanceStatus] = useState<string | undefined>("All");
  const [invoicePercentStatusFilter, setInvoicePercentStatusFilter] = useState<string[]>(["All"]);
  const onInvoiceRequestClicked = (item: InvoiceRequest) => {
    setSelectedInvoiceRequest(item);
    setIsInvoiceRequestViewPanelOpen(true);
  };
  const getColumnDistinctValues = (columnKey: string): string[] => {
    const col = columns.find(c => c.key === columnKey);
    if (!col || !col.fieldName) return [];
    const field = col.fieldName;
    const values = Array.from(
      new Set(
        invoiceRequests
          .map(i => (i as any)[field])
          .filter(v => v !== null && v !== undefined && v !== '')
          .map(v => v.toString())
      )
    );
    return values.sort((a, b) => a.localeCompare(b));
  };

  const clearColumnFilter = (columnKey: string) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      delete next[columnKey];
      return next;
    });
    setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  };

  const getVisibleColumns = (): IColumn[] =>
    columns
      .filter(col => visibleColumns.includes(col.key as string));

  const moveColumn = (columnKey: string, direction: 'up' | 'down') => {
    const currentIndex = visibleColumns.indexOf(columnKey)
    if (direction === 'up' && currentIndex > 0) {
      const newOrder = [...visibleColumns]
        ;[newOrder[currentIndex - 1], newOrder[currentIndex]] = [newOrder[currentIndex], newOrder[currentIndex - 1]]
      setVisibleColumns(newOrder)
    } else if (direction === 'down' && currentIndex < visibleColumns.length - 1) {
      const newOrder = [...visibleColumns]
        ;[newOrder[currentIndex], newOrder[currentIndex + 1]] = [newOrder[currentIndex + 1], newOrder[currentIndex]]
      setVisibleColumns(newOrder)
    }
  }

  const toggleColumnVisibility = (columnKey: string) => {
    setVisibleColumns(prev =>
      prev.includes(columnKey) ? prev.filter(k => k !== columnKey) : [...prev, columnKey]
    )
  }
  // const menuItems = [
  //   {
  //     key: 'asc',
  //     text: 'Sort A to Z',
  //     iconProps: { iconName: 'SortUp' },
  //     onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc'),
  //   },
  //   {
  //     key: 'desc',
  //     text: 'Sort Z to A',
  //     iconProps: { iconName: 'SortDown' },
  //     onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc'),
  //   },
  //   { key: 'divider', itemType: ContextualMenuItemType.Divider },
  // ];

  // const [columnFilterMenu, setColumnFilterMenu] = useState<{ visible: boolean; target: HTMLElement | null; columnKey: string | null }>({
  //   visible: false,
  //   target: null,
  //   columnKey: null,
  // });
  const menuItems = [
    {
      key: 'asc', text: 'Sort A→Z', iconProps: { iconName: 'SortUp' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'asc')
    },
    {
      key: 'desc', text: 'Sort Z→A', iconProps: { iconName: 'SortDown' },
      onClick: () => sortColumn(columnFilterMenu.columnKey!, 'desc')
    },
    { key: 'divider1', itemType: ContextualMenuItemType.Divider },
    {
      key: 'filter', text: 'Filter Column', iconProps: { iconName: 'Filter' },
      onClick: () => {
        if (!columnFilterMenu.columnKey) return;
        setCurrentFilterColumn(columnFilterMenu.columnKey);
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

  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 200 },
    callout: { minWidth: 200 },
    dropdownItem: {
      whiteSpace: 'normal',
      textOverflow: 'clip',
      overflow: 'visible',
      maxWidth: 'none'
    },
    dropdownItemSelected: {
      whiteSpace: 'normal',
      textOverflow: 'clip',
      overflow: 'visible',
      maxWidth: 'none'
    }
  };

  // const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
  //   const sortedItems = [...sortedFilteredItems].sort((a, b) => {

  //     let aVal = (a as any)[columnKey];
  //     let bVal = (b as any)[columnKey];


  //     if (aVal == null && bVal == null) return 0;
  //     if (aVal == null) return 1;
  //     if (bVal == null) return -1;


  //     if (aVal instanceof Date) aVal = aVal.getTime();
  //     if (bVal instanceof Date) bVal = bVal.getTime();


  //     if (typeof aVal === 'number' && typeof bVal === 'number') {
  //       return direction === 'asc' ? aVal - bVal : bVal - aVal;
  //     }

  //     const aAsDate = Date.parse(aVal);
  //     const bAsDate = Date.parse(bVal);
  //     if (!isNaN(aAsDate) && !isNaN(bAsDate)) {
  //       return direction === 'asc' ? aAsDate - bAsDate : bAsDate - aAsDate;
  //     }

  //     const aStr = aVal.toString();
  //     const bStr = bVal.toString();
  //     return direction === 'asc'
  //       ? aStr.localeCompare(bStr)
  //       : bStr.localeCompare(aStr);
  //   });

  //   setSortedFilteredItems(sortedItems);
  //   setColumnFilterMenu({ visible: false, target: null, columnKey: null });
  // };

  const sortColumn = (columnKey: string, direction: 'asc' | 'desc') => {
    const isAmountField = ['POItemx0020Value', 'InvoiceAmount'].includes(columnKey)

    const sortedItems = [...sortedFilteredItems.length ? sortedFilteredItems : filteredItems].sort((a: any, b: any) => {
      let aVal = a[columnKey]
      let bVal = b[columnKey]

      // EMPTY/NULL FIRST in ASC (0 first for numbers)
      if (aVal === null || aVal === undefined || aVal === '') {
        return direction === 'asc' ? -1 : 1  // empties first in asc, last in desc
      }
      if (bVal === null || bVal === undefined || bVal === '') {
        return direction === 'asc' ? 1 : -1   // empties first in asc, last in desc
      }

      // NUMERIC FIELDS (POItemValue, InvoiceAmount) - 0 first in ASC
      if (isAmountField) {
        const aNum = Number(aVal) || 0
        const bNum = Number(bVal) || 0
        return direction === 'asc' ? aNum - bNum : bNum - aNum
      }

      // DATES
      if (aVal instanceof Date) {
        aVal = aVal.getTime()
      }
      if (bVal instanceof Date) {
        bVal = bVal.getTime()
      }
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

    setSortedFilteredItems(sortedItems)
    setColumnFilterMenu({ visible: false, target: null, columnKey: null })
  }

  const onColumnHeaderClick = (
    ev?: React.MouseEvent<HTMLElement>,
    column?: IColumn
  ) => {
    if (!column || !ev?.currentTarget) return;

    setColumnFilterMenu({
      visible: true,
      target: ev.currentTarget as HTMLElement,
      columnKey: column.key,
    });
  };

  const handleExportToExcel = (): void => {
    const exportSource = sortedFilteredItems && sortedFilteredItems.length > 0
      ? sortedFilteredItems
      : filteredItems || invoiceRequests;

    const exportData = exportSource.map((item) => {
      const currencySymbol = getCurrencySymbol(item.Currency);
      const invoicePercent = calculateInvoicedPercentForPO(item.PurchaseOrder, invoiceRequests).toFixed(0) + "%";
      const poItemInvoicePercent = calculateInvoicedPercentForPOItem(
        item.PurchaseOrder,
        item.POItem_x0020_Title,
        item.POItem_x0020_Value,
        invoiceRequests
      ).toFixed(0) + "%";
      const createddate = new Date(item.Created).toLocaleDateString()
      const modifieddate = new Date(item.Modified).toLocaleDateString()

      return {
        POID: item.PurchaseOrder,
        Project: item.ProjectName,
        CurrentStatus: item.CurrentStatus,
        InvoiceStatus: item.Status,
        POItemTitle: item.POItem_x0020_Title,
        POItemValue: `${currencySymbol}${Number(item.POItem_x0020_Value || 0).toLocaleString()}`,
        "Invoiced Amount": `${currencySymbol}${Number(item.InvoiceAmount || 0).toLocaleString()}`,
        "Invoice %": invoicePercent,
        "PO Item Invoice %": poItemInvoicePercent,
        Created: createddate,
        "Created By": item.Author?.Title,
        Modified: modifieddate,
        "Modified By": item.Editor?.Title,
      };
    });

    // Convert JSON to worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);

    // Create a new workbook and append the worksheet
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'InvoiceRequests');

    // Write workbook and convert to binary
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    // Save file using file-saver
    const fileName = `InvoiceRequests_${new Date().toISOString()}.xlsx`;
    saveAs(new Blob([wbout], { type: 'application/octet-stream' }), fileName);
  };

  // Columns for Invoice requests list:
  const invoiceColumns: IColumn[] = [
    { key: "PurchaseOrder", name: "Purchase Order", fieldName: "PurchaseOrder", minWidth: 110, isCollapsible: true, isResizable: true, onColumnClick: onColumnHeaderClick, },
    { key: "ProjectName", name: "Project", fieldName: "ProjectName", minWidth: 130, isCollapsible: true, isResizable: true, onColumnClick: onColumnHeaderClick },
    {
      key: "CurrentStatus",
      name: "Current Status",
      fieldName: "CurrentStatus",
      minWidth: 120,
      isCollapsible: true,
      isResizable: true,
      onRender: (item) => item.CurrentStatus || "-",
      onColumnClick: onColumnHeaderClick
    },
    { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 120, isCollapsible: true, isResizable: true, onColumnClick: onColumnHeaderClick },
    // { key: "CurrentStatus", name: "Current Status", fieldName: "CurrentStatus", minWidth: 120, isCollapsible: true, isResizable: true, onColumnClick: onColumnHeaderClick },
    {
      key: "POItem_x0020_Title",
      name: "PO Item Title",
      fieldName: "POItem_x0020_Title",
      minWidth: 130,
      isCollapsible: true,
      isResizable: true,
      onRender: item => item["POItem_x0020_Title"] || "-",
      onColumnClick: onColumnHeaderClick
    },
    {
      key: "POItem_x0020_Value",
      name: "PO Item Value",
      fieldName: "POItem_x0020_Value",
      minWidth: 130,
      isCollapsible: true,
      isResizable: true,
      onRender: (item) => {
        // item["POItem_x0020_Value"] != null && !isNaN(Number(item["POItem_x0020_Value"]))
        //   ? ` ${Number(item["POItem_x0020_Value"]).toLocaleString()} ${item.Currency ?? ''}`.trim()
        //   : "-",
        const symbol = getCurrencySymbol(item.Currency);
        const value = item.POItem_x0020_Value ?? '';
        return <span>{symbol} {Number(value).toLocaleString()}</span>;
      },
      onColumnClick: onColumnHeaderClick
    },
    // {
    //   key: "InvoiceAmount",
    //   name: "Invoiced Amount",
    //   fieldName: "InvoiceAmount",
    //   minWidth: 150,
    //   isCollapsible: true,
    //   isResizable: true,
    //   onRender: (item) => {
    //     const symbol = getCurrencySymbol(item.Currency);
    //     const value = item.InvoiceAmount ?? '';
    //     return <span>{symbol} {Number(value).toLocaleString()}</span>;
    //   },
    //   // }`${item.InvoiceAmount.toLocaleString()} ${item.Currency ?? ''}`.trim(),
    //   onColumnClick: onColumnHeaderClick
    // },
    {
      key: "RequestedAmount",
      name: "Requested Amount",
      minWidth: 130,
      isCollapsible: true,
      isResizable: true,
      onColumnClick: onColumnHeaderClick,
      onRender: (item: InvoiceRequest) => {
        // Sum InvoiceAmount where Status = "Request Submitted" for this PO
        const poRequests = invoiceRequests.filter(ir =>
          ir.PurchaseOrder === item.PurchaseOrder &&
          ir.Status === "Request Submitted"
        );
        const requestedAmount = poRequests.reduce((sum, ir) => sum + (ir.InvoiceAmount ?? 0), 0);
        const symbol = getCurrencySymbol(item.Currency);
        return <span>{symbol}{requestedAmount.toLocaleString()}</span>;
      }
    },

    // ✅ INVOICED AMOUNT (Invoice Raised + Pending Payment)
    {
      key: "InvoicedAmount",
      name: "Invoiced Amount",
      fieldName: "InvoiceAmount",
      minWidth: 130,
      isCollapsible: true,
      isResizable: true,
      onColumnClick: onColumnHeaderClick,
      onRender: (item: InvoiceRequest) => {
        // Sum InvoiceAmount where Status = "Invoice Raised" OR "Pending Payment" for this PO
        const poInvoices = invoiceRequests.filter(ir =>
          ir.PurchaseOrder === item.PurchaseOrder &&
          (ir.Status === "Invoice Raised" || ir.Status === "Pending Payment")
        );
        const invoicedAmount = poInvoices.reduce((sum, ir) => sum + (ir.InvoiceAmount ?? 0), 0);
        const symbol = getCurrencySymbol(item.Currency);
        return <span>{symbol}{invoicedAmount.toLocaleString()}</span>;
      }
    },

    // ✅ PAID AMOUNT (Payment Received)
    {
      key: "PaidAmount",
      name: "Paid Amount",
      minWidth: 100,
      isCollapsible: true,
      isResizable: true,
      onColumnClick: onColumnHeaderClick,
      onRender: (item: InvoiceRequest) => {
        // Sum InvoiceAmount where Status = "Payment Received" for this PO
        const poPayments = invoiceRequests.filter(ir =>
          ir.PurchaseOrder === item.PurchaseOrder &&
          ir.Status === "Payment Received"
        );
        const paidAmount = poPayments.reduce((sum, ir) => sum + (ir.InvoiceAmount ?? 0), 0);
        const symbol = getCurrencySymbol(item.Currency);
        return <span>{symbol}{paidAmount.toLocaleString()}</span>;
      }
    },
    {
      key: 'InvoicedPercent',
      name: 'Invoiced %',
      fieldName: 'InvoicedPercent',
      minWidth: 100,
      isCollapsible: true,
      isResizable: true,
      onRender: (item: InvoiceRequest) => {
        // const poAmountRaw = invoicePOs.find(po => po.POID === item.PurchaseOrder)?.POAmount || "0"; // string | "0"
        // const poAmount = parseFloat(poAmountRaw); // number

        const percent = calculateInvoicedPercentForPO(
          item.PurchaseOrder,      // mainPOID: string
          invoiceRequests          // invoiceRequests: InvoiceRequest[]
        );

        return `${percent.toFixed(0)}%`;
      }
    },
    {
      key: 'POItemInvoicedPercent',
      name: 'PO Item Invoiced %',
      fieldName: 'POItemInvoicedPercent',
      minWidth: 120,
      isCollapsible: true,
      isResizable: true,
      // onRender: item => item.POItemInvoicedPercent?.toFixed(0)
      onRender: (item: InvoiceRequest) => {
        const poItemAmount = item.POItem_x0020_Value || 0;
        return `${calculateInvoicedPercentForPOItem(item.PurchaseOrder, item.POItem_x0020_Title || '', parseFloat(poItemAmount.toString()), invoiceRequests).toFixed(0)}%`;
      }
    },
    {
      key: "Created",
      name: "Created",
      fieldName: "Created",
      minWidth: 110,
      isCollapsible: true,
      isResizable: true,
      onRender: item => item.Created ? new Date(item.Created).toLocaleDateString() : "-",
      onColumnClick: onColumnHeaderClick
    },
    {
      key: "Author",
      name: "Created By",
      fieldName: "Author",
      minWidth: 110,
      isCollapsible: true,
      isResizable: true,
      onRender: item => item.Author?.Title || "-",
      onColumnClick: onColumnHeaderClick
    },
    {
      key: "Modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 110,
      isCollapsible: true,
      isResizable: true,
      onRender: item => item.Modified ? new Date(item.Modified).toLocaleDateString() : "-",
      onColumnClick: onColumnHeaderClick
    },
    {
      key: "Editor",
      name: "Modified By",
      fieldName: "Editor",
      minWidth: 110,
      isCollapsible: true,
      isResizable: true,
      onRender: item => item.Editor?.Title || "-",
      onColumnClick: onColumnHeaderClick
    },

  ];

  const [columns,] = useState<IColumn[]>(invoiceColumns);
  const [visibleColumns, setVisibleColumns] = useState<string[]>(invoiceColumns.map(c => c.key as string));
  // Columns for PO items:
  const poColumns: IColumn[] = [
    { key: "POID", name: "POItem Title", fieldName: "POID", minWidth: 150, maxWidth: 220, isResizable: true },
    {
      key: "POAmount", name: "POItem Amount", fieldName: "POAmount", minWidth: 140, maxWidth: 160, isResizable: true, onRender: (item) => {
        const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
        const value = item.POItem_x0020_Value ?? 0;
        return <span>{symbol} {Number(value).toLocaleString()}</span>;
      }
    },
  ];

  const poColumnsLine: IColumn[] = [
    { key: "POItem_x0020_Title", name: "POItem Title", fieldName: "POItem_x0020_Title", minWidth: 150, maxWidth: 220, isResizable: true },
    {
      key: "POItem_x0020_Value", name: "POItem Amount", fieldName: "POItem_x0020_Value", minWidth: 140, maxWidth: 160, isResizable: true, onRender: (item) => {
        if (!selectedReq || !selectedReq.PurchaseOrder) {
          // fallback when no request selected yet
          const value = item.POItem_x0020_Value ?? 0;
          return <span>{Number(value).toLocaleString()}</span>;
        }
        const currencyCode = getCurrencyByPOID(selectedReq.PurchaseOrder, invoicePOs);
        const symbol = getCurrencySymbol(currencyCode);
        const value = item.POItem_x0020_Value ?? 0;
        return <span>{symbol} {Number(value).toLocaleString()}</span>;
      }
    },
    // { key: "Comments", name: "Description", fieldName: "Comments", minWidth: 170, maxWidth: 270, isResizable: true }, // Optional
  ];

  // Columns for invoice requests grouped by PO:

  // const enrichedItems = sortedFilteredItems.map(po => {

  //   return {
  //     ...po,
  //     InvoicedPercent: calculateInvoicedPercentForPO(po.PurchaseOrder, invoiceRequests),
  //     POItemInvoicedPercent: calculateInvoicedPercentForPOItem(po.PurchaseOrder, po.POItem_x0020_Title, po.POItem_x0020_Value, invoiceRequests),
  //   };
  // });

  const groupedInvColumns: IColumn[] = [
    { key: "POItem_x0020_Title", name: "PO Item Title", fieldName: "POItem_x0020_Title", minWidth: 150, maxWidth: 220, isResizable: true },
    {
      key: "POItem_x0020_Value", name: "PO Item Value", fieldName: "POItem_x0020_Value", minWidth: 140, maxWidth: 160, isResizable: true, onRender: (item) => {
        const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
        const value = item.POItem_x0020_Value ?? 0;
        return <span>{symbol} {Number(value).toLocaleString()}</span>;
      }
    },
    {
      key: "RequestedAmount",
      name: "Requested Amount",
      minWidth: 110,
      onRender: (item: any) => {
        const { requested } = getAmountBuckets(item);
        if (!requested) return null;
        const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
        return <span>{symbol} {requested.toLocaleString()}</span>;
      },
    },
    {
      key: "InvoicedAmount",
      name: "Invoiced Amount",
      minWidth: 110,
      onRender: (item: any) => {
        const { invoiced } = getAmountBuckets(item);
        if (!invoiced) return null;
        const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
        return <span>{symbol} {invoiced.toLocaleString()}</span>;
      },
    },
    {
      key: "PaidAmount",
      name: "Paid Amount",
      minWidth: 110,
      onRender: (item: any) => {
        const { paid } = getAmountBuckets(item);
        if (!paid) return null;
        const symbol = item.Currency ? getCurrencySymbol(item.Currency) : "";
        return <span>{symbol} {paid.toLocaleString()}</span>;
      },
    },
    { key: "Status", name: "Invoice Status", fieldName: "Status", minWidth: 120, maxWidth: 160, isResizable: true },
    {
      key: "Created",
      name: "Created",
      fieldName: "Created",
      minWidth: 130,
      maxWidth: 180,
      isResizable: true,
      onRender: item => item.Created ? new Date(item.Created).toLocaleDateString() : "-"
    },
    {
      key: "Author",
      name: "Created By",
      fieldName: "Author",
      minWidth: 160,
      maxWidth: 200,
      isResizable: true,
      onRender: item => item.Author?.Title || "-"
    },
    {
      key: "Modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 130,
      maxWidth: 180,
      isResizable: true,
      onRender: item => item.Modified ? new Date(item.Modified).toLocaleDateString() : "-"
    },
    {
      key: "Editor",
      name: "Modified By",
      fieldName: "Editor",
      minWidth: 160,
      maxWidth: 200,
      isResizable: true,
      onRender: item => item.Editor?.Title || "-"
    },
    {
      key: "PMCommentsHistory",
      name: "Requestor Comments",
      fieldName: "PMCommentsHistory",
      minWidth: 200,
      maxWidth: 350,
      isResizable: true,
      onRender: item => formatCommentsHistory(item.PMCommentsHistory)
    },
    {
      key: "FinanceCommentsHistory",
      name: "Finance Comments",
      fieldName: "FinanceCommentsHistory",
      minWidth: 200,
      maxWidth: 350,
      isResizable: true,
      onRender: item => formatCommentsHistory(item.FinanceCommentsHistory)
    },

  ];
  // Helper to render text or fallback
  const renderValue = (value: any) => value ? value : <span style={{ color: '#999' }}>—</span>;

  const [selection] = useState(
    new Selection({
      onSelectionChanged: () => {
        const selected = selection.getSelection()[0] as InvoiceRequest | undefined;
        onInvoiceRequestSelect(selected);
      }
    })
  );

  const [clearCounter, setClearCounter] = useState(0);
  // const projectOptions = React.useMemo(() => {
  //   return Array.from(new Set(
  //     invoiceRequests
  //       .filter(item => {
  //         const matchesCurrentStatus = !filterCurrentStatus || filterCurrentStatus === "All"
  //           ? true : item.CurrentStatus === filterCurrentStatus;
  //         const matchesInvoiceStatus = !filterInvoiceStatus || filterInvoiceStatus === "All"
  //           ? true : item.Status === filterInvoiceStatus;
  //         return matchesCurrentStatus && matchesInvoiceStatus;
  //       })
  //       .map(item => item.ProjectName)
  //       .filter(Boolean)
  //   ));
  // }, [invoiceRequests, filterCurrentStatus, filterInvoiceStatus]);
  const projectOptions: IComboBoxOption[] = React.useMemo(() =>
    Array.from(
      new Set(
        invoiceRequests
          .filter(item => {
            const matchesCurrentStatus =
              !filterCurrentStatus.length ||
              filterCurrentStatus.includes("All") ||
              filterCurrentStatus.includes(item.CurrentStatus || "");
            const matchesInvoiceStatus =
              !filterInvoiceStatus.length ||
              filterInvoiceStatus.includes("All") ||
              filterInvoiceStatus.includes(item.Status || "");
            return matchesCurrentStatus && matchesInvoiceStatus;
          })
          .map(item => item.ProjectName)
          .filter(Boolean)
      )
    )
      .sort((a, b) => a.localeCompare(b)) // ASC
      .map(proj => ({ key: proj, text: proj })),
    [invoiceRequests, filterCurrentStatus, filterInvoiceStatus]
  );

  const projectMultiOptions: IDropdownOption[] = React.useMemo(
    () =>
      projectOptions.map(p => ({
        key: p.key,              // use the key value, not the whole object
        text: p.text as string,  // ComboBox text is string | undefined
      })),
    [projectOptions]
  )

  const filteredProjectOptions: IDropdownOption[] = React.useMemo(
    () => [
      { key: 'All', text: 'All' },
      ...projectMultiOptions.filter(o =>
        (o.text || '').toLowerCase().includes(projectSearch.toLowerCase())
      ),
    ],
    [projectMultiOptions, projectSearch]
  )

  // const renderProjectTitle: IRenderFunction<IDropdownOption[]> = (options, defaultRender) => (
  //   <Stack tokens={{ childrenGap: 4 }}>
  //     <TextField
  //       placeholder="Search project"
  //       value={projectSearch}
  //       onChange={(_, v) => setProjectSearch(v || '')}
  //       styles={{ root: { padding: '4px 8px 0 8px' } }}
  //       borderless
  //     />
  //     {defaultRender && defaultRender(options)}
  //   </Stack>
  // )
  const clearAllFilters = () => {
    setSearchText("");
    setFilterProjectName(["All"]);
    setFilterCurrentStatus(["All"]);
    setFilterInvoiceStatus(["All"]);
    setFilterFinanceStatus("All");
    setInvoicePercentStatusFilter(["All"]);
    setClearCounter(clearCounter + 1);
  };


  // const isClearDisabled =
  //   !searchText &&
  //   (filterProjectName === "All" || !filterProjectName) &&
  //   (filterInvoiceStatus === "All" || !filterInvoiceStatus) &&
  //   (filterCurrentStatus === "All" || !filterCurrentStatus) &&
  //   (filterFinanceStatus === "All" || !filterFinanceStatus) &&
  //   (invoicePercentStatusFilter === null || invoicePercentStatusFilter === "All");

  const isClearDisabled =
    !searchText &&
    (filterProjectName.includes("All") || !filterProjectName) &&
    (!filterInvoiceStatus || !filterInvoiceStatus.length || filterInvoiceStatus.includes("All")) &&
    (!filterCurrentStatus || !filterCurrentStatus.length || filterCurrentStatus.includes("All")) &&
    (filterFinanceStatus === "All" || !filterFinanceStatus) &&
    (!invoicePercentStatusFilter || !invoicePercentStatusFilter.length || invoicePercentStatusFilter.includes("All"));

  // useEffect(() => {
  //   async function loadData() {
  //     setLoading(true);
  //     try {
  //       const [pos, reqs] = await Promise.all([
  //         sp.web.lists.getByTitle("InvoicePO").items(),
  //         sp.web.lists.getByTitle("Invoice Requests").items.select("*", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail").expand("AttachmentFiles", "Author", "Editor")(),
  //       ]);
  //       setInvoicePOs(pos);
  //       setInvoiceRequests(reqs);
  //       setProjectOptions(Array.from(new Set(reqs.map(r => r.ProjectName).filter(Boolean))));
  //       setStatusOptions(Array.from(new Set(reqs.map(r => r.Status).filter(Boolean))));
  //       setError(null);
  //     } catch (err: any) {
  //       setError(`Error loading data: ${err.message || err}`);
  //     } finally {
  //       setLoading(false);
  //     }
  //   }
  //   loadData();
  // }, [sp]);

  useEffect(() => {
    async function loadRole() {
      try {
        const admin = await isCurrentUserAdmin(sp, context);
        setIsAdmin(admin);
      } catch (e) {
        console.error("Failed to resolve admin role", e);
        setIsAdmin(false);
      }
    }
    loadRole();
  }, [sp, context]);

  useEffect(() => {
    async function getAllItemsPaged(sp: any, listTitle: string, select?: string[], expand?: string[]) {
      let allItems: any[] = [];
      let pageSize = 2000;
      let query = sp.web.lists.getByTitle(listTitle).items.top(pageSize);
      if (select) query = query.select(...select);
      if (expand) query = query.expand(...expand);
      let items = await query();
      while (items && items.length > 0) {
        allItems.push(...items);
        if (items['@odata.nextLink']) {
          items = await sp.web.get(items['@odata.nextLink'])();
        } else {
          break;
        }
      }
      return allItems;
    }

    async function loadData() {
      setLoading(true);
      try {
        const pos = await getAllItemsPaged(sp, "InvoicePO");
        const reqs = await getAllItemsPaged(sp, "Invoice Requests",
          ["*", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail"],
          ["Author", "Editor", "AttachmentFiles"]
        );
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
      // if (initialFilters.projectName !== undefined) setFilterProjectName(initialFilters.projectName);
      // if (initialFilters.Status !== undefined) setFilterInvoiceStatus(initialFilters.Status);
      if (initialFilters.projectName !== undefined) {
        setFilterProjectName([initialFilters.projectName]);
      }
      if (initialFilters.Status !== undefined) {
        setFilterInvoiceStatus([initialFilters.Status]);
      }
      if (initialFilters.FinanceStatus !== undefined) setFilterFinanceStatus(initialFilters.FinanceStatus);
      if (initialFilters.CurrentStatus !== undefined) setFilterCurrentStatus(initialFilters.CurrentStatus);
    }
    // empty dep array if initialFilters won't change after mount
  }, [initialFilters]);

  useEffect(() => {
    console.log("Filters set from initialFilters:", {
      searchText,
      filterFinanceStatus,
      filterProjectName,
      filterInvoiceStatus,
      filterCurrentStatus,
    });
  }, [searchText, filterProjectName, filterInvoiceStatus, filterCurrentStatus, filterFinanceStatus]);

  useEffect(() => {
    if (selectedReq?.ProjectName) {
      loadProject(selectedReq.ProjectName);
    } else {
      setSelectedProject(null);
    }
  }, [selectedReq]);

  useEffect(() => {
    async function loadSelectedInvoiceAndHierarchy() {
      if (initialFilters?.selectedInvoice) {
        const invoiceId = Number(initialFilters.selectedInvoice);
        if (invoiceId) {
          const invoice = await fetchInvoiceRequestById(sp, invoiceId);
          if (invoice) {
            setSelectedReq(invoice);
            setShowHierPanel(true);
            const mainPO = findMainPO(invoice, invoicePOs);
            if (mainPO) {
              const hierarchy = getHierarchyForPO(mainPO, invoicePOs, invoiceRequests);
              setPOHierarchy(hierarchy);
            }
          }
        }
      }
    }
    loadSelectedInvoiceAndHierarchy();
  }, [initialFilters, invoicePOs, invoiceRequests]);

  useEffect(() => {
    const style = document.createElement('style');
    style.innerHTML = '[class*="contentContainer-"] { inset: unset !important; }';
    document.head.appendChild(style);
    return () => { document.head.removeChild(style); };
  }, []);


  const filteredItems = React.useMemo(() => {
    const searchLower = searchText?.toLowerCase().trim() || "";
    return invoiceRequests.filter(item => {
      const matchesSearch = !searchLower || Object.values(item).some(val =>
        val !== undefined && val !== null && val.toString().toLowerCase().includes(searchLower)
      );
      const matchesFinanceStatus = !filterFinanceStatus || filterFinanceStatus === "All"
        ? true
        : item.FinanceStatus === filterFinanceStatus;
      const matchesProject =
        !filterProjectName.length ||
        filterProjectName.includes("All") ||
        filterProjectName.includes(item.ProjectName || "");
      const matchesCurrentStatus = !filterCurrentStatus || filterCurrentStatus.length === 0 || filterCurrentStatus.includes("All")
        ? true : filterCurrentStatus.includes(item.CurrentStatus || "");

      const matchesInvoiceStatus = !filterInvoiceStatus || filterInvoiceStatus.length === 0 || filterInvoiceStatus.includes("All")
        ? true : filterInvoiceStatus.includes(item.Status || "");

      const matchesInvoicePercent = !invoicePercentStatusFilter || invoicePercentStatusFilter.length === 0 || invoicePercentStatusFilter.includes("All")
        ? true
        : (() => {
          const percent = calculateInvoicedPercentForPO(item.PurchaseOrder, invoiceRequests);
          const epsilon = 0.0001;
          // if (invoicePercentStatusFilter.includes("NotPaid")) return Math.abs(percent) < epsilon;
          if (invoicePercentStatusFilter.includes("NotPaid"))
            return Math.abs(percent) < epsilon;
          // if (invoicePercentStatusFilter.includes("PartiallyInvoiced")) return !(Math.abs(percent) < epsilon || Math.abs(percent - 100) < epsilon);
          if (invoicePercentStatusFilter.includes("PartiallyInvoiced"))
            return percent > epsilon && percent < (100 - epsilon);
          if (invoicePercentStatusFilter.includes("CompletelyInvoiced")) return Math.abs(percent - 100) < epsilon;
          // return true;
        })();

      return matchesSearch && matchesProject && matchesCurrentStatus && matchesInvoiceStatus && matchesInvoicePercent && matchesFinanceStatus;
    });
  }, [invoiceRequests, searchText, filterProjectName, filterCurrentStatus, filterInvoiceStatus, filterFinanceStatus, invoicePercentStatusFilter]);

  React.useEffect(() => {
    if (!filteredItems) return;
    const sorted = _copyAndSort(
      filteredItems,
      sortedColumnKey ?? 'PurchaseOrder', // default sort key
      isSortedDescending
    );
    setSortedFilteredItems(sorted);
  }, [filteredItems, sortedColumnKey, isSortedDescending]);

  async function isCurrentUserAdmin(sp: SPFI, context: any): Promise<boolean> {
    const email = context.pageContext.user.email.toLowerCase();
    const admins = await sp.web.siteGroups
      .getByName("admin")
      .users();
    return admins.some(u => u.Email?.toLowerCase() === email);
  }

  function getAmountBuckets(item: any) {
    const status = (item.Status || "").toLowerCase();
    const amount = Number(item.InvoiceAmount || 0);

    return {
      requested:
        status === "invoice requested" ? amount : 0,
      invoiced:
        status === "invoice raised" || status === "pending payment"
          ? amount
          : 0,
      paid:
        status === "payment received" ? amount : 0,
    };
  }

  function getCurrencyByPOID(poID: string, mainPOs: Array<{ POID: string; Currency?: string }>): string {
    const mainPO = mainPOs.find(po => po.POID === poID);
    return mainPO?.Currency ?? '';  // fallback to empty string if not found
  }
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

  async function sendMailWithGraph(
    graphClient: MSGraphClient,
    toRecipients: string[],
    subject: string,
    body: string
  ): Promise<void> {
    const message = {
      subject,
      body: { contentType: 'HTML', content: body },
      toRecipients: toRecipients.map(address => ({ emailAddress: { address } })),
    };

    await graphClient.api('/me/sendMail').post({ message });
  }

  async function sendClarificationGivenEmail(item: InvoiceRequest) {
    const toArray = await getFinanceEmails();
    if (!toArray.length) return;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const appPageUrl = getCurrentPageUrl ? getCurrentPageUrl() : `${siteUrl}/SitePages/InvoiceTracker.aspx`;
    const financeLink = `${appPageUrl}#updaterequests?selectedInvoice=${item.Id}`;
    const siteTitle = context.pageContext.web.title;
    // financeEmails string can come from a config list / env; here assumed on the item or from your settings
    // const financeEmails = item.FinanceEmails || ''; // adjust to your source
    // if (!financeEmails) return;

    // const toArray = financeEmails.split(',').map((e: any) => e.trim()).filter(Boolean);

    const body = `
<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;background:#f9f9f9;border-radius:10px;padding:24px;">
  <div style="font-size:18px;font-weight:600;color:#0078d4;margin-bottom:16px;">
    Clarification Provided for Invoice Request
  </div>
  <div style="font-size:16px;color:#444;margin-bottom:18px;">
    The requestor has submitted clarification for the invoice request.
  </div>
  <table style="width:100%;border-collapse:collapse;font-size:15px;color:#333;margin-bottom:20px;">
    <tr>
      <td style="font-weight:600;padding:6px 0;">PO ID:</td>
      <td>${item.PurchaseOrder}</td>
    </tr>
    <tr>
      <td style="font-weight:600;padding:6px 0;">Project Name:</td>
      <td>${item.ProjectName ?? ''}</td>
    </tr>
    <tr>
      <td style="font-weight:600;padding:6px 0;">PO Item Title:</td>
      <td>${item["POItem_x0020_Title"] ?? ''}</td>
    </tr>
  </table>
  <div style="margin-bottom:24px;">
    <a href="${financeLink}" style="font-size:15px;color:#0078d4;text-decoration:underline;">
      Click here to review the clarification
    </a>
  </div>
  <div style="border-top:1px solid #eee;margin-top:22px;padding-top:10px;font-size:13px;color:#999;">
    Invoice Tracker | SACHA Group
  </div>
</div>
`;

    const subject = `[${siteTitle}]Clarification submitted for PO ${item.PurchaseOrder}`;

    const graphClient = await context.msGraphClientFactory.getClient();
    await sendMailWithGraph(graphClient, toArray, subject, body);
  }

  async function fetchInvoiceRequestById(sp: SPFI, id: number): Promise<InvoiceRequest | null> {
    try {
      const item = await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.getById(id)
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
          "AttachmentFiles"
        )
        .expand("Author", "Editor", "AttachmentFiles")();

      return {
        Id: item.Id,
        Title: item.Title || "",                      // Add Title if required
        PurchaseOrderId: item.PurchaseOrder || "",   // mapped to PurchaseOrderId as interface expects
        PurchaseOrder: item.PurchaseOrder,            // include any other that are in interface
        POAmount: item.InvoiceAmount,
        Status: item.Status,
        ProjectName: item.ProjectName,
        POItem_x0020_Title: item.POItem_x0020_Title,           // Use "POItemx0020Title" not "POItem_x0020_Title"
        POItem_x0020_Value: item.POItem_x0020_Value,           // Use correct casing as in SP
        Customer_x0020_Contact: item.Customer_x0020_Contact,
        PMCommentsHistory: item.PMCommentsHistory,
        FinanceCommentsHistory: item.FinanceCommentsHistory,
        Created: item.Created,
        Author: item.Author?.Title || "",
        Modified: item.Modified,
        Editor: item.Editor?.Title || "",
        CurrentStatus: item.CurrentStatus,
        AttachmentFiles: item.AttachmentFiles,
        InvoiceAmount: item.InvoiceAmount,
      };

    } catch (error) {
      console.error("Error fetching invoice request by ID", error);
      return null;
    }
  }

  function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    return items.slice().sort((a, b) => {
      const aVal = (a as any)[columnKey];
      const bVal = (b as any)[columnKey];

      if (aVal === undefined || aVal === null) return 1;
      if (bVal === undefined || bVal === null) return -1;

      if (typeof aVal === 'string' && typeof bVal === 'string') {
        const comparison = aVal.localeCompare(bVal);
        return isSortedDescending ? -comparison : comparison;
      }

      if (typeof aVal === 'number' && typeof bVal === 'number') {
        return isSortedDescending ? bVal - aVal : aVal - bVal;
      }

      if (aVal instanceof Date && bVal instanceof Date) {
        return isSortedDescending ? bVal.getTime() - aVal.getTime() : aVal.getTime() - bVal.getTime();
      }

      return 0;
    });
  }


  function findMainPO(request: InvoiceRequest, allPOs: InvoicePO[]): InvoicePO | undefined {
    let po = allPOs.find((p) => p.POID === request.PurchaseOrder);
    while (po && po.ParentPOID) {
      po = allPOs.find((p) => p.POID === po.ParentPOID);
    }
    console.log(po)
    return po;
  }

  function getLineItemsList(h: POHierarchy | null) {
    if (!h) return [];
    // Adapt line items to POItems table structure for the chilPO table UI
    return h.lineItemGroups.map(g => ({
      // Map 'Title' to POItem Title, 'Value' to POItem Value
      POItem_x0020_Title: g.poItem.Title,              // Displayed as POItem Title
      POItem_x0020_Value: g.poItem.Value,              // Displayed as POItem Value
      Comments: g.poItem.Comments || "",
    }));
  }

  async function getFinanceEmails(): Promise<string[]> {
    const items = await sp.web.lists
      .getByTitle("InvoiceConfiguration")
      .items.filter("Title eq 'FinanceEmail'")();

    const financeEmails: string =
      items.length > 0 && items[0].Value ? items[0].Value : "";

    return financeEmails
      .split(",")
      .map(e => e.trim())
      .filter(e => !!e);
  }

  async function handleClarifySubmit() {
    setClarifyLoading(true);

    const remaining = getRemainingAmountForCurrentPOItem(selectedReq);
    const newAmount = clarifyInvoiceAmount || 0;
    if (newAmount > remaining) {
      setDialogMessage(
        `Invoiced Amount (${newAmount.toLocaleString()}) cannot be greater than the remaining PO Item value (${remaining.toLocaleString()}).`
      );
      setDialogType("error");
      setDialogVisible(true);
      return;
    }

    try {
      // Fetch the item
      const item = await sp.web.lists.getByTitle('Invoice Requests').items.getById(selectedReq.Id).select('PMCommentsHistory')();
      let history = [];

      // Parse existing history only if clarifyComment is non-empty
      if (clarifyComment && clarifyComment.trim().length > 0) {
        if (item.PMCommentsHistory) {
          try {
            const decodedJson = decodeHtmlEntities(item.PMCommentsHistory);
            history = JSON.parse(decodedJson);
            if (!Array.isArray(history)) history = [history];
          } catch {
            history = [];
          }
        }

        // Append new comment
        const userRole = await getCurrentUserRole(context, selectedReq);
        history.push({
          Date: new Date().toISOString(),
          Title: 'Clarification',
          User: context.pageContext.user.displayName,
          Role: userRole,
          Data: clarifyComment,
        });
      } else {
        // If no comment, keep history unchanged (no new entry)
        if (item.PMCommentsHistory) {
          try {
            const decodedJson = decodeHtmlEntities(item.PMCommentsHistory);
            history = JSON.parse(decodedJson);
            if (!Array.isArray(history)) history = [history];
          } catch {
            history = [];
          }
        }
      }

      // Prepare the update payload
      const updatePayload: any = {
        InvoiceAmount: clarifyInvoiceAmount,
        PMStatus: "Submitted",
        FinanceStatus: "Pending",
        Customer_x0020_Contact: clarifyCustomerContact,
        CurrentStatus: `Clarified by ${await getCurrentUserRole(context, selectedReq)}`
      };

      // Only update PMCommentsHistory if comment was provided
      if (clarifyComment && clarifyComment.trim().length > 0) {
        updatePayload.PMCommentsHistory = JSON.stringify(history);
      }

      await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.getById(selectedReq.Id)
        .update(updatePayload);
      try {
        await sendClarificationGivenEmail(selectedReq);
      } catch (e) {
        console.error("Failed to send clarification email to finance", e);
      }
      setShowClarifyPanel(false);
      setShowHierPanel(false);
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

  function getCurrencySymbol(currencyCode: string, locale = 'en-US'): string {
    return new Intl.NumberFormat(locale, {
      style: 'currency',
      currency: currencyCode,
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    })
      .formatToParts(1)
      .find(part => part.type === 'currency')?.value || currencyCode;
  }

  function calculateInvoicedPercentForPO(
    mainPOID: string,
    invoiceRequests: InvoiceRequest[],
  ): number {
    const mainPO = invoicePOs.find(po => po.POID === mainPOID);
    if (!mainPO) return 0;
    const mainPOAmount = Number(mainPO.POAmount);
    if (!mainPOAmount) return 0;

    const totalInvoiced = invoiceRequests
      .filter(inv => inv.PurchaseOrder === mainPOID && (inv.Status?.toLowerCase() === "payment received" || inv.Status?.toLowerCase() === "invoice raised" || inv.Status?.toLowerCase() === "pending payment"))
      .reduce((sum, inv) => sum + (inv.InvoiceAmount || 0), 0);  // use Amount here

    return (totalInvoiced / mainPOAmount) * 100;
  }

  function calculateInvoicedPercentForPOItem(poID: string, poItemTitle: string, poItemAmount: number, invoiceRequests: InvoiceRequest[]): number {
    if (!poItemAmount) return 0;

    // Filter out cancelled and sum amounts for POItem with POID and POItem Title
    const totalInvoiced = invoiceRequests
      .filter(inv => inv.PurchaseOrder === poID && inv.POItem_x0020_Title === poItemTitle && (inv.Status?.toLowerCase() === "payment received" || inv.Status?.toLowerCase() === "invoice raised" || inv.Status?.toLowerCase() === "pending payment"))
      .reduce((sum, inv) => sum + (inv.InvoiceAmount || 0), 0);

    return (totalInvoiced / poItemAmount) * 100;
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
        ();

      // const projectNameFromInvoice = poId?.ProjectName;
      // const matchedProject = projects.find((p: any) => {
      //   const projectTitle = (p.Title ?? "").toString().trim().toLowerCase();
      //   const invoiceProjectName = (projectNameFromInvoice ?? "").toString().trim().toLowerCase();

      //   return projectTitle === invoiceProjectName;
      // });

      // const matchedProject = projects[0];
      if (isAdmin) {
        return "Admin";
      }
      const matchedProject = projects[0];
      if (!matchedProject) {
        return "";
      }
      // if (!matchedProject) return "Unknown Role";
      if (matchedProject.DH?.EMail.toLowerCase() === currentUserEmail) return "DH";
      if (matchedProject.DM?.EMail.toLowerCase() === currentUserEmail) return "DM";
      if (matchedProject.PM?.EMail.toLowerCase() === currentUserEmail) return "PM";
      return "";
    } catch (error) {
      console.error("Error determining user role:", error);
      return "";
    }
  }
  function getAllPOItemsForList(h: POHierarchy | null) {
    if (!h) {
      return [];
    }

    const lineItems = h.lineItemGroups.map((g) => {
      return {
        POID: g.poItem.POID,
        POAmount: g.poItem.POAmount,
      };
    });

    const childPOs = h.childPOGroups.map((g) => {
      return {
        POID: g.childPO.POID,
        POAmount: g.childPO.POAmount,
      };
    });

    const combined = [...lineItems, ...childPOs];
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

  // const filteredInvoiceRequests = invoiceRequests.filter(item => {
  //   const searchLower = searchText.toLowerCase().trim();
  //   const matchesSearch = !searchLower || Object.values(item).some(val =>
  //     val !== undefined &&
  //     val !== null &&
  //     val.toString().toLowerCase().includes(searchLower)
  //   );
  //   // const matchesProject = !filterProjectName || filterProjectName === "All" || item.ProjectName === filterProjectName;
  //   // const matchesInvoiceStatus = !filterInvoiceStatus || filterInvoiceStatus === "All" || item.Status === filterInvoiceStatus;
  //   const matchesProject =
  //     !filterProjectName.length ||
  //     filterProjectName.includes("All") ||
  //     filterProjectName.includes(item.ProjectName || "");
  //   const matchesInvoiceStatus =
  //     !filterInvoiceStatus.length ||
  //     filterInvoiceStatus.includes("All") ||
  //     filterInvoiceStatus.includes(item.Status || "");
  //   const matchesCurrentStatus =
  //     !filterCurrentStatus.length ||
  //     filterCurrentStatus.includes("All") ||
  //     filterCurrentStatus.includes(item.CurrentStatus || "");

  //   // const matchesCurrentStatus = !filterCurrentStatus || filterCurrentStatus === "All" || item.CurrentStatus === filterCurrentStatus;
  //   if (invoicePercentStatusFilter) {
  //     // Calculate percent invoiced for this item's PurchaseOrder
  //     const percent = calculateInvoicedPercentForPO(item.PurchaseOrder, invoiceRequests);
  //     if (invoicePercentStatusFilter.includes("NotPaid") && percent !== 0) return false;
  //     if (invoicePercentStatusFilter.includes("PartiallyInvoiced") && !(percent > 0 && percent < 100)) return false;
  //     if (invoicePercentStatusFilter.includes("CompletelyInvoiced") && percent !== 100) return false;
  //   }
  //   return matchesSearch && matchesProject && matchesInvoiceStatus && matchesCurrentStatus;
  // });

  const filteredInvoiceRequests = React.useMemo(() => {
    const searchLower = searchText.toLowerCase().trim();

    return invoiceRequests.filter(item => {
      const matchesSearch =
        !searchLower ||
        Object.values(item).some(val =>
          val !== undefined &&
          val !== null &&
          val.toString().toLowerCase().includes(searchLower)
        );

      const matchesColumnFilters = Object.entries(columnFilters).every(([colKey, selectedVals]) => {
        if (!selectedVals || selectedVals.length === 0) return true;
        const col = columns.find(c => c.key === colKey);
        if (!col || !col.fieldName) return true;
        const value = (item as any)[col.fieldName];
        if (value === null || value === undefined || value === '') return false;
        const vStr = value.toString();
        return selectedVals.includes(vStr);
      });

      // existing project/status/currentStatus/invoice % filters stay as‑is
      const matchesProject =
        !filterProjectName.length ||
        filterProjectName.includes("All") ||
        filterProjectName.includes(item.ProjectName || "");
      const matchesInvoiceStatus =
        !filterInvoiceStatus.length ||
        filterInvoiceStatus.includes("All") ||
        filterInvoiceStatus.includes(item.Status || "");
      const matchesCurrentStatus =
        !filterCurrentStatus.length ||
        filterCurrentStatus.includes("All") ||
        filterCurrentStatus.includes(item.CurrentStatus || "");

      let matchesInvoicePercent = true;
      // if (invoicePercentStatusFilter?.length) {
      //   const percent = calculateInvoicedPercentForPO(item.PurchaseOrder, invoiceRequests);
      //   if (invoicePercentStatusFilter.includes("NotPaid") && percent !== 0) matchesInvoicePercent = false;
      //   if (invoicePercentStatusFilter.includes("PartiallyInvoiced") && !(percent > 0 && percent < 100)) matchesInvoicePercent = false;
      //   if (invoicePercentStatusFilter.includes("CompletelyInvoiced") && percent !== 100) matchesInvoicePercent = false;
      // }

      if (invoicePercentStatusFilter?.length && !invoicePercentStatusFilter.includes("All")) {
        const raw = calculateInvoicedPercentForPO(item.PurchaseOrder, invoiceRequests);
        const percent = isNaN(raw) ? 0 : raw;   // treat empty / NaN as 0
        const epsilon = 0.0001;

        const isNotPaid = Math.abs(percent) < epsilon;                       // 0 or empty
        const isPartial = percent > epsilon && percent < 100 - epsilon;     // strictly between 0 and 100
        const isComplete = Math.abs(percent - 100) < epsilon;                 // ~= 100

        const wantsNotPaid = invoicePercentStatusFilter.includes("NotPaid");
        const wantsPartial = invoicePercentStatusFilter.includes("PartiallyInvoiced");
        const wantsComplete = invoicePercentStatusFilter.includes("CompletelyInvoiced");

        const matches =
          (wantsNotPaid && isNotPaid) ||
          (wantsPartial && isPartial) ||
          (wantsComplete && isComplete);

        if (!matches) {
          matchesInvoicePercent = false;
        }
      }


      return matchesSearch &&
        matchesColumnFilters &&
        matchesProject &&
        matchesInvoiceStatus &&
        matchesCurrentStatus &&
        matchesInvoicePercent;
    });
  }, [invoiceRequests, columns, searchText, columnFilters,
    filterProjectName, filterInvoiceStatus, filterCurrentStatus,
    invoicePercentStatusFilter]);

  // console.log(filteredInvoiceRequests)

  async function onInvoiceRequestSelect(item?: InvoiceRequest) {
    try {
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

            const baseUrl = getCurrentPageUrl ? getCurrentPageUrl() : window.location.href;
            const newUrl = `${baseUrl}#myrequests?selectedInvoice=${item.Id}`;
            window.history.replaceState(null, '', newUrl);

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
                ? { POID: selectedPO.poItem.Title, POAmount: selectedPO.poItem.Value, Currency: selectedPO.poItem.Currency }
                : { POID: selectedPO.childPO.POID, POAmount: selectedPO.childPO.POAmount, Currency: (selectedPO.childPO as any).Currency ?? "" }
              : null
          );
        } else {
          setSelectedPOItem(null);
        }
      } else {
        setSelectedReq(null);
        setSelectedProject(null);
        setPOHierarchy(null);
        setSelectedPOItem(null);
        setShowHierPanel(false);
      }
    } catch (error) {
      console.error("Error in onInvoiceRequestSelect", error);
      // Optionally reset states on unexpected error
      setSelectedReq(null);
      setSelectedProject(null);
      setPOHierarchy(null);
      setSelectedPOItem(null);
      setShowHierPanel(false);
    }
  }

  function getRemainingAmountForCurrentPOItem(selected: InvoiceRequest | null): number {
    if (!selected) return 0;

    const poId = selected.PurchaseOrder;
    const poItemTitle = selected["POItem_x0020_Title"];
    const poItemValue = selected["POItem_x0020_Value"] || 0;

    // Sum all non-cancelled requests for same PO + PO item, excluding the current request
    const alreadyInvoiced = invoiceRequests
      .filter(req =>
        req.PurchaseOrder === poId &&
        req["POItem_x0020_Title"] === poItemTitle &&
        req.Id !== selected.Id &&
        req.Status !== "Cancelled"
      )
      .reduce((sum, req) => sum + (req.InvoiceAmount || 0), 0);

    const remaining = poItemValue - alreadyInvoiced;
    return remaining < 0 ? 0 : remaining;
  }

  function normalizeSelectedPOItem(item: any): { POID: string, POAmount: string, Currency: string } | null {
    if (!item) return null;
    return {
      POID: item.POID ?? item.POItem_x0020_Title,
      POAmount: item.POAmount ?? item.POItem_x0020_Value,
      Currency: item.Currency ?? "",
    };
  }

  return (
    <div style={{ padding: 16 }} className="rootContainer">
      <h2>My Invoice Requests</h2>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div className={styles.filterAndHeaderSection}>
        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" styles={{ root: { marginBottom: 12 } }}>
          <div>
            <Label>Search</Label>
            <TextField
              placeholder="Search"
              value={searchText}
              onChange={(e, val) => setSearchText(val || "")}
              styles={{ root: { width: 270 } }}  // Double length
            />
          </div>
          <div>
            <SearchableDropdown
              labelText="Project Name"
              multiSelect={true}
              options={filteredProjectOptions}
              selectedItems={
                filterProjectName && filterProjectName.length ? filterProjectName : ["All"]
              }
              placeholder="Select project(s)"
              onChangeHandler={(_, option) => {
                if (!option) return;
                const key = option.key.toString();
                setFilterProjectName(prev => {
                  if (!prev || !prev.length || prev.includes("All")) prev = [];
                  if (key === "All") return ["All"];          // reset
                  const exists = prev.includes(key);
                  const next = exists ? prev.filter(k => k !== key) : [...prev, key];
                  return next.length ? next : ["All"];
                });
              }}
              // onRenderOption={renderProjectTitle}          // keep your custom title/option
              styles={dropdownStyles}
              disabled={false}
            />
            {/* <Dropdown
              label="Project Name"
              multiSelect
              options={filteredProjectOptions}
              selectedKeys={
                filterProjectName && filterProjectName.length
                  ? filterProjectName
                  : ['All']
              }
              onChange={(_, option) => {
                if (!option) return
                const key = option.key.toString()
                setFilterProjectName(prev => {
                  if (!prev || !prev.length || prev.includes('All')) prev = []
                  if (key === 'All') return ['All']        // reset
                  const exists = prev.includes(key)
                  const next = exists ? prev.filter(k => k !== key) : [...prev, key]
                  return next.length ? next : ['All']
                })
              }}
              onRenderTitle={renderProjectTitle}
              styles={dropdownStyles}
            /> */}
          </div>
          <div>
            {/* <Label>Current Status</Label> */}
            {/* <Dropdown
              label="Current Status"
              options={CURRENT_STATUS_OPTIONS}
              multiSelect
              selectedKeys={filterCurrentStatus}        // <- array
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key as string;
                setFilterCurrentStatus(prev =>
                  prev.includes(key)
                    ? prev.filter(k => k !== key)       // unselect
                    : [...prev, key]                    // select
                );
              }}
              placeholder="Current Status"
              styles={dropdownStyles}
            /> */}
            <Dropdown
              label="Current Status"
              options={CURRENT_STATUS_OPTIONS}
              multiSelect
              selectedKeys={
                filterCurrentStatus && filterCurrentStatus.length
                  ? filterCurrentStatus
                  : ["All"]
              }
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key as string;
                setFilterCurrentStatus(prev => {
                  if (!prev || !prev.length || prev.includes("All")) prev = [];
                  if (key === "All") return ["All"];
                  const exists = prev.includes(key);
                  const next = exists ? prev.filter(k => k !== key) : [...prev, key];
                  return next.length ? next : ["All"];
                });
              }}
              styles={dropdownStyles}
            />
          </div>
          <div>
            <Dropdown
              label="Invoice Status"
              options={INVOICE_STATUS_OPTIONS}
              multiSelect
              selectedKeys={
                filterInvoiceStatus && filterInvoiceStatus.length
                  ? filterInvoiceStatus
                  : ["All"]
              }
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key as string;

                // Click on "All" resets to only "All"
                if (key === "All") {
                  setFilterInvoiceStatus(["All"]);
                  return;
                }

                setFilterInvoiceStatus(prev => {
                  // remove "All" when selecting a specific value
                  const withoutAll = prev.filter(k => k !== "All");
                  return withoutAll.includes(key)
                    ? withoutAll.filter(k => k !== key)   // unselect
                    : [...withoutAll, key];               // select
                });
              }}
              placeholder="Invoice Status"
              styles={dropdownStyles}
            />
          </div>
          <div>
            <Dropdown
              label="Invoice % Status"
              options={invoicePercentStatusOptions}
              multiSelect
              selectedKeys={
                invoicePercentStatusFilter && invoicePercentStatusFilter.length
                  ? invoicePercentStatusFilter
                  : ["All"]
              }
              onChange={(_, option) => {
                if (!option) return;
                const key = option.key as string;

                if (key === "All") {
                  // reset to only All
                  setInvoicePercentStatusFilter(["All"]);
                  return;
                }

                setInvoicePercentStatusFilter(prev => {
                  // drop "All" when choosing specific values
                  const withoutAll = prev.filter(k => k !== "All");
                  return withoutAll.includes(key)
                    ? withoutAll.filter(k => k !== key)   // unselect
                    : [...withoutAll, key];               // select
                });
              }}
              styles={dropdownStyles}
            />
          </div>
          <div>
            <PrimaryButton
              text="Clear"
              onClick={clearAllFilters}
              style={{ alignSelf: "center", marginLeft: 20 }}
              disabled={isClearDisabled}
            />
          </div>
          <IconButton
            iconProps={{ iconName: 'ExcelDocument' }}
            title="Export to Excel"
            ariaLabel="Export to Excel"
            onClick={handleExportToExcel}
            styles={{ root: { color: primaryColor } }}
          />
          <IconButton
            iconProps={{ iconName: 'Columns' }}
            title="Manage Columns"
            ariaLabel="Manage Columns"
            onClick={() => setIsColumnPanelOpen(true)}
            styles={{ root: { color: primaryColor } }}
          />
        </Stack>
      </div>
      {loading ? (
        <Spinner label="Loading..." />
      ) : (
        <>
          <div className={`ms-Grid-row ${styles.detailsListContainer}`}>
            <div style={{ height: 900, position: 'relative' }}>
              <ScrollablePane>
                {/* <ScrollablePane> */}
                <div
                  className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.detailsList_Scrollablepane_Container}`}
                >
                  <div style={{ width: '100%', overflowX: 'auto' }}>
                    <DetailsList
                      items={filteredInvoiceRequests}
                      columns={getVisibleColumns()}
                      isHeaderVisible={true}
                      setKey="invoiceRequestList"
                      // layoutMode={DetailsListLayoutMode.justified}
                      onActiveItemChanged={onInvoiceRequestSelect}
                      selectionPreservedOnEmptyClick={true}
                      selectionMode={SelectionMode.single}
                      onRenderDetailsHeader={onRenderDetailsHeader}
                      styles={{ root: { height: '100%', minWidth: 1400 } }}
                    />
                  </div>
                </div>
                {/* {columnFilterMenu.visible && (
                  <ContextualMenu
                    items={menuItems}
                    target={columnFilterMenu.target}
                    onDismiss={() =>
                      setColumnFilterMenu({ visible: false, target: null, columnKey: null })
                    }
                  />
                )} */}
                {columnFilterMenu.visible && columnFilterMenu.target && (
                  <ContextualMenu
                    target={columnFilterMenu.target}
                    items={menuItems}
                    onDismiss={() => setColumnFilterMenu({ visible: false, target: null, columnKey: null })}
                  />
                )}
              </ScrollablePane>
            </div>
          </div>
          <Panel
            isOpen={showHierPanel}
            isBlocking={true}
            isLightDismiss={false}
            onDismiss={() => {
              // Prevent closing if viewer panel is open
              const fragment = window.location.hash.substring(1);
              const [tab, query] = fragment.split("?");
              const params = new URLSearchParams(query || "");
              params.delete("selectedInvoice");

              const newFragment = params.toString() ? `${tab}?${params.toString()}` : tab;
              window.history.replaceState(null, '', `#${newFragment}`);

              if (showClarifyPanel) return;
              if (!viewerUrl) {
                setShowHierPanel(false);
                setSelectedReq(null);
                setSelectedPOItem(null);
                setPOHierarchy(null);
              }

            }}
            // isBlocking={!!viewerUrl}
            // headerText={`Invoice Details:`}
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
                {selectedReq.CurrentStatus === "Finance asked Clarification" && (
                  <div style={{ margin: "16px 0" }}>
                    <PrimaryButton
                      onClick={() => {
                        setClarifyInvoiceAmount(selectedReq.InvoiceAmount);
                        setClarifyCustomerContact(selectedReq.Customer_x0020_Contact);
                        setClarifyComment("");
                        setShowClarifyPanel(true);
                      }}
                    // style={{ padding: '8px 24px', background: '#166BDD', color: '#fff', borderRadius: 4, border: 'none' }}
                    >
                      Clarify
                    </PrimaryButton>
                  </div>
                )}
              </>
            )}
            {poHierarchy && poHierarchy.lineItemGroups.length > 0 && (
              <div style={{ marginTop: 16 }}>
                <strong>All PO Items for {poHierarchy.mainPO.POID}</strong>
                <DetailsList
                  items={getLineItemsList(poHierarchy)}
                  onActiveItemChanged={(item) => setSelectedPOItem(normalizeSelectedPOItem(item))}
                  columns={poColumnsLine}
                  selectionMode={SelectionMode.single}
                  setKey="lineItemsList"
                  styles={{ root: { marginBottom: 16 } }}
                />
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                  <strong>
                    Invoice Requests for {selectedPOItem ? selectedPOItem.POID : poHierarchy.mainPO.POID}
                  </strong>
                  <PrimaryButton
                    text={`Show all Invoice Requests`}
                    onClick={() => setSelectedPOItem(null)}
                    disabled={!selectedPOItem}
                  // style={{
                  //   marginLeft: 12,
                  //   // color: "white",
                  //   // background: primaryColor,
                  //   fontWeight: 600,
                  //   borderRadius: 4,
                  //   padding: "4px 16px"
                  // }}
                  />
                </div>

                <DetailsList
                  items={getFilteredRequests()}
                  columns={groupedInvColumns}
                  selectionMode={SelectionMode.single}
                  onActiveItemChanged={onInvoiceRequestClicked}
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
                  <PrimaryButton
                    text="Show all Invoice Requests"
                    onClick={() => setSelectedPOItem(null)}
                    disabled={!selectedPOItem}
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

                </div>

                <DetailsList
                  items={getFilteredRequests()}
                  columns={groupedInvColumns}
                  selectionMode={SelectionMode.single}
                  // selection={selection}
                  onActiveItemChanged={onInvoiceRequestClicked}
                  setKey="invoiceRequestsListByPO"
                />
              </div>
            )}
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
                  label="Invoiced Amount"
                  value={clarifyInvoiceAmount?.toString() || ''}
                  onChange={(_, val) => setClarifyInvoiceAmount(val ? Number(val) : undefined)}
                  required
                />
                <TextField
                  label="Add Comment"
                  multiline
                  value={clarifyComment}
                  required
                  onChange={(_, val) => setClarifyComment(val || '')}
                />
                <TextField
                  label="Requestor Comment"
                  value={formatCommentsHistory(selectedReq.PMCommentsHistory)}
                  multiline
                  disabled
                />
                <TextField
                  label="Finance Comments"
                  value={formatCommentsHistory(selectedReq.FinanceCommentsHistory)}
                  multiline
                  disabled
                />
                <div style={{ marginTop: 12 }}>
                  <PrimaryButton
                    // type="button"
                    disabled={clarifyLoading || clarifyInvoiceAmount === undefined}
                    onClick={handleClarifySubmit}
                    style={{
                      background: primaryColor,
                      color: '#fff',
                      padding: '8px 24px',
                      border: 'none',
                      // borderRadius: 4,
                      cursor: 'pointer'
                    }}
                  >
                    Submit
                  </PrimaryButton>
                </div>
              </>
            )}
          </Panel>
          {/* <Panel
            isOpen={!!viewerUrl}
            onDismiss={() => {
              setViewerUrl(null);
              setViewerName(null);
              // Do NOT close parent panel here to keep parent open
            }}
            headerText={viewerName ?? "Document Viewer"}
            type={PanelType.large}
            closeButtonAriaLabel="Close"
          >
            {viewerUrl && viewerName && (
              <DocumentViewer
                url={viewerUrl}
                isOpen
                onDismiss={() => {
                  setViewerUrl(null);
                  setViewerName(null);
                }}
                fileName={viewerName}
              />
            )}
          </Panel> */}
          <Panel
            isOpen={!!viewerUrl}
            onDismiss={() => {
              setViewerUrl(null);
              setViewerName(null);
              // Do NOT close parent panel here to keep parent open
            }}
            headerText={viewerName ?? 'Document Viewer'}
            type={PanelType.large}
            isLightDismiss
            closeButtonAriaLabel="Close"
            styles={{
              content: { height: "100vh", display: "flex", flexDirection: "column", padding: 0 }
            }}
          >
            <div style={{ flex: 1, minHeight: "100vh", display: 'flex', flexDirection: 'column' }}>
              {viewerUrl && viewerName && (
                <DocumentViewer
                  url={viewerUrl}
                  isOpen={!!viewerUrl}
                  fileName={viewerName}
                  onDismiss={() => {
                    setViewerUrl(null);
                    setViewerName(null);
                  }}
                />
              )}
            </div>
          </Panel>

          <Panel
            isOpen={isInvoiceRequestViewPanelOpen}
            onDismiss={() => {
              setIsInvoiceRequestViewPanelOpen(false);
              setSelectedInvoiceRequest(null);
            }}
            headerText="Invoice Request Details"
            type={PanelType.medium}
            styles={{
              content: { padding: 20 }, // More generous padding
              headerText: { fontWeight: 600, fontSize: 22, color: primaryColor }
            }}
          >
            {selectedInvoiceRequest && (() => {
              const { requested, invoiced, paid } = getAmountBuckets(selectedInvoiceRequest);
              const symbol = getCurrencySymbol(selectedInvoiceRequest.Currency);
              return (
                < Stack tokens={{ childrenGap: 24 }}>
                  {/* <Icon iconName="PageDetails" styles={{ root: { fontSize: 36, color: primaryColor, marginBottom: 12 } }} /> */}

                  {/* Main Data Grid */}
                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(2, 1fr)', // two columns
                    gap: 24,
                    background: '#f9f9f9',
                    padding: 20,
                    borderRadius: 12,
                    marginBottom: 24
                  }}>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>PO Item Title: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{renderValue(selectedInvoiceRequest.POItem_x0020_Title)}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>PO Item Value: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{getCurrencySymbol(selectedInvoiceRequest.Currency)} {renderValue(selectedInvoiceRequest.POItem_x0020_Value)}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>
                        Requested Amount:{" "}
                      </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>
                        {requested ? `${symbol} ${requested.toLocaleString()}` : "-"}
                      </Text>
                    </div>

                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>
                        Invoiced Amount:{" "}
                      </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>
                        {invoiced ? `${symbol} ${invoiced.toLocaleString()}` : "-"}
                      </Text>
                    </div>

                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>
                        Paid Amount:{" "}
                      </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>
                        {paid ? `${symbol} ${paid.toLocaleString()}` : "-"}
                      </Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>Invoice Status: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{renderValue(selectedInvoiceRequest.Status)}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>POID Invoiced % (PO): </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>
                        {calculateInvoicedPercentForPO(
                          selectedInvoiceRequest.PurchaseOrder,
                          invoiceRequests
                        ).toFixed(0)}%
                      </Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>POItem Invoiced %: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>
                        {calculateInvoicedPercentForPOItem(
                          selectedInvoiceRequest.PurchaseOrder,
                          selectedInvoiceRequest.POItem_x0020_Title,
                          selectedInvoiceRequest.POItem_x0020_Value,
                          invoiceRequests
                        ).toFixed(0)}%
                      </Text>
                    </div>
                  </div>

                  {/* Comments Section */}
                  <Stack tokens={{ childrenGap: 18 }}>
                    {formatCommentsHistory(selectedInvoiceRequest.PMCommentsHistory)?.trim() && (
                      <TextField
                        label="Requestor Comments"
                        value={(formatCommentsHistory(selectedInvoiceRequest.PMCommentsHistory))}
                        multiline
                        disabled
                        styles={{
                          root: {},
                          subComponentStyles: {
                            label: {
                              root: {
                                color: primaryColor,
                                fontWeight: 600
                              }
                            }
                          }
                        }}
                      />)}
                    {formatCommentsHistory(selectedInvoiceRequest.FinanceCommentsHistory)?.trim() && (
                      <TextField
                        label="Finance Comments"
                        value={formatCommentsHistory(selectedInvoiceRequest.FinanceCommentsHistory)}
                        multiline
                        disabled
                        styles={{
                          root: {},
                          subComponentStyles: {
                            label: {
                              root: {
                                color: primaryColor,
                                fontWeight: 600
                              }
                            }
                          }
                        }}
                      />)}
                    {/* </div> */}
                  </Stack>

                  {/* Metadata */}
                  <Separator styles={{ root: { marginTop: 16, marginBottom: 16 } }} />
                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: 'repeat(2, 1fr)',
                    gap: 18,
                    padding: 14,
                    background: '#f5f5fa',
                    borderRadius: 8
                  }}>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>Created: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{new Date(selectedInvoiceRequest.Created).toLocaleDateString()}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>Created By: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{renderValue(selectedInvoiceRequest.Author?.Title)}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>Modified: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{new Date(selectedInvoiceRequest.Modified).toLocaleDateString()}</Text>
                    </div>
                    <div>
                      <Text variant="small" styles={{ root: { color: primaryColor } }}>Modified By: </Text>
                      <Text styles={{ root: { fontWeight: 400, fontSize: 12 } }}>{renderValue(selectedInvoiceRequest.Editor?.Title)}</Text>
                    </div>
                  </div>
                </Stack>
              );
            })()}
          </Panel>
          <Panel
            isOpen={isFilterPanelOpen}
            onDismiss={() => setIsFilterPanelOpen(false)}
            headerText={`Filter: ${currentFilterColumn}`}
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
                          const existing = prev[currentFilterColumn] || [];
                          const next = checked
                            ? [...existing, val]
                            : existing.filter(v => v !== val);
                          return { ...prev, [currentFilterColumn]: next };
                        });
                      }}
                    />
                  );
                })}
                <PrimaryButton
                  text="Clear"
                  onClick={() => clearColumnFilter(currentFilterColumn)}
                />
              </Stack>
            )}
          </Panel>
          {/* <Panel
            isOpen={isColumnPanelOpen}
            onDismiss={() => setIsColumnPanelOpen(false)}
            headerText="Manage Columns"
            closeButtonAriaLabel="Close"
            type={PanelType.smallFixedFar}
            isLightDismiss={false}
          >
            <Stack tokens={{ childrenGap: 12 }}>
              <Text styles={{ root: { fontWeight: 600 } }}>
                Choose which columns you want to see
              </Text>

              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text="Select all"
                  onClick={() => setVisibleColumns(columns.map(c => c.key as string))}
                />
                <PrimaryButton
                  text="Clear all"
                  onClick={() => setVisibleColumns([])}
                />
              </Stack>

              <Stack
                styles={{
                  root: {
                    maxHeight: 320,
                    overflowY: "auto",
                    border: "1px solid #eee",
                    borderRadius: 4,
                    padding: 8,
                    marginTop: 8,
                  },
                }}
                tokens={{ childrenGap: 4 }}
              >
                {columns.map(col => (
                  <Checkbox
                    key={col.key}
                    label={col.name}
                    checked={visibleColumns.includes(col.key as string)}
                    onChange={(_: React.FormEvent<HTMLElement>, checked?: boolean) => {
                      const key = col.key as string;
                      setVisibleColumns(prev =>
                        checked ? [...prev, key] : prev.filter(k => k !== key)
                      );
                    }}
                  />
                ))}
              </Stack>

              <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text="Save"
                  onClick={() => setIsColumnPanelOpen(false)}
                />
                <PrimaryButton
                  text="Cancel"
                  onClick={() => {
                    // optionally reset visibleColumns to last saved snapshot
                    setIsColumnPanelOpen(false);
                  }}
                />
              </Stack>
            </Stack>
          </Panel> */}
          <Panel
            isOpen={isColumnPanelOpen}
            onDismiss={() => setIsColumnPanelOpen(false)}
            headerText="Customize Columns"
            type={PanelType.medium}
            isBlocking={true}
          >
            <Stack tokens={{ childrenGap: 16 }}>
              <div>
                <Text styles={{ root: { fontWeight: 600 } }}>Choose which columns you want to see</Text>
              </div>

              <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                {/* <PrimaryButton
                  text="Select all"
                  onClick={() => setVisibleColumns(columns.map(c => c.key as string))}
                />
                <PrimaryButton
                  text="Clear all"
                  onClick={() => setVisibleColumns([])}
                /> */}
              </Stack>

              <div style={{ overflow: 'auto', border: '1px solid #edebe9', borderRadius: 4, padding: 12 }}>
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
                    const defaults = ['PurchaseOrder', 'ProjectName', 'Status', 'InvoiceAmount', 'DueDate']
                    setVisibleColumns(defaults)
                  }}
                />
                <PrimaryButton
                  text="Apply Changes"
                  onClick={() => setIsColumnPanelOpen(false)}
                /> */}
              </Stack>
            </Stack>
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
            <DialogFooter styles={{ actions: { justifyContent: 'center' } }}>
              <PrimaryButton onClick={() => setDialogVisible(false)} text="OK" />
            </DialogFooter>
          </Dialog>

        </>
      )
      }
    </div >
  );
}
