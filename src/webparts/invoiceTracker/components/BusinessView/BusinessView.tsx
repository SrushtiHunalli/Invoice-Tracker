import * as React from "react";
import { useEffect, useState } from "react";
import {
  Stack, Spinner, Text,
  DetailsList, IColumn,
  MessageBar, MessageBarType,
  Separator, Dropdown, IDropdownOption,
  Panel, PanelType, SearchBox
} from "office-ui-fabric-react";
import { SPFI } from "@pnp/sp";

interface BusinessViewProps {
  sp: SPFI;
  context: any;
  onNavigate?: (view: string) => void;
  projectsp: SPFI;
}

interface InvoiceRequest {
  Id: number;
  PurchaseOrder: string;
  ProjectName?: string;
  Status?: string;
  CurrentStatus?: string;
  InvoiceAmount?: number;
  POItem_x0020_Title?: string;
  POItem_x0020_Value?: number;
  DueDate?: string;
  POAmount?: number;
  Currency?: string;
  Created?: string;
  Author?: { Title?: string };
  Modified?: string;
  Editor?: { Title?: string };
  PMCommentsHistory?: string;
  FinanceCommentsHistory?: string;
  AttachmentFiles?: any[];
  ParentPOID?: string;
  [key: string]: any;
}

interface InvoicePO {
  Id: number;
  POID: string;
  ProjectName?: string;
  POAmount?: number;
  Currency?: string;
  LineItemsJSON?: string;
  ParentPOID?: string;
  [key: string]: any;
}

interface POItem {
  POItem_x0020_Title: string;
  POItem_x0020_Value: number;
  POComments?: string;
  Currency?: string;
}

const getCurrencySymbol = (cur?: string) => {
  if (!cur) return "";
  try {
    return new Intl.NumberFormat("en-IN", {
      style: "currency",
      currency: cur,
      minimumFractionDigits: 0,
      maximumFractionDigits: 0,
    })
      .formatToParts(1)
      .find((part) => part.type === "currency")?.value || "";
  } catch {
    return "";
  }
};

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

const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";

export default function BusinessView({
  sp,
  context,
  projectsp,
}: BusinessViewProps) {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [poList, setPOList] = useState<InvoicePO[]>([]);
  const [reqList, setReqList] = useState<InvoiceRequest[]>([]);
  const [totals, setTotals] = useState<any>({});
  const [poPanel, setPoPanel] = useState<{
    open: boolean;
    po: InvoicePO | null;
    poItems: POItem[];
    invoiceRequests: InvoiceRequest[];
  }>({ open: false, po: null, poItems: [], invoiceRequests: [] });
  const [projects, setProjects] = useState<any[]>([]);
  const [departments, setDepartments] = useState<IDropdownOption[]>([]);
  const [selectedDepartment, setSelectedDepartment] = useState<string>("__all__");
  const [searchText, setSearchText] = useState<string>("");
  const [invoiceRequestPanel, setInvoiceRequestPanel] = useState<{
    open: boolean;
    invoiceRequest: InvoiceRequest | null;
  }>({ open: false, invoiceRequest: null });

  useEffect(() => {
    async function loadProjects() {
      const projectItems = await projectsp.web
        .lists.getByTitle("Projects")
        .items.select("Id", "Title", "Department")();
      setProjects(projectItems);

      // Extract unique departments
      const uniqueDeps = Array.from(
        new Set(projectItems.map((p) => p.Department).filter(Boolean))
      );

      // Map to dropdown options
      const options: IDropdownOption[] = [
        { key: "__all__", text: "All" },
        ...uniqueDeps.map((dep) => ({
          key: dep,
          text: dep,
        })),
      ];
      setDepartments(options);

    }
    loadProjects();
  }, [projectsp]);

  useEffect(() => {
    async function loadData() {
      setLoading(true);
      try {
        const POData: InvoicePO[] = await sp.web.lists.getByTitle("InvoicePO").items();
        const ReqData: InvoiceRequest[] = await sp.web
          .lists.getByTitle("Invoice Requests")
          .items.select(
            "Id,PurchaseOrder,ProjectName,Status,CurrentStatus,InvoiceAmount,POItem_x0020_Title,POItem_x0020_Value,DueDate,POAmount,Currency,Created,Author/Title,Modified,Editor/Title,PMCommentsHistory,FinanceCommentsHistory,AttachmentFiles"
          )
          .expand("Author", "Editor", "AttachmentFiles")();
        setPOList(POData);
        setReqList(ReqData);

        // Aggregates
        const totalPOValue = POData.reduce((s, p) => s + (+p.POAmount || 0), 0);
        const totalInvoiced = ReqData.filter(
          (r) => r.Status && r.Status.toLowerCase() !== "cancelled"
        ).reduce((s, r) => s + (+r.InvoiceAmount || 0), 0);

        const outstanding = totalPOValue - totalInvoiced;

        const paidPercent = totalPOValue > 0 ? (totalInvoiced / totalPOValue) * 100 : 0;

        const overdueCount = ReqData.filter(
          (r) => r.Status && r.Status.toLowerCase() === "pending payment"
        ).length;
        const statusMap: { [status: string]: number } = {};
        ReqData.forEach((r) => {
          const k = r.Status || "Unknown";
          statusMap[k] = (statusMap[k] || 0) + 1;
        });

        setTotals({
          totalPOValue,
          totalInvoiced,
          outstanding,
          paidPercent,
          overdueCount,
          statusMap,
        });
      } catch (err: any) {
        setError(err.message || String(err));
      }
      setLoading(false);
    }
    loadData();
  }, [sp]);

  const onDepartmentChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    setSelectedDepartment(option?.key as string);
  };

  // Columns for main PO summary DetailsList
  const columns: IColumn[] = [
    { key: "poid", name: "PO ID", fieldName: "POID", minWidth: 90, maxWidth: 140 },
    {
      key: "poamount",
      name: "PO Amount",
      fieldName: "POAmount",
      minWidth: 120,
      maxWidth: 150,
      onRender: (i) => getCurrencySymbol(i.Currency) + (+i.POAmount || 0).toLocaleString(),
    },
    { key: "project", name: "Project", fieldName: "ProjectName", minWidth: 120, maxWidth: 180 },
    {
      key: "invoiced",
      name: "Invoiced",
      fieldName: "invoiced",
      minWidth: 110,
      maxWidth: 130,
      onRender: (i) =>
        getCurrencySymbol(i.Currency) + ((+i.invoiced || 0).toLocaleString()),
    },
    {
      key: "percent",
      name: "Invoiced %",
      fieldName: "percentPaid",
      minWidth: 90,
      maxWidth: 100,
      onRender: (i) => ((i.percentPaid ?? 0).toFixed(0) + "%"),
    },
  ];

  function openInvoiceRequestPanel(request: InvoiceRequest) {
    setInvoiceRequestPanel({ open: true, invoiceRequest: request });
  }

  // Function to get PO items for a main PO - either from LineItemsJSON or children POs
  function getPOItems(po: InvoicePO, allPOs: InvoicePO[]): POItem[] {
    if (po.LineItemsJSON) {
      try {
        const items = JSON.parse(po.LineItemsJSON);
        if (Array.isArray(items)) {
          // Map to POItem structure safely
          return items.map((item: any) => ({
            POItem_x0020_Title: item.POItem_x0020_Title || item.Title || "-",
            POItem_x0020_Value: +item.POItem_x0020_Value || +item.Value || 0,
            POComments: item.POComments || item.Comments || "",
            Currency: po.Currency,
          }));
        }
      } catch {
        // Failed to parse JSON, fallback to empty array
      }
    }

    // If no LineItemsJSON, look for child POs by ParentPOID
    const childPOs = allPOs.filter((p) => p.ParentPOID === po.POID);
    if (childPOs.length > 0) {
      return childPOs.map((child) => ({
        POItem_x0020_Title: child.POID,
        POItem_x0020_Value: child.POAmount || 0,
        POComments: "",
        Currency: child.Currency || po.Currency,
      }));
    }

    // No line items or children, show main PO as single item
    return [
      {
        POItem_x0020_Title: po.POID,
        POItem_x0020_Value: po.POAmount || 0,
        POComments: "",
        Currency: po.Currency,
      },
    ];
  }

  // Calculate poSummary with invoiced values and percentages
  const poSummary = poList.map((po) => {
    const related = reqList.filter(
      (r) => r.PurchaseOrder === po.POID && r.Status && r.Status.toLowerCase() !== "cancelled"
    );
    const invoiced = related.reduce((s, r) => s + (+r.InvoiceAmount || 0), 0);
    return {
      ...po,
      invoiced,
      percentPaid: !po.POAmount ? 0 : (invoiced / +po.POAmount) * 100,
    };
  });

  // Filter by Department if selected
  const searchFilteredPoSummary = poSummary.filter(po => poMatchesSearch(po, searchText));

  const filteredPoSummary = (selectedDepartment && selectedDepartment !== "__all__"
    ? searchFilteredPoSummary.filter((po: any) => {
      const project = projects.find((p: any) => p.Title === po.ProjectName);
      return project?.Department === selectedDepartment;
    })
    : searchFilteredPoSummary
  ).filter(po => !po.ParentPOID);

  // Open PO Panel handler
  function openPoPanel(po: InvoicePO) {
    const poItems = getPOItems(po, poList);
    // Collect invoice requests for main PO and its child POs (if any)
    const poIdsForRequests = [po.POID];
    // Add child POIDs if any
    poList.forEach((p) => {
      if (p.ParentPOID === po.POID) poIdsForRequests.push(p.POID);
    });

    const invoiceRequests = reqList.filter((r) =>
      poIdsForRequests.includes(r.PurchaseOrder)
    );

    setPoPanel({ open: true, po, poItems, invoiceRequests });
  }

  function poMatchesSearch(po: any, text: string): boolean {
    if (!text) return true;
    const lower = text.toLowerCase();
    // Combine all searchable fields as strings, handle null/undefined
    return [
      po.POID,
      po.ProjectName,
      po.Currency,
      po.POAmount,
      po.invoiced,
      (po.percentPaid ?? 0) + "%",
    ]
      .map(val => (val == null ? "" : val.toString().toLowerCase()))
      .some(val => val.includes(lower));
  }

  // Panel Columns for PO Items table
  const poItemColumns: IColumn[] = [
    {
      key: "title",
      name: "PO Item Title",
      fieldName: "POItem_x0020_Title",
      minWidth: 150,
    },
    {
      key: "value",
      name: "PO Item Value",
      fieldName: "POItem_x0020_Value",
      minWidth: 120,
      onRender: (item: POItem) =>
        getCurrencySymbol(item.Currency) + (item.POItem_x0020_Value?.toLocaleString() ?? ""),
    },
    {
      key: "comments",
      name: "Comments",
      fieldName: "POComments",
      minWidth: 150,
      onRender: (item: POItem) => item.POComments || "-",
    },
  ];

  // Panel columns for Invoice Requests table
  const invoiceColumns: IColumn[] = [
    {
      key: "poitemtitle",
      name: "PO Item Title",
      fieldName: "POItem_x0020_Title",
      minWidth: 140,
    },
    {
      key: "poitemvalue",
      name: "PO Item Value",
      fieldName: "POItem_x0020_Value",
      minWidth: 120,
      onRender: (i) =>
        getCurrencySymbol(i.Currency) + (i.POItem_x0020_Value?.toLocaleString() ?? ""),
    },
    {
      key: "invoiceamount",
      name: "Invoiced Amount",
      fieldName: "InvoiceAmount",
      minWidth: 120,
      onRender: (i) =>
        getCurrencySymbol(i.Currency) + (i.InvoiceAmount?.toLocaleString() ?? ""),
    },
    { key: "status", name: "Invoice Status", fieldName: "Status", minWidth: 120 },
    {
      key: "created",
      name: "Created",
      fieldName: "Created",
      minWidth: 110,
      onRender: (i) => (i.Created ? new Date(i.Created).toLocaleDateString() : "-"),
    },
    {
      key: "author",
      name: "Created By",
      fieldName: "Author",
      minWidth: 120,
      onRender: (i) => i.Author?.Title ?? "-",
    },
  ];

  if (loading) return <Spinner label="Loading dashboard..." />;
  if (error) return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;

  return (
    <Stack
      tokens={{ childrenGap: 28 }}
      styles={{ root: { padding: 32, background: "#fafafa", minHeight: 600 } }}
    >
      <Separator />
      <Stack horizontal tokens={{ childrenGap: 24, padding: 0 }} styles={{ root: { marginBottom: 16 } }}>
        <SearchBox
          placeholder="Search"
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue ?? "")}
          styles={{ root: { maxWidth: 300 } }}
        />
        <Dropdown
          placeholder="Filter by Department"
          // label="Department"
          options={departments}
          selectedKey={selectedDepartment}
          onChange={onDepartmentChange}
          styles={{ root: { minWidth: 300 } }}
        />
      </Stack>
      <Stack>
        <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
          Invoice Status Breakdown
        </Text>
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          {Object.entries(totals.statusMap ?? {}).map(([status, count]) => (
            <Stack
              key={status}
              styles={{
                root: {
                  minWidth: 100,
                  background: "#fff",
                  borderRadius: 6,
                  boxShadow: "0 2px 7px #f6f6f6",
                  padding: "10px 14px",
                  margin: "6px 0",
                },
              }}
            >
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 700 } }}>
                {count}
              </Text>
              <div style={{ color: "#666" }}>{status}</div>
            </Stack>
          ))}
        </Stack>
      </Stack>
      <Separator />
      <Text variant="large" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
        All Purchase Orders (Summary)
      </Text>
      <DetailsList
        items={filteredPoSummary}
        columns={columns}
        compact
        isHeaderVisible
        styles={{ root: { background: "#fff", borderRadius: 8, marginTop: 4 } }}
        setKey="businessSummary"
        selectionMode={0}
        onActiveItemChanged={openPoPanel}
      />

      {/* PO Details Panel */}
      {/* <Panel
        isOpen={poPanel.open}
        onDismiss={() => setPoPanel({ open: false, po: null, poItems: [], invoiceRequests: [] })}
        headerText={`Invoice Details for PO: ${poPanel.po?.POID ?? ""}`}
        closeButtonAriaLabel="Close"
        isLightDismiss
        type={PanelType.largeFixed}
        styles={{ main: { maxWidth: 1000 } }}
      >
        {poPanel.po && (
          <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 10 } }}>
            <Stack horizontal tokens={{ childrenGap: 18 }} wrap>
              <Stack>
                <Text variant="mediumPlus">
                  <b>Purchase Order:</b> {poPanel.po.POID}
                </Text>
                <Text variant="mediumPlus">
                  <b>Project Name:</b> {poPanel.po.ProjectName}
                </Text>
                <Text variant="mediumPlus">
                  <b>PO Amount:</b> {getCurrencySymbol(poPanel.po.Currency)}
                  {(poPanel.po.POAmount ?? 0).toLocaleString()}
                </Text>
              </Stack>
            </Stack>

            <Separator />
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              PO Items
            </Text>
            <DetailsList
              items={poPanel.poItems}
              columns={poItemColumns}
              compact
              isHeaderVisible
            />

            <Separator />
            <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center" styles={{ root: { marginBottom: 8, marginTop: 10 } }}>
              <Text variant="mediumPlus">
                Invoice Requests for PO {poPanel.po.POID}
              </Text>
              {/* Add button if needed to show all */}
      {/* <PrimaryButton text="Show all Invoice Requests" /> */}
      {/* </Stack>

            {poPanel.invoiceRequests.length > 0 ? (
              <DetailsList
                items={poPanel.invoiceRequests}
                columns={invoiceColumns}
                compact
                isHeaderVisible
                onActiveItemChanged={openInvoiceRequestPanel}
              />

            ) : (
              <Text>No invoice requests for this POID</Text>
            )}
          </Stack>
        )}
      </Panel> */}

      <Panel
        isOpen={poPanel.open}
        onDismiss={() => setPoPanel({ open: false, po: null, poItems: [], invoiceRequests: [] })}
        headerText={poPanel.po ? `Invoice Details for PO: ${poPanel.po.POID}` : ""}
        closeButtonAriaLabel="Close"
        isLightDismiss
        type={PanelType.largeFixed}
        styles={{ main: { maxWidth: 1000 } }}
      >
        {poPanel.po && (
          <Stack tokens={{ childrenGap: 24 }} styles={{ root: { padding: 18, background: "#f4f9fc", borderRadius: 12, minHeight: 700 } }}>
            {/* PO Summary */}
            <div style={{
              background: "#fff",
              borderRadius: 8,
              padding: "18px 26px 10px",
              boxShadow: "0 1px 7px #ebecf0",
              marginBottom: 14
            }}>
              <Stack horizontal tokens={{ childrenGap: 50 }}>
                <Stack>
                  <Text variant="medium" styles={{ root: { marginTop: 6 } }}>
                    <span style={{ fontWeight: 600, color: primaryColor }}>Purchase Order:</span> {poPanel.po.POID}
                  </Text>
                  <Text variant="medium" styles={{ root: { marginTop: 6 } }}>
                    <span style={{ fontWeight: 600, color: primaryColor }}>Project Name:</span> {poPanel.po.ProjectName}
                  </Text>
                  <Text variant="medium" styles={{ root: { marginTop: 5 } }}>
                    <span style={{ fontWeight: 600, color: primaryColor }}>PO Amount:</span>
                    {" "}{getCurrencySymbol(poPanel.po.Currency)}
                    {(poPanel.po.POAmount ?? 0).toLocaleString()}
                  </Text>
                  {/* Add more summary fields as needed */}
                </Stack>
              </Stack>
            </div>

            <Text variant="large" styles={{ root: { fontWeight: 600, marginTop: 4, marginBottom: 0, color: primaryColor } }}>
              PO Items
            </Text>
            {/* PO Item List */}
            <div style={{
              background: "#fff",
              borderRadius: 7,
              padding: 18,
              marginBottom: 10,
              boxShadow: "0 1px 6px #f1f2f8"
            }}>
              <DetailsList
                items={poPanel.poItems}
                columns={poItemColumns}
                compact
                isHeaderVisible
                styles={{
                  root: { background: "transparent" }
                }}
              />
            </div>

            <Text variant="large" styles={{ root: { fontWeight: 600, marginTop: 4, color: primaryColor } }}>
              Invoice Requests for PO {poPanel.po.POID}
            </Text>
            {/* Invoice Requests */}
            <div style={{
              background: "#fff",
              borderRadius: 7,
              padding: 18,
              marginBottom: 2,
              boxShadow: "0 1px 6px #ecedfc"
            }}>
              {poPanel.invoiceRequests.length > 0 ? (
                <DetailsList
                  items={poPanel.invoiceRequests}
                  columns={invoiceColumns}
                  compact
                  isHeaderVisible
                  onActiveItemChanged={openInvoiceRequestPanel}
                  styles={{ root: { background: "transparent" } }}
                />
              ) : (
                <Text styles={{ root: { color: "#888" } }}>No invoice requests for this POID</Text>
              )}
            </div>
          </Stack>
        )}
      </Panel>

      <Panel
        isOpen={invoiceRequestPanel.open}
        onDismiss={() => setInvoiceRequestPanel({ open: false, invoiceRequest: null })}
        headerText={`Invoice Request Details:`}
        closeButtonAriaLabel="Close"
        isLightDismiss
        type={PanelType.medium}
        styles={{ main: { maxWidth: 620 } }}
      >
        {invoiceRequestPanel.invoiceRequest && (
          <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16, background: "#f4f9fc", borderRadius: 10 } }}>

            {/* Main Details Card */}
            <div style={{
              background: "#fff",
              borderRadius: 8,
              padding: "18px 22px",
              marginBottom: 12,
              boxShadow: "0 2px 8px #f2f2f7"
            }}>
              <div style={{
                display: "grid",
                gridTemplateColumns: "155px 1fr",
                rowGap: 14,
                columnGap: 22
              }}>
                <div style={{ fontWeight: 600, color: primaryColor }}>Purchase Order:</div>
                <div>{invoiceRequestPanel.invoiceRequest.PurchaseOrder}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Project Name:</div>
                <div>{invoiceRequestPanel.invoiceRequest.ProjectName ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>PO Item Title:</div>
                <div>{invoiceRequestPanel.invoiceRequest.POItem_x0020_Title ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>PO Item Value:</div>
                <div>
                  {getCurrencySymbol(invoiceRequestPanel.invoiceRequest.Currency)}
                  {invoiceRequestPanel.invoiceRequest.POItem_x0020_Value?.toLocaleString() ?? "-"}
                </div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Invoiced Amount:</div>
                <div>
                  {getCurrencySymbol(invoiceRequestPanel.invoiceRequest.Currency)}
                  {invoiceRequestPanel.invoiceRequest.InvoiceAmount?.toLocaleString() ?? "-"}
                </div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Invoice Status:</div>
                <div>
                  <span style={{
                    fontWeight: 700,
                    background: "#e5f1fa",
                    color: primaryColor,
                    borderRadius: 12,
                    padding: "2px 14px",
                    display: "inline-block"
                  }}>
                    {invoiceRequestPanel.invoiceRequest.Status ?? "-"}
                  </span>
                </div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Current Status:</div>
                <div>{invoiceRequestPanel.invoiceRequest.CurrentStatus ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Due Date:</div>
                <div>{invoiceRequestPanel.invoiceRequest.DueDate ? new Date(invoiceRequestPanel.invoiceRequest.DueDate).toLocaleDateString() : "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Created:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Created ? new Date(invoiceRequestPanel.invoiceRequest.Created).toLocaleDateString() : "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Created By:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Author?.Title ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Modified:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Modified ? new Date(invoiceRequestPanel.invoiceRequest.Modified).toLocaleDateString() : "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Modified By:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Editor?.Title ?? "-"}</div>
              </div>
            </div>

            <Separator />

            {/* PM Comments Section */}
            {invoiceRequestPanel.invoiceRequest.PMCommentsHistory && (
              <div style={{
                background: "#fcfcfd",
                borderRadius: 7,
                padding: 14,
                marginBottom: 10,
                border: "1px solid #edf3fa"
              }}>
                <Text variant="medium" styles={{ root: { fontWeight: 600, color: primaryColor, marginBottom: 4 } }}>PM Comments</Text>
                <pre style={{
                  whiteSpace: "pre-wrap",
                  maxHeight: 160,
                  overflowY: "auto",
                  backgroundColor: "#f6f9fd",
                  padding: 8,
                  borderRadius: 5,
                  fontFamily: "Segoe UI"
                }}>
                  {formatCommentHistory(invoiceRequestPanel.invoiceRequest.PMCommentsHistory)}
                </pre>
              </div>
            )}

            {/* Finance Comments Section */}
            {invoiceRequestPanel.invoiceRequest.FinanceCommentsHistory && (
              <div style={{
                background: "#fcfcfd",
                borderRadius: 7,
                padding: 14,
                marginBottom: 10,
                border: "1px solid #edf3fa"
              }}>
                <Text variant="medium" styles={{ root: { fontWeight: 600, color: primaryColor, marginBottom: 4 } }}>Finance Comments</Text>
                <pre style={{
                  whiteSpace: "pre-wrap",
                  maxHeight: 160,
                  overflowY: "auto",
                  backgroundColor: "#f6f9fd",
                  padding: 8,
                  borderRadius: 5,
                  fontFamily: "Segoe UI"
                }}>
                  {formatCommentHistory(invoiceRequestPanel.invoiceRequest.FinanceCommentsHistory)}
                </pre>
              </div>
            )}

            {/* Attachments Section */}
            <div style={{
              background: "#fcfcfd",
              borderRadius: 7,
              padding: 14,
              marginBottom: 2,
              border: "1px solid #edf3fa"
            }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600, color: primaryColor, marginBottom: 4 } }}>Attachments</Text>
              {invoiceRequestPanel.invoiceRequest.AttachmentFiles && invoiceRequestPanel.invoiceRequest.AttachmentFiles.length > 0 ? (
                <ul style={{ paddingLeft: 20, marginTop: 4 }}>
                  {invoiceRequestPanel.invoiceRequest.AttachmentFiles.map((file: any) => (
                    <li key={file.UniqueId} style={{ marginBottom: 7 }}>
                      <a href={file.ServerRelativeUrl} target="_blank" rel="noopener noreferrer"
                        style={{
                          color: primaryColor,
                          textDecoration: "underline",
                          fontWeight: 500,
                          fontSize: 15
                        }}>
                        {file.FileName}
                      </a>
                    </li>
                  ))}
                </ul>
              ) : (
                <Text styles={{ root: { color: "#888", marginTop: 3 } }}>No attachments</Text>
              )}
            </div>

          </Stack>
        )}
      </Panel>

    </Stack>
  );
}
