// import * as React from "react";
// import {
//   Stack,
//   Text,
//   Spinner,
//   MessageBar,
//   MessageBarType,
//   Icon,
//   Dropdown,
//   IDropdownOption,
//   PrimaryButton,
// } from "@fluentui/react";
// import {
//   ResponsiveContainer,
//   BarChart,
//   Bar,
//   Tooltip,
//   Legend,
//   CartesianGrid,
//   PieChart,
//   Pie,
//   Cell,
//   LabelList,
// } from "recharts";
// import { SPFI } from "@pnp/sp";
// import { XAxis as _XAxis, YAxis as _YAxis } from "recharts";
// const XAxis = _XAxis as any;
// const YAxis = _YAxis as any;
// import * as XLSX from "xlsx";
// import { saveAs } from "file-saver";

// interface DashboardProps {
//   sp: SPFI;
//   context: any;
//   projectsp: SPFI;
// }
// interface AgingRow {
//   projectName: string;
//   poid: string;
//   invoiceNumber: string;
//   dueDate: string;
//   status: string;
//   amount: number;
//   currency?: string;
// }
// interface DashboardState {
//   loading: boolean;
//   error: string | null;
//   totalInvoices: number;
//   pendingInvoices: number;
//   paidInvoices: number;
//   monthlyData: { month: string; amount: number }[];
//   statusData: { name: string; value: number }[];
//   projectComparisonData: { project: string; poAmount: number; invoiceAmount: number }[];
//   currencyData: { name: string; value: number }[];
//   agingData: AgingRow[];
//   period: string;
//   startDate: Date | null;
//   endDate: Date | null;
//   business: string;
//   businessUnit: string;
//   customer: string;
//   project: string;
//   businessOptions: IDropdownOption[];
//   businessUnitOptions: IDropdownOption[];
//   customerOptions: IDropdownOption[];
//   projectOptions: IDropdownOption[];
//   statusFilter: string;
//   currentStatusFilter: string;
//   monthlyProjectComparisonData: { month: string; invoiceAmount: number; poAmount: number; paidAmount: number }[];
//   poStatusStackedData: any[];
//   currentStatusPieData?: { name: string; value: number }[];
//   overdueDataCount: Record<string, number[]>;
//   overdueDataAmount: Record<string, number[]>;
//   showOverdueBy: "count" | "amount";
//   dueDateBucketFilter: number | null,
//   currentStatusOptions?: IDropdownOption[];
//   statusOptions?: IDropdownOption[];
// }

// const periodOptions: IDropdownOption[] = [
//   { key: "all", text: "All" },
//   { key: "week_to_date", text: "Week till date" },
//   { key: "last_week", text: "Last Week" },
//   { key: "month_to_date", text: "Month till date" },
//   { key: "current_month", text: "Current Month" },
//   { key: "last_month", text: "Last Month" },
//   { key: "year_to_date", text: "Year till date" },
//   { key: "current_year", text: "Current Year" },
//   { key: "last_year", text: "Last Year" },
// ];
// // const statusOptions: IDropdownOption[] = [
// //   { key: "", text: "All Statuses" },
// //   { key: "Pending", text: "Pending" },
// //   { key: "Paid", text: "Paid" },
// //   { key: "Others", text: "Others" },
// // ];
// // const currentStatusOptions: IDropdownOption[] = [
// //   { key: "", text: "All Current Statuses" },
// //   { key: "Payment Received", text: "Payment Received" },
// //   { key: "Pending Payment", text: "Pending Payment" },
// //   { key: "Not Generated", text: "Not Generated" },
// //   { key: "Invoice Raised", text: "Invoice Raised" },
// // ];
// const pieColors = ["#0078d4", "#ffaa44", "#107c10", "#d83b01", "#605e5c", "#fcbd73", "#a1a7b3"];
// const thStyle: React.CSSProperties = {
//   padding: "12px 16px",
//   background: "#f0f2f7",
//   color: "#222",
//   textAlign: "left",
//   fontWeight: 700,
//   borderBottom: "2px solid #dde0eb",
// };
// const tdStyle: React.CSSProperties = {
//   padding: "12px 16px",
//   color: "#333",
//   verticalAlign: "middle",
//   borderBottom: "1px solid #eee",
// };

// function getChartData(state: DashboardState) {
//   const source = state.showOverdueBy === "count" ? state.overdueDataCount : state.overdueDataAmount;
//   return Object.entries(source)
//     .map(([project, values]) => ({
//       name: project,
//       value:
//         state.dueDateBucketFilter === null
//           ? values.reduce((a, b) => a + b, 0) // sum all buckets if All Periods
//           : values[state.dueDateBucketFilter!], // or pick bucket
//     }))
//     .filter(d => d.value > 0);
// }

// function getCurrencySymbol(currencyCode: string, locale: string = "en-US") {
//   if (!currencyCode || !currencyCode.trim()) {
//     // default if missing
//     return "USD";
//   }
//   try {
//     return new Intl.NumberFormat(locale, {
//       style: "currency",
//       currency: currencyCode,
//       minimumFractionDigits: 0,
//       maximumFractionDigits: 0,
//     })
//       .formatToParts(1)
//       .find((part) => part.type === "currency")?.value ?? currencyCode;
//   } catch (error) {
//     console.warn("Invalid currency code", currencyCode, error);
//     return currencyCode;
//   }
// }

// // Utility for period range
// function getPeriodRange(period: string) {
//   const now = new Date();
//   switch (period) {
//     case "week_to_date": {
//       const first = new Date(now);
//       first.setDate(now.getDate() - now.getDay());
//       return { start: first, end: now };
//     }
//     case "last_week": {
//       const first = new Date(now);
//       first.setDate(now.getDate() - now.getDay() - 7);
//       const last = new Date(first);
//       last.setDate(first.getDate() + 6);
//       return { start: first, end: last };
//     }
//     case "month_to_date":
//       return { start: new Date(now.getFullYear(), now.getMonth(), 1), end: now };
//     case "current_month":
//       return {
//         start: new Date(now.getFullYear(), now.getMonth(), 1),
//         end: new Date(now.getFullYear(), now.getMonth() + 1, 0),
//       };
//     case "last_month":
//       return {
//         start: new Date(now.getFullYear(), now.getMonth() - 1, 1),
//         end: new Date(now.getFullYear(), now.getMonth(), 0),
//       };
//     case "year_to_date":
//       return { start: new Date(now.getFullYear(), 0, 1), end: now };
//     case "current_year":
//       return { start: new Date(now.getFullYear(), 0, 1), end: new Date(now.getFullYear(), 11, 31) };
//     case "last_year":
//       return {
//         start: new Date(now.getFullYear() - 1, 0, 1),
//         end: new Date(now.getFullYear() - 1, 11, 31),
//       };
//     default:
//       return { start: null, end: null };
//   }
// }
// function inDateRange(dateStr: string | undefined, start: Date | null, end: Date | null) {
//   if (!dateStr || dateStr === "") return true;
//   const date = new Date(dateStr);
//   if (start && date < start) return false;
//   if (end && date > end) return false;
//   return true;
// }

// // Overdue buckets for overdue breakdown
// interface OverdueBucket {
//   label: string;
//   minDays: number;
//   maxDays: number | null; // null means no upper bound
// }

// const overdueBuckets: OverdueBucket[] = [
//   { label: "0-30 days", minDays: 0, maxDays: 30 },
//   { label: "31-60 days", minDays: 31, maxDays: 60 },
//   { label: "61-90 days", minDays: 61, maxDays: 90 },
//   { label: "90+ days", minDays: 91, maxDays: null },
// ];

// function parseDMY(d: string) {
//   const [day, month, year] = d.split("/").map(Number);
//   return new Date(year, month - 1, day);
// }

// function getOverdueDays(dueDateStr: string) {
//   if (!dueDateStr || dueDateStr === "") return null;
//   let dueDate: Date;
//   // Try detect "DD/MM/YYYY" pattern, else fallback
//   if (/^\d{2}\/\d{2}\/\d{4}$/.test(dueDateStr)) {
//     dueDate = parseDMY(dueDateStr);
//   } else {
//     dueDate = new Date(dueDateStr);
//   }
//   if (isNaN(dueDate.getTime())) return null;
//   const now = new Date();
//   const diff = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 3600 * 24));
//   return diff > 0 ? diff : 0;
// }

// function aggregateOverdueData(invoices: AgingRow[]) {
//   const aggregation: Record<string, { count: number[]; amount: number[] }> = {};
//   invoices.forEach((inv) => {
//     const overdueDays = getOverdueDays(inv.dueDate);
//     if (overdueDays === null || overdueDays === 0) return; // Not overdue

//     let bucketIndex = overdueBuckets.findIndex((b) => {
//       if (b.maxDays === null) return overdueDays >= b.minDays;
//       return overdueDays >= b.minDays && overdueDays <= b.maxDays;
//     });
//     if (bucketIndex === -1) bucketIndex = overdueBuckets.length - 1;

//     const key = inv.projectName;
//     if (!aggregation[key]) {
//       aggregation[key] = {
//         count: Array(overdueBuckets.length).fill(0),
//         amount: Array(overdueBuckets.length).fill(0),
//       };
//     }
//     aggregation[key].count[bucketIndex]++;
//     aggregation[key].amount[bucketIndex] += inv.amount;
//   });

//   return aggregation;
// }

// // Monthly Project PO/Invoice/Paid Amounts by month (month-wise only)
// function getMonthlyInvoicePaidComparison(
//   invoiceItems: any[],
//   poItems: any[],
//   startDate: Date | null,
//   endDate: Date | null
// ) {
//   const raw: Record<string, { month: string; invoiceAmount: number; poAmount: number; paidAmount: number }> = {};

//   invoiceItems.forEach((i) => {
//     const date = i.Created ? new Date(i.Created) : null;
//     if (!date) return;
//     if (startDate && date < startDate) return;
//     if (endDate && date > endDate) return;

//     const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
//     if (!raw[key]) raw[key] = { month: key, invoiceAmount: 0, poAmount: 0, paidAmount: 0 };

//     raw[key].invoiceAmount += i.InvoiceAmount || 0;
//     if (i.Status === "Payment Received") {
//       raw[key].paidAmount += i.InvoiceAmount || 0;
//     }
//   });

//   poItems.forEach((po) => {
//     const date = po.Created ? new Date(po.Created) : null;
//     if (!date) return;
//     if (startDate && date < startDate) return;
//     if (endDate && date > endDate) return;

//     const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
//     if (!raw[key]) raw[key] = { month: key, invoiceAmount: 0, poAmount: 0, paidAmount: 0 };

//     raw[key].poAmount += po.POAmount || 0;
//   });

//   // Sort months chronologically
//   return Object.values(raw).sort((a, b) => {
//     const [am, ay] = a.month.split(" ");
//     const [bm, by] = b.month.split(" ");
//     return new Date(`${am} 1, ${ay}`).getTime() - new Date(`${bm} 1, ${by}`).getTime();
//   });
// }

// function getPOStatusStackedData(
//   filteredInvoiceItems: any[],
//   poItems: any[],
//   projectToBusiness: any,
//   projectToBusinessUnit: any,
//   selectedBusiness: string,
//   selectedBusinessUnit: string
// ) {
//   return poItems
//     .filter((po) => {
//       const project = po.ProjectName;
//       const biz = projectToBusiness[project];
//       const bu = projectToBusinessUnit[project];
//       return (!selectedBusiness || biz === selectedBusiness) && (!selectedBusinessUnit || bu === selectedBusinessUnit);
//     })
//     .map((po) => {
//       const project = po.ProjectName;
//       const invoicesForProject = filteredInvoiceItems.filter((i) => i.ProjectName === project);
//       const invoiced = invoicesForProject.reduce((sum, i) => sum + (i.InvoiceAmount || 0), 0);
//       const invoicePaid = invoicesForProject
//         .filter((i) => i.Status === "Payment Received")
//         .reduce((sum, i) => sum + (i.InvoiceAmount || 0), 0);
//       // const invoicePending = invoicesForProject
//       //   .filter((i) => i.Status === "Pending Payment")
//       //   .reduce((sum, i) => sum + (i.InvoiceAmount || 0), 0);
//       const invoicePending = invoiced - invoicePaid;
//       const notInvoiced = (po.POAmount || 0) - invoiced;
//       return {
//         project,
//         POAmount: po.POAmount || 0,
//         Invoiced: invoiced,
//         Paid: invoicePaid,
//         Pending: invoicePending,
//         // Others: invoiceOther,
//         NotInvoiced: notInvoiced,
//       };
//     });
// }

// export default function Dashboard({ sp, projectsp, context }: DashboardProps) {
//   const [state, setState] = React.useState<DashboardState>({
//     loading: true,
//     error: null,
//     totalInvoices: 0,
//     pendingInvoices: 0,
//     paidInvoices: 0,
//     monthlyData: [],
//     statusData: [],
//     projectComparisonData: [],
//     currencyData: [],
//     agingData: [],
//     statusFilter: "",
//     currentStatusFilter: "",
//     period: "all",
//     startDate: null,
//     endDate: null,
//     business: "",
//     businessUnit: "",
//     customer: "",
//     project: "",
//     businessOptions: [],
//     businessUnitOptions: [],
//     customerOptions: [],
//     projectOptions: [],
//     monthlyProjectComparisonData: [],
//     poStatusStackedData: [],
//     currentStatusPieData: [],
//     overdueDataCount: {},
//     overdueDataAmount: {},
//     showOverdueBy: "count",
//     dueDateBucketFilter: null,
//   });
//   const [rawInvoiceItems, setRawInvoiceItems] = React.useState<any[]>([]);
//   const [rawPoItems, setRawPoItems] = React.useState<any[]>([]);
//   const [projectToBusiness, setProjectToBusiness] = React.useState<{ [key: string]: string }>({});
//   const [projectToBusinessUnit, setProjectToBusinessUnit] = React.useState<{ [key: string]: string }>({});
//   React.useEffect(() => {
//     loadDashboardData();
//   }, []);

//   async function loadDashboardData() {
//     try {
//       setState((prev) => ({ ...prev, loading: true, error: null }));
//       const projectBiz: { [key: string]: string } = {};
//       const projectBU: { [key: string]: string } = {};
//       const invoiceStatusSet = new Set<string>();
//       const currentStatusSet = new Set<string>();
//       setProjectToBusiness(projectBiz);
//       setProjectToBusinessUnit(projectBU);

//       const invoiceItems = await sp.web.lists
//         .getByTitle("Invoice Requests")
//         .items.select(
//           "Id",
//           "Status",
//           "InvoiceAmount",
//           "Created",
//           "InvoiceNumber",
//           "DueDate",
//           "ProjectName",
//           "Currency",
//           "PurchaseOrder",
//           "POItem_x0020_Value",
//           "POItem_x0020_Title"
//         )();
//       const poItems = await sp.web.lists
//         .getByTitle("InvoicePO")
//         .items.select("Id", "POAmount", "ProjectName", "Currency", "Customer")();
//       setRawInvoiceItems(invoiceItems);
//       setRawPoItems(poItems);

//       const projectOpts = Array.from(
//         new Set([...invoiceItems, ...poItems].map((i) => i.ProjectName).filter(Boolean))
//       ).map((b) => ({ key: b, text: b }));
//       const customerOpts = Array.from(new Set(poItems.map((p) => p.Customer).filter(Boolean))).map((b) => ({
//         key: b,
//         text: b,
//       }));

//       invoiceItems.forEach((item) => {
//         if (item.Status) invoiceStatusSet.add(item.Status);
//         if (item.CurrentStatus) currentStatusSet.add(item.CurrentStatus);
//       });

//       // Convert sets to sorted arrays
//       const sortedStatusOptions = Array.from(invoiceStatusSet).sort().map(s => ({ key: s, text: s }));
//       const sortedCurrentStatusOptions = Array.from(currentStatusSet).sort().map(s => ({ key: s, text: s }));

//       sortedStatusOptions.unshift({ key: "", text: "All Statuses" });
//       sortedCurrentStatusOptions.unshift({ key: "", text: "All Current Statuses" });

//       setState((prev) => ({
//         ...prev,
//         projectOptions: [{ key: "", text: "All Projects" }, ...projectOpts],
//         customerOptions: [{ key: "", text: "All Customers" }, ...customerOpts],
//         statusOptions: sortedStatusOptions,
//         currentStatusOptions: sortedCurrentStatusOptions,
//       }));

//       processDashboardData(invoiceItems, poItems, projectBiz, projectBU, "", "", "all", null, null, "", "", "", "", false);
//     } catch (err: any) {
//       setState((prev) => ({ ...prev, loading: false, error: err.message || "Error loading data" }));
//     }
//   }

//   function exportTableToExcel() {
//     // Prepare data for sheet: headers + rows
//     const headers = ["Invoice No", "Project Name", "POID", "Due Date", "Invoice Status", "Invoiced Amount"];
//     const data = state.agingData.map(row => [
//       row.invoiceNumber,
//       row.projectName,
//       row.poid,
//       row.dueDate,
//       row.status,
//       row.currency ? row.currency + " " + row.amount.toLocaleString() : row.amount.toLocaleString(),
//     ]);
//     const worksheetData = [headers, ...data];

//     // Create worksheet and workbook
//     const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
//     const workbook = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(workbook, worksheet, "Invoice Requests");

//     // Generate buffer
//     const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

//     // Save file
//     const blob = new Blob([wbout], { type: "application/octet-stream" });
//     saveAs(blob, "invoice_requests.xlsx");
//   }


//   function processDashboardData(
//     invoiceItems: any[],
//     poItems: any[],
//     projectToBusinessMap: { [k: string]: string },
//     projectToBusinessUnitMap: { [k: string]: string },
//     statusFilter: string,
//     currentStatusFilter: string,
//     period: string,
//     startDate: Date | null,
//     endDate: Date | null,
//     business: string,
//     businessUnit: string,
//     customer: string,
//     project: string,
//     updateOnly?: boolean
//   ) {
//     let actualStart = startDate,
//       actualEnd = endDate;
//     if (period !== "all") {
//       const { start, end } = getPeriodRange(period);
//       actualStart = start;
//       actualEnd = end;
//     }
//     let filteredInvoiceItems = invoiceItems.filter((inv) => {
//       const projectName = inv.ProjectName;
//       const projectBusiness = projectToBusinessMap[projectName] || "";
//       const projectBusinessUnit = projectToBusinessUnitMap[projectName] || "";
//       const poRow = poItems.find((po) => po.ProjectName === projectName);
//       const poCustomer = poRow?.Customer || "";
//       let statusIncluded = true;
//       if (statusFilter === "Pending")
//         statusIncluded =
//           inv.Status &&
//           (inv.Status.toLowerCase().includes("pending") ||
//             inv.Status.toLowerCase().includes("invoice requested") ||
//             inv.Status.toLowerCase().includes("invoice raised"));
//       else if (statusFilter === "Paid") statusIncluded = inv.Status && inv.Status.toLowerCase().includes("received");
//       else if (statusFilter === "Others")
//         statusIncluded =
//           inv.Status &&
//           !(
//             inv.Status.toLowerCase().includes("pending") ||
//             inv.Status.toLowerCase().includes("received") ||
//             inv.Status.toLowerCase().includes("invoice requested") ||
//             inv.Status.toLowerCase().includes("invoice raised")
//           );
//       let currentStatusIncluded = currentStatusFilter ? inv.Status === currentStatusFilter : true;
//       let dateIncluded = true;
//       if (actualStart || actualEnd) {
//         const baseDate = inv.DueDate ? new Date(inv.DueDate) : inv.Created ? new Date(inv.Created) : null;
//         dateIncluded = baseDate ? inDateRange(baseDate.toISOString(), actualStart, actualEnd) : true;
//       }
//       let bizIncluded = business ? projectBusiness === business : true;
//       let unitIncluded = businessUnit ? projectBusinessUnit === businessUnit : true;
//       let custIncluded = customer ? poCustomer === customer : true;
//       let projectIncluded = project ? projectName === project : true;
//       return statusIncluded && currentStatusIncluded && dateIncluded && bizIncluded && unitIncluded && custIncluded && projectIncluded;
//     });
//     let filteredPOItems = poItems.filter((po) => {
//       const projectName = po.ProjectName;
//       const projectBusiness = projectToBusinessMap[projectName] || "";
//       const projectBusinessUnit = projectToBusinessUnitMap[projectName] || "";
//       let bizIncluded = business ? projectBusiness === business : true;
//       let unitIncluded = businessUnit ? projectBusinessUnit === businessUnit : true;
//       let custIncluded = customer ? po.Customer === customer : true;
//       let projectIncluded = project ? projectName === project : true;
//       return bizIncluded && unitIncluded && custIncluded && projectIncluded;
//     });

//     // Map POID to Currency
//     const poCurrencyMap: { [key: string]: string } = {};
//     filteredPOItems.forEach((po) => (poCurrencyMap[po.Id] = po.Currency || ""));

//     const totalInvoices = filteredInvoiceItems.length;
//     const paidInvoices = filteredInvoiceItems.filter((i) => i.Status?.toLowerCase().includes("received")).length;
//     const pendingInvoices = filteredInvoiceItems.filter(
//       (i) =>
//         i.Status?.toLowerCase().includes("pending") ||
//         i.Status?.toLowerCase().includes("invoice requested") ||
//         i.Status?.toLowerCase().includes("invoice raised")
//     ).length;

//     const monthlyMap: Record<string, number> = {};
//     filteredInvoiceItems.forEach((i) => {
//       if (!i.Created) return;
//       const date = new Date(i.Created);
//       const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
//       monthlyMap[key] = (monthlyMap[key] || 0) + (i.InvoiceAmount || 0);
//     });
//     const monthlyData = Object.keys(monthlyMap)
//       .map((m) => ({ month: m, amount: monthlyMap[m] }))
//       .sort((a, b) => new Date(`1 ${a.month}`).getTime() - new Date(`1 ${b.month}`).getTime());

//     const projectMap: Record<string, { poAmount: number; invoiceAmount: number }> = {};
//     filteredPOItems.forEach((po) => {
//       const project = po.ProjectName;
//       if (!projectMap[project]) projectMap[project] = { poAmount: 0, invoiceAmount: 0 };
//       projectMap[project].poAmount += po.POAmount || 0;
//     });
//     filteredInvoiceItems.forEach((inv) => {
//       const project = inv.ProjectName;
//       if (!projectMap[project]) projectMap[project] = { poAmount: 0, invoiceAmount: 0 };
//       projectMap[project].invoiceAmount += inv.InvoiceAmount || 0;
//     });
//     const projectComparisonData = Object.keys(projectMap).map((project) => ({
//       project,
//       poAmount: projectMap[project].poAmount,
//       invoiceAmount: projectMap[project].invoiceAmount,
//     }));
//     const currencyMap: Record<string, number> = {};
//     filteredInvoiceItems.forEach((inv) => {
//       const currency = inv.Currency;
//       currencyMap[currency] = (currencyMap[currency] || 0) + (inv.InvoiceAmount || 0);
//     });
//     const currencyData = Object.keys(currencyMap).map((c) => ({ name: c, value: currencyMap[c] }));

//     // Add currency for each row (tries by PO first, then Invoice currency)
//     const agingData: AgingRow[] = filteredInvoiceItems.map((inv) => {
//       let poMatch = filteredPOItems.find((po) => po.ProjectName === inv.ProjectName);
//       return {
//         projectName: inv.ProjectName || "",
//         poid: inv.PurchaseOrder || "",
//         invoiceNumber: inv.InvoiceNumber || "",
//         dueDate: inv.DueDate ? new Date(inv.DueDate).toLocaleDateString() : "",
//         status: inv.Status || "",
//         amount: inv.InvoiceAmount || 0,
//         currency: poMatch?.Currency || inv.Currency || "",
//       };
//     });

//     // Aggregate overdue data
//     const overdueAggregation = aggregateOverdueData(agingData);
//     const overdueDataCount: Record<string, number[]> = {};
//     const overdueDataAmount: Record<string, number[]> = {};
//     Object.entries(overdueAggregation).forEach(([key, val]) => {
//       overdueDataCount[key] = val.count;
//       overdueDataAmount[key] = val.amount;
//     });

//     // --- Pie data for Current Status --- //
//     const currentStatusCount: Record<string, number> = {};
//     filteredInvoiceItems.forEach((inv) => {
//       const stat = inv.Status;
//       currentStatusCount[stat] = (currentStatusCount[stat] || 0) + 1;
//     });
//     const currentStatusPieData = Object.keys(currentStatusCount).map((s) => ({
//       name: s,
//       value: currentStatusCount[s],
//     }));

//     // The new monthly PO vs Invoice/Paid data:
//     const monthlyProjectComparisonData = getMonthlyInvoicePaidComparison(filteredInvoiceItems, filteredPOItems, actualStart, actualEnd);
//     const poStatusStackedData = getPOStatusStackedData(filteredInvoiceItems, filteredPOItems, projectToBusinessMap, projectToBusinessUnitMap, business, businessUnit);

//     setState((prev) => ({
//       ...prev,
//       loading: false,
//       error: null,
//       totalInvoices,
//       pendingInvoices,
//       paidInvoices,
//       monthlyData,
//       statusData: [
//         { name: "Pending", value: pendingInvoices },
//         { name: "Paid", value: paidInvoices },
//         { name: "Others", value: totalInvoices - (pendingInvoices + paidInvoices) },
//       ],
//       projectComparisonData,
//       currencyData,
//       agingData,
//       statusFilter,
//       currentStatusFilter,
//       period,
//       startDate: actualStart,
//       endDate: actualEnd,
//       business,
//       businessUnit,
//       customer,
//       project,
//       monthlyProjectComparisonData,
//       poStatusStackedData,
//       currentStatusPieData,
//       overdueDataCount,
//       overdueDataAmount,
//     }));
//   }

//   React.useEffect(() => {
//     if (rawInvoiceItems.length > 0 && rawPoItems.length > 0) {
//       processDashboardData(
//         rawInvoiceItems,
//         rawPoItems,
//         projectToBusiness,
//         projectToBusinessUnit,
//         state.statusFilter,
//         state.currentStatusFilter,
//         state.period,
//         state.startDate,
//         state.endDate,
//         state.business,
//         state.businessUnit,
//         state.customer,
//         state.project,
//         true
//       );
//     }
//     // eslint-disable-next-line
//   }, [
//     state.statusFilter,
//     state.currentStatusFilter,
//     state.period,
//     state.startDate,
//     state.endDate,
//     state.business,
//     state.businessUnit,
//     state.customer,
//     state.project,
//     projectToBusiness,
//     projectToBusinessUnit,
//     rawInvoiceItems,
//     rawPoItems,
//   ]);

//   if (state.loading) {
//     return (
//       <Stack horizontalAlign="center" verticalAlign="center" style={{ height: "80vh" }}>
//         <Spinner label="Loading Dashboard..." />
//       </Stack>
//     );
//   }
//   if (state.error) {
//     return <MessageBar messageBarType={MessageBarType.error}>{state.error}</MessageBar>;
//   }
//   return (
//     <div style={{ padding: 20 }}>
//       <Text variant="xxLarge" styles={{ root: { fontWeight: 600, marginBottom: 20 } }}>
//         Invoice Dashboard
//       </Text>
//       <Stack horizontal tokens={{ childrenGap: 18 }} style={{ marginBottom: 18, flexWrap: "wrap" }}>
//         <Dropdown
//           label="Period"
//           options={periodOptions}
//           selectedKey={state.period}
//           onChange={(_, o) => {
//             const { start, end } = getPeriodRange(o?.key as string);
//             setState((prev) => ({ ...prev, period: o?.key as string, startDate: start, endDate: end }));
//           }}
//           style={{ minWidth: 150 }}
//         />
//         <Dropdown
//           label="Project"
//           options={state.projectOptions}
//           selectedKey={state.project}
//           onChange={(_, o) => setState((prev) => ({ ...prev, project: o?.key as string }))}
//           style={{ minWidth: 180 }}
//         />
//         <Dropdown
//           label="Customer"
//           options={state.customerOptions}
//           selectedKey={state.customer}
//           onChange={(_, o) => setState((prev) => ({ ...prev, customer: o?.key as string }))}
//           style={{ minWidth: 160 }}
//         />
//         <Dropdown
//           label="Current Status"
//           options={state.currentStatusOptions}
//           selectedKey={state.currentStatusFilter}
//           onChange={(_, o) => setState((prev) => ({ ...prev, currentStatusFilter: o?.key as string }))}
//           style={{ minWidth: 170 }}
//         />
//         <Dropdown
//           label="Invoice Status"
//           options={state.statusOptions}
//           selectedKey={state.statusFilter}
//           onChange={(_, o) => setState((prev) => ({ ...prev, statusFilter: o?.key as string }))}
//           style={{ minWidth: 150 }}
//         />

//       </Stack>
//       <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 20 }} wrap>
//         <DashboardCard title="Total Invoices" value={state.totalInvoices} icon="NumberField" color="#0078d4" />
//         <DashboardCard title="Pending" value={state.pendingInvoices} icon="Clock" color="#ffaa44" />
//         <DashboardCard title="Paid" value={state.paidInvoices} icon="CheckMark" color="#107c10" />
//       </Stack>
//       <Stack horizontal tokens={{ childrenGap: 24 }} wrap styles={{ root: { marginTop: 20, marginBottom: 32 } }}>
//         <ChartContainer title="Invoices by Status">
//           <ResponsiveContainer width="100%" height={320}>
//             <PieChart>
//               <Pie
//                 dataKey="value"
//                 data={state.currentStatusPieData}
//                 cx="50%"
//                 cy="50%"
//                 outerRadius={100}
//                 label={({ name, value }) => `${name}: ${value}`}
//               >
//                 {(state.currentStatusPieData || []).map((entry, idx) => (
//                   <Cell key={entry.name} fill={pieColors[idx % pieColors.length]} />
//                 ))}
//               </Pie>
//               <Tooltip />
//               <Legend />
//             </PieChart>
//           </ResponsiveContainer>
//         </ChartContainer>
//         {/* <ChartContainer title="Monthly PO vs Invoice by Project">
//           <ResponsiveContainer width="100%" height="90%">
//             <BarChart data={state.monthlyProjectComparisonData} margin={{ top: 10, right: 20, left: 0, bottom: 40 }} barCategoryGap="25%">
//               <CartesianGrid strokeDasharray="3 3" />
//               <XAxis dataKey="month" angle={-25} textAnchor="end" interval={0} />
//               <YAxis />
//               <Tooltip />
//               <Legend />
//               <Bar dataKey="poAmount" name="PO Amount" fill="#0078d4" />
//               <Bar dataKey="invoiceAmount" name="Invoiced Amount" fill="#107c10" />
//               <Bar dataKey="paidAmount" name="Paid Amount" fill="#ffaa44" />
//             </BarChart>
//           </ResponsiveContainer>
//         </ChartContainer> */}
//         <ChartContainer title="Monthly PO vs Invoice by Project">
//           <ResponsiveContainer width="100%" height="90%">
//             <BarChart
//               data={state.monthlyProjectComparisonData}
//               margin={{ top: 10, right: 20, left: 0, bottom: 40 }}
//               barCategoryGap="25%"
//             >
//               <CartesianGrid strokeDasharray="3 3" />
//               <XAxis dataKey="month" angle={-25} textAnchor="end" interval={0} />
//               <YAxis />
//               <Tooltip />
//               <Legend />
//               <Bar dataKey="poAmount" name="PO Amount" fill="#0078d4">
//                 <LabelList dataKey="poAmount" position="top" formatter={(v: number) => v.toLocaleString()} />
//               </Bar>
//               <Bar dataKey="invoiceAmount" name="Invoiced Amount" fill="#107c10">
//                 <LabelList dataKey="invoiceAmount" position="top" formatter={(v: number) => v.toLocaleString()} />
//               </Bar>
//               <Bar dataKey="paidAmount" name="Paid Amount" fill="#ffaa44">
//                 <LabelList dataKey="paidAmount" position="top" formatter={(v: number) => v.toLocaleString()} />
//               </Bar>
//             </BarChart>
//           </ResponsiveContainer>
//         </ChartContainer>
//       </Stack>
//       <Stack
//         styles={{
//           root: { background: "#fff", borderRadius: 12, padding: 16, marginTop: 32, marginBottom: 24, boxShadow: "0 2px 8px rgba(0,0,0,0.1)" }
//         }}
//       >
//         <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 12 } }}>
//           Overdue Invoices Breakdown ({state.showOverdueBy === "count" ? "Count" : "Amount"})
//         </Text>
//         <Stack horizontal tokens={{ childrenGap: 14 }} styles={{ root: { marginBottom: 12 } }}>
//           <Dropdown
//             label="Show by"
//             selectedKey={state.showOverdueBy}
//             options={[{ key: "count", text: "Count" }, { key: "amount", text: "Amount" }]}
//             onChange={(_, option) => setState(prev => ({ ...prev, showOverdueBy: option?.key as "count" | "amount" }))}
//             styles={{ root: { width: 140 } }}
//           />
//           <Dropdown
//             label="Due Date Period"
//             selectedKey={state.dueDateBucketFilter !== null ? state.dueDateBucketFilter : "all"}
//             options={[
//               { key: "all", text: "All Periods" },
//               ...overdueBuckets.map((b, idx) => ({ key: idx, text: b.label })),
//             ]}
//             onChange={(_, option) => setState(prev => ({
//               ...prev,
//               dueDateBucketFilter: option?.key === "all" ? null : Number(option?.key)
//             }))}
//             styles={{ root: { width: 150 } }}
//           />
//         </Stack>
//         <Stack.Item grow>
//           <Text variant="medium" styles={{ root: { marginBottom: 6 } }}>Bar by Project</Text>
//           <div style={{ height: 340, maxHeight: 340, overflowY: "auto", overflowX: "hidden" }}>
//             <ResponsiveContainer width="100%" height={Math.max(60, getChartData(state).length * 44)}>
//               {/* <BarChart layout="vertical" data={getChartData(state)} margin={{ left: 90 }}>
//                 <CartesianGrid strokeDasharray="3 3" />
//                 <XAxis type="number" />
//                 <YAxis type="category" dataKey="name" width={120} />
//                 <Tooltip />
//                 <Bar dataKey="value" fill="#0078d4" name="Overdue" />
//               </BarChart> */}
//               <BarChart layout="vertical" data={getChartData(state)} margin={{ left: 90 }}>
//                 <CartesianGrid strokeDasharray="3 3" />
//                 <XAxis type="number" />
//                 <YAxis type="category" dataKey="name" width={120} />
//                 <Tooltip />
//                 <Bar dataKey="value" fill="#0078d4" name="Overdue">
//                   <LabelList
//                     dataKey="value"
//                     position="right"
//                     formatter={(v: number) =>
//                       state.showOverdueBy === "amount" ? v.toLocaleString() : v
//                     }
//                   />
//                 </Bar>
//               </BarChart>
//             </ResponsiveContainer>
//             {getChartData(state).length === 0 && (
//               <Text variant="medium" styles={{ root: { color: "#605e5c", margin: 16 } }}>
//                 No overdue invoices for this selection
//               </Text>
//             )}
//           </div>
//         </Stack.Item>
//       </Stack>
//       <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
//         <ChartContainer title="PO Utilization by Project">
//           <div style={{ height: 400, overflowY: "auto", overflowX: "hidden" }}>
//             <ResponsiveContainer width="100%" height={state.poStatusStackedData.length * 40}>
//               <BarChart layout="vertical" data={state.poStatusStackedData} margin={{ top: 5, right: 30, left: 220, bottom: 25 }}>
//                 <CartesianGrid strokeDasharray="3 3" />
//                 <XAxis type="number" />
//                 <YAxis
//                   dataKey="project"
//                   type="category"
//                   width={220}
//                   tickFormatter={(name: any) => (name && name.length > 30 ? name.substr(0, 28) + "..." : name)}
//                   interval={0}
//                   fontSize={12}
//                 />
//                 <Tooltip />
//                 <Legend />
//                 <Bar dataKey="NotInvoiced" stackId="a" fill="#0078d4" name="Not Invoiced" />
//                 <Bar dataKey="Paid" stackId="a" fill="#107c10" name="Invoiced - Paid" />
//                 <Bar dataKey="Pending" stackId="a" fill="#f99923ff" name="Invoiced - Requested" />
//                 {/* <Bar dataKey="Others" stackId="a" fill="#d83b01" name="Invoiced - Others" /> */}
//               </BarChart>
//             </ResponsiveContainer>
//           </div>
//         </ChartContainer>
//       </Stack>
//       <Stack
//         styles={{
//           root: {
//             marginTop: 32,
//             background: "#fff",
//             borderRadius: 12,
//             padding: 16,
//             boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
//           },
//         }}
//       >
//         <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 16 } }}>
//           <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
//             Invoice Requests Table
//           </Text>
//           <PrimaryButton text="Export" onClick={exportTableToExcel} />
//         </Stack>
//         <div style={{ overflowX: "auto" }}>
//           <table
//             style={{
//               width: "100%",
//               borderCollapse: "separate",
//               borderSpacing: 0,
//               minWidth: 800,
//               fontSize: 15,
//               background: "#fafbfc",
//             }}
//           >
//             <thead>
//               <tr>
//                 <th style={{ ...thStyle }}>Invoice No:</th>
//                 <th style={{ ...thStyle }}>Project Name</th>
//                 <th style={{ ...thStyle }}>POID</th>
//                 <th style={{ ...thStyle }}>Due Date</th>
//                 <th style={{ ...thStyle }}>Invoice Status</th>
//                 <th style={{ ...thStyle, textAlign: "right" }}>Invoiced Amount</th>
//               </tr>
//             </thead>
//             <tbody>
//               {state.agingData.map((row, idx) => (
//                 <tr
//                   key={idx}
//                   style={{
//                     background: idx % 2 === 0 ? "#fff" : "#f4f6fa",
//                     transition: "background 0.2s",
//                     borderBottom: "1px solid #eee",
//                   }}
//                 >
//                   <td style={tdStyle}>{row.invoiceNumber}</td>
//                   <td style={tdStyle}>{row.projectName}</td>
//                   <td style={tdStyle}>{row.poid}</td>
//                   <td style={tdStyle}>{row.dueDate}</td>
//                   <td
//                     style={{
//                       ...tdStyle,
//                       color: row.status.toLowerCase().includes("pending")
//                         ? "#ffaa44"
//                         : row.status.toLowerCase().includes("received")
//                           ? "#107c10"
//                           : row.status.toLowerCase().includes("not")
//                             ? "#605e5c"
//                             : "#d83b01",
//                       fontWeight: 500,
//                     }}
//                   >
//                     {row.status}
//                   </td>
//                   <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>
//                     {row.currency ? row.currency + " " : ""}
//                     {row.amount.toLocaleString()}
//                   </td>
//                 </tr>
//               ))}
//             </tbody>
//           </table>
//         </div>
//       </Stack>
//     </div>
//   );
// }

// function DashboardCard({ title, value, icon, color }: { title: string; value: number; icon: string; color: string }) {
//   return (
//     <div
//       style={{
//         flex: 1,
//         backgroundColor: "#fff",
//         borderRadius: 12,
//         padding: 16,
//         minWidth: 180,
//         boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
//         display: "flex",
//         alignItems: "center",
//         justifyContent: "space-between", // pushes left and right content apart
//       }}
//     >
//       {/* Left: label */}
//       <Text variant="mediumPlus" styles={{ root: { color: "#333" } }}>
//         {title}
//       </Text>
//       {/* Right: value and icon (grouped) */}
//       <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
//         <Text variant="xxLarge" styles={{ root: { fontWeight: 600, color } }}>
//           {value}
//         </Text>
//         <Icon iconName={icon} styles={{ root: { fontSize: 40, color, opacity: 0.8 } }} />
//       </div>
//     </div>
//   );
// }

// function ChartContainer({ title, children }: { title: string; children: React.ReactNode }) {
//   return (
//     <Stack.Item
//       grow={1}
//       styles={{
//         root: {
//           minWidth: 400,
//           height: 350,
//           background: "#fff",
//           borderRadius: 12,
//           boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
//           padding: 16,
//         },
//       }}
//     >
//       <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>
//         {title}
//       </Text>
//       {children}
//     </Stack.Item>
//   );
// }
import * as React from "react";
import {
  Stack,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  Icon,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
} from "@fluentui/react";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  Tooltip,
  Legend,
  CartesianGrid,
  PieChart,
  Pie,
  Cell,
  LabelList,
} from "recharts";
import { SPFI } from "@pnp/sp";
import { XAxis as _XAxis, YAxis as _YAxis } from "recharts";
const XAxis = _XAxis as any;
const YAxis = _YAxis as any;
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";

interface DashboardProps {
  sp: SPFI;
  context: any;
  projectsp: SPFI;
}

interface AgingRow {
  projectName: string;
  poid: string;
  invoiceNumber: string;
  dueDate: string;
  status: string;
  amount: number;
  currency?: string;
}

interface DashboardState {
  loading: boolean;
  error: string | null;
  totalInvoices: number;
  pendingInvoices: number;
  paidInvoices: number;
  monthlyData: { month: string; amount: number }[];
  statusData: { name: string; value: number }[];
  projectComparisonData: { project: string; poAmount: number; invoiceAmount: number }[];
  currencyData: { name: string; value: number }[];
  agingData: AgingRow[];
  period: string;
  startDate: Date | null;
  endDate: Date | null;
  business: string;
  businessUnit: string;
  customer: string;
  project: string;
  businessOptions: IDropdownOption[];
  businessUnitOptions: IDropdownOption[];
  customerOptions: IDropdownOption[];
  projectOptions: IDropdownOption[];
  statusFilter: string;
  currentStatusFilter: string;
  monthlyProjectComparisonData: { month: string; invoiceAmount: number; poAmount: number; paidAmount: number }[];
  poStatusStackedData: any[];
  currentStatusPieData?: { name: string; value: number }[];
  overdueDataCount: Record<string, number[]>;
  overdueDataAmount: Record<string, number[]>;
  showOverdueBy: "count" | "amount";
  dueDateBucketFilter: number | null;
  currentStatusOptions?: IDropdownOption[];
  statusOptions?: IDropdownOption[];
}

const periodOptions: IDropdownOption[] = [
  { key: "all", text: "All" },
  { key: "week_to_date", text: "Week till date" },
  { key: "last_week", text: "Last Week" },
  { key: "month_to_date", text: "Month till date" },
  { key: "current_month", text: "Current Month" },
  { key: "last_month", text: "Last Month" },
  { key: "year_to_date", text: "Year till date" },
  { key: "current_year", text: "Current Year" },
  { key: "last_year", text: "Last Year" },
];

const pieColors = ["#0078d4", "#ffaa44", "#107c10", "#d83b01", "#605e5c", "#fcbd73", "#a1a7b3"];

const thStyle: React.CSSProperties = {
  padding: "12px 16px",
  background: "#f0f2f7",
  color: "#222",
  textAlign: "left",
  fontWeight: 700,
  borderBottom: "2px solid #dde0eb",
};
const tdStyle: React.CSSProperties = {
  padding: "12px 16px",
  color: "#333",
  verticalAlign: "middle",
  borderBottom: "1px solid #eee",
};

// ===== Currency helper =====
function getCurrencySymbol(currencyCode: string, locale: string = "en-US") {
  if (!currencyCode || !currencyCode.trim()) {
    return "";
  }
  try {
    return (
      new Intl.NumberFormat(locale, {
        style: "currency",
        currency: currencyCode,
        minimumFractionDigits: 0,
        maximumFractionDigits: 0,
      })
        .formatToParts(1)
        .find((part) => part.type === "currency")?.value ?? currencyCode
    );
  } catch (error) {
    console.warn("Invalid currency code", currencyCode, error);
    return currencyCode;
  }
}

function formatAmountWithCurrency(amount: number, currencyCode: string | undefined) {
  if (!currencyCode) {
    return amount.toLocaleString();
  }
  const symbol = getCurrencySymbol(currencyCode);
  return `${symbol} ${amount.toLocaleString()}`;
}

// ===== Utility for period range =====
function getPeriodRange(period: string) {
  const now = new Date();
  switch (period) {
    case "week_to_date": {
      const first = new Date(now);
      first.setDate(now.getDate() - now.getDay());
      return { start: first, end: now };
    }
    case "last_week": {
      const first = new Date(now);
      first.setDate(now.getDate() - now.getDay() - 7);
      const last = new Date(first);
      last.setDate(first.getDate() + 6);
      return { start: first, end: last };
    }
    case "month_to_date":
      return { start: new Date(now.getFullYear(), now.getMonth(), 1), end: now };
    case "current_month":
      return {
        start: new Date(now.getFullYear(), now.getMonth(), 1),
        end: new Date(now.getFullYear(), now.getMonth() + 1, 0),
      };
    case "last_month":
      return {
        start: new Date(now.getFullYear(), now.getMonth() - 1, 1),
        end: new Date(now.getFullYear(), now.getMonth(), 0),
      };
    case "year_to_date":
      return { start: new Date(now.getFullYear(), 0, 1), end: now };
    case "current_year":
      return { start: new Date(now.getFullYear(), 0, 1), end: new Date(now.getFullYear(), 11, 31) };
    case "last_year":
      return {
        start: new Date(now.getFullYear() - 1, 0, 1),
        end: new Date(now.getFullYear() - 1, 11, 31),
      };
    default:
      return { start: null, end: null };
  }
}

function inDateRange(dateStr: string | undefined, start: Date | null, end: Date | null) {
  if (!dateStr || dateStr === "") return true;
  const date = new Date(dateStr);
  if (start && date < start) return false;
  if (end && date > end) return false;
  return true;
}

// ===== Overdue buckets =====
interface OverdueBucket {
  label: string;
  minDays: number;
  maxDays: number | null;
}

const overdueBuckets: OverdueBucket[] = [
  { label: "0-30 days", minDays: 0, maxDays: 30 },
  { label: "31-60 days", minDays: 31, maxDays: 60 },
  { label: "61-90 days", minDays: 61, maxDays: 90 },
  { label: "90+ days", minDays: 91, maxDays: null },
];

function parseDMY(d: string) {
  const [day, month, year] = d.split("/").map(Number);
  return new Date(year, month - 1, day);
}

function getOverdueDays(dueDateStr: string) {
  if (!dueDateStr || dueDateStr === "") return null;
  let dueDate: Date;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dueDateStr)) {
    dueDate = parseDMY(dueDateStr);
  } else {
    dueDate = new Date(dueDateStr);
  }
  if (isNaN(dueDate.getTime())) return null;
  const now = new Date();
  const diff = Math.floor((now.getTime() - dueDate.getTime()) / (1000 * 3600 * 24));
  return diff > 0 ? diff : 0;
}

function aggregateOverdueData(invoices: AgingRow[]) {
  const aggregation: Record<string, { count: number[]; amount: number[] }> = {};
  invoices.forEach((inv) => {
    const overdueDays = getOverdueDays(inv.dueDate);
    if (overdueDays === null || overdueDays === 0) return;

    let bucketIndex = overdueBuckets.findIndex((b) => {
      if (b.maxDays === null) return overdueDays >= b.minDays;
      return overdueDays >= b.minDays && overdueDays <= b.maxDays;
    });
    if (bucketIndex === -1) bucketIndex = overdueBuckets.length - 1;

    const key = inv.projectName;
    if (!aggregation[key]) {
      aggregation[key] = {
        count: Array(overdueBuckets.length).fill(0),
        amount: Array(overdueBuckets.length).fill(0),
      };
    }
    aggregation[key].count[bucketIndex]++;
    aggregation[key].amount[bucketIndex] += inv.amount;
  });

  return aggregation;
}

// ===== Monthly Project PO/Invoice/Paid Amounts =====
function getMonthlyInvoicePaidComparison(
  invoiceItems: any[],
  poItems: any[],
  startDate: Date | null,
  endDate: Date | null
) {
  const raw: Record<string, { month: string; invoiceAmount: number; poAmount: number; paidAmount: number }> = {};

  invoiceItems.forEach((i) => {
    const date = i.Created ? new Date(i.Created) : null;
    if (!date) return;
    if (startDate && date < startDate) return;
    if (endDate && date > endDate) return;

    const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
    if (!raw[key]) raw[key] = { month: key, invoiceAmount: 0, poAmount: 0, paidAmount: 0 };

    raw[key].invoiceAmount += i.InvoiceAmount || 0;
    if (i.Status === "Payment Received") {
      raw[key].paidAmount += i.InvoiceAmount || 0;
    }
  });

  poItems.forEach((po) => {
    const date = po.Created ? new Date(po.Created) : null;
    if (!date) return;
    if (startDate && date < startDate) return;
    if (endDate && date > endDate) return;

    const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
    if (!raw[key]) raw[key] = { month: key, invoiceAmount: 0, poAmount: 0, paidAmount: 0 };

    raw[key].poAmount += po.POAmount || 0;
  });

  return Object.values(raw).sort((a, b) => {
    const [am, ay] = a.month.split(" ");
    const [bm, by] = b.month.split(" ");
    return new Date(`${am} 1, ${ay}`).getTime() - new Date(`${bm} 1, ${by}`).getTime();
  });
}

function getPOStatusStackedData(
  filteredInvoiceItems: any[],
  poItems: any[],
  projectToBusiness: any,
  projectToBusinessUnit: any,
  selectedBusiness: string,
  selectedBusinessUnit: string
) {
  return poItems
    .filter((po) => {
      const project = po.ProjectName;
      const biz = projectToBusiness[project];
      const bu = projectToBusinessUnit[project];
      return (!selectedBusiness || biz === selectedBusiness) && (!selectedBusinessUnit || bu === selectedBusinessUnit);
    })
    .map((po) => {
      const project = po.ProjectName;
      const invoicesForProject = filteredInvoiceItems.filter((i) => i.ProjectName === project);
      const invoiced = invoicesForProject.reduce((sum, i) => sum + (i.InvoiceAmount || 0), 0);
      const invoicePaid = invoicesForProject
        .filter((i) => i.Status === "Payment Received")
        .reduce((sum, i) => sum + (i.InvoiceAmount || 0), 0);
      const invoicePending = invoiced - invoicePaid;
      const notInvoiced = (po.POAmount || 0) - invoiced;
      return {
        project,
        POAmount: po.POAmount || 0,
        Invoiced: invoiced,
        Paid: invoicePaid,
        Pending: invoicePending,
        NotInvoiced: notInvoiced,
      };
    });
}

// ===== Overdue chart data helper =====
function getChartData(state: DashboardState) {
  const source = state.showOverdueBy === "count" ? state.overdueDataCount : state.overdueDataAmount;
  return Object.entries(source)
    .map(([project, values]) => ({
      name: project,
      value:
        state.dueDateBucketFilter === null
          ? values.reduce((a, b) => a + b, 0)
          : values[state.dueDateBucketFilter!],
    }))
    .filter((d) => d.value > 0);
}

export default function Dashboard({ sp, projectsp, context }: DashboardProps) {
  const [state, setState] = React.useState<DashboardState>({
    loading: true,
    error: null,
    totalInvoices: 0,
    pendingInvoices: 0,
    paidInvoices: 0,
    monthlyData: [],
    statusData: [],
    projectComparisonData: [],
    currencyData: [],
    agingData: [],
    statusFilter: "",
    currentStatusFilter: "",
    period: "all",
    startDate: null,
    endDate: null,
    business: "",
    businessUnit: "",
    customer: "",
    project: "",
    businessOptions: [],
    businessUnitOptions: [],
    customerOptions: [],
    projectOptions: [],
    monthlyProjectComparisonData: [],
    poStatusStackedData: [],
    currentStatusPieData: [],
    overdueDataCount: {},
    overdueDataAmount: {},
    showOverdueBy: "count",
    dueDateBucketFilter: null,
  });
  const [rawInvoiceItems, setRawInvoiceItems] = React.useState<any[]>([]);
  const [rawPoItems, setRawPoItems] = React.useState<any[]>([]);
  const [projectToBusiness, setProjectToBusiness] = React.useState<{ [key: string]: string }>({});
  const [projectToBusinessUnit, setProjectToBusinessUnit] = React.useState<{ [key: string]: string }>({});
  const [projectCurrency, setProjectCurrency] = React.useState<Record<string, string>>({});

  React.useEffect(() => {
    loadDashboardData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function loadDashboardData() {
    try {
      setState((prev) => ({ ...prev, loading: true, error: null }));
      const projectBiz: { [key: string]: string } = {};
      const projectBU: { [key: string]: string } = {};
      const invoiceStatusSet = new Set<string>();
      const currentStatusSet = new Set<string>();
      setProjectToBusiness(projectBiz);
      setProjectToBusinessUnit(projectBU);

      const invoiceItems = await sp.web.lists
        .getByTitle("Invoice Requests")
        .items.select(
          "Id",
          "Status",
          "InvoiceAmount",
          "Created",
          "InvoiceNumber",
          "DueDate",
          "ProjectName",
          "Currency",
          "PurchaseOrder",
          "POItem_x0020_Value",
          "POItem_x0020_Title"
        )();

      const poItems = await sp.web.lists
        .getByTitle("InvoicePO")
        .items.select("Id", "POAmount", "ProjectName", "Currency", "Customer", "Created")();

      setRawInvoiceItems(invoiceItems);
      setRawPoItems(poItems);

      const projectCurrencyMap: Record<string, string> = {};
      poItems.forEach((po) => {
        if (po.ProjectName && po.Currency) {
          projectCurrencyMap[po.ProjectName] = po.Currency;
        }
      });
      setProjectCurrency(projectCurrencyMap);

      const projectOpts = Array.from(
        new Set([...invoiceItems, ...poItems].map((i) => i.ProjectName).filter(Boolean))
      ).map((b) => ({ key: b, text: b }));

      const customerOpts = Array.from(new Set(poItems.map((p) => p.Customer).filter(Boolean))).map((b) => ({
        key: b,
        text: b,
      }));

      invoiceItems.forEach((item) => {
        if (item.Status) invoiceStatusSet.add(item.Status);
        if (item.CurrentStatus) currentStatusSet.add(item.CurrentStatus);
      });

      const sortedStatusOptions = Array.from(invoiceStatusSet)
        .sort()
        .map((s) => ({ key: s, text: s }));
      const sortedCurrentStatusOptions = Array.from(currentStatusSet)
        .sort()
        .map((s) => ({ key: s, text: s }));

      sortedStatusOptions.unshift({ key: "", text: "All Statuses" });
      sortedCurrentStatusOptions.unshift({ key: "", text: "All Current Statuses" });

      setState((prev) => ({
        ...prev,
        projectOptions: [{ key: "", text: "All Projects" }, ...projectOpts],
        customerOptions: [{ key: "", text: "All Customers" }, ...customerOpts],
        statusOptions: sortedStatusOptions,
        currentStatusOptions: sortedCurrentStatusOptions,
      }));

      processDashboardData(
        invoiceItems,
        poItems,
        projectBiz,
        projectBU,
        "",
        "",
        "all",
        null,
        null,
        "",
        "",
        "",
        "",
        false,
        projectCurrencyMap
      );
    } catch (err: any) {
      setState((prev) => ({ ...prev, loading: false, error: err.message || "Error loading data" }));
    }
  }

  function exportTableToExcel() {
    const headers = ["Invoice No", "Project Name", "POID", "Due Date", "Invoice Status", "Invoiced Amount"];
    const data = state.agingData.map((row) => [
      row.invoiceNumber,
      row.projectName,
      row.poid,
      row.dueDate,
      row.status,
      formatAmountWithCurrency(row.amount, row.currency),
    ]);
    const worksheetData = [headers, ...data];

    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invoice Requests");

    const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    saveAs(blob, "invoice_requests.xlsx");
  }

  function processDashboardData(
    invoiceItems: any[],
    poItems: any[],
    projectToBusinessMap: { [k: string]: string },
    projectToBusinessUnitMap: { [k: string]: string },
    statusFilter: string,
    currentStatusFilter: string,
    period: string,
    startDate: Date | null,
    endDate: Date | null,
    business: string,
    businessUnit: string,
    customer: string,
    project: string,
    updateOnly?: boolean,
    projectCurrencyMapParam?: Record<string, string>
  ) {
    const currencyMapToUse = projectCurrencyMapParam || projectCurrency;

    let actualStart = startDate,
      actualEnd = endDate;
    if (period !== "all") {
      const { start, end } = getPeriodRange(period);
      actualStart = start;
      actualEnd = end;
    }

    let filteredInvoiceItems = invoiceItems.filter((inv) => {
      const projectName = inv.ProjectName;
      const projectBusiness = projectToBusinessMap[projectName] || "";
      const projectBusinessUnit = projectToBusinessUnitMap[projectName] || "";
      const poRow = poItems.find((po) => po.ProjectName === projectName);
      const poCustomer = poRow?.Customer || "";
      let statusIncluded = true;

      if (statusFilter === "Pending")
        statusIncluded =
          inv.Status &&
          (inv.Status.toLowerCase().includes("pending") ||
            inv.Status.toLowerCase().includes("invoice requested") ||
            inv.Status.toLowerCase().includes("invoice raised"));
      else if (statusFilter === "Paid")
        statusIncluded = inv.Status && inv.Status.toLowerCase().includes("received");
      else if (statusFilter === "Others")
        statusIncluded =
          inv.Status &&
          !(
            inv.Status.toLowerCase().includes("pending") ||
            inv.Status.toLowerCase().includes("received") ||
            inv.Status.toLowerCase().includes("invoice requested") ||
            inv.Status.toLowerCase().includes("invoice raised")
          );

      let currentStatusIncluded = currentStatusFilter ? inv.Status === currentStatusFilter : true;

      let dateIncluded = true;
      if (actualStart || actualEnd) {
        const baseDate = inv.DueDate ? new Date(inv.DueDate) : inv.Created ? new Date(inv.Created) : null;
        dateIncluded = baseDate ? inDateRange(baseDate.toISOString(), actualStart, actualEnd) : true;
      }

      let bizIncluded = business ? projectBusiness === business : true;
      let unitIncluded = businessUnit ? projectBusinessUnit === businessUnit : true;
      let custIncluded = customer ? poCustomer === customer : true;
      let projectIncluded = project ? projectName === project : true;

      return statusIncluded && currentStatusIncluded && dateIncluded && bizIncluded && unitIncluded && custIncluded && projectIncluded;
    });

    let filteredPOItems = poItems.filter((po) => {
      const projectName = po.ProjectName;
      const projectBusiness = projectToBusinessMap[projectName] || "";
      const projectBusinessUnit = projectToBusinessUnitMap[projectName] || "";
      let bizIncluded = business ? projectBusiness === business : true;
      let unitIncluded = businessUnit ? projectBusinessUnit === businessUnit : true;
      let custIncluded = customer ? po.Customer === customer : true;
      let projectIncluded = project ? projectName === project : true;
      return bizIncluded && unitIncluded && custIncluded && projectIncluded;
    });

    const totalInvoices = filteredInvoiceItems.length;
    const paidInvoices = filteredInvoiceItems.filter((i) => i.Status?.toLowerCase().includes("received")).length;
    const pendingInvoices = filteredInvoiceItems.filter(
      (i) =>
        i.Status?.toLowerCase().includes("pending") ||
        i.Status?.toLowerCase().includes("invoice requested") ||
        i.Status?.toLowerCase().includes("invoice raised")
    ).length;

    const monthlyMap: Record<string, number> = {};
    filteredInvoiceItems.forEach((i) => {
      if (!i.Created) return;
      const date = new Date(i.Created);
      const key = `${date.toLocaleString("default", { month: "short" })} ${date.getFullYear()}`;
      monthlyMap[key] = (monthlyMap[key] || 0) + (i.InvoiceAmount || 0);
    });
    const monthlyData = Object.keys(monthlyMap)
      .map((m) => ({ month: m, amount: monthlyMap[m] }))
      .sort((a, b) => new Date(`1 ${a.month}`).getTime() - new Date(`1 ${b.month}`).getTime());

    const projectMapAmounts: Record<string, { poAmount: number; invoiceAmount: number }> = {};
    filteredPOItems.forEach((po) => {
      const projectName = po.ProjectName;
      if (!projectMapAmounts[projectName]) projectMapAmounts[projectName] = { poAmount: 0, invoiceAmount: 0 };
      projectMapAmounts[projectName].poAmount += po.POAmount || 0;
    });
    filteredInvoiceItems.forEach((inv) => {
      const projectName = inv.ProjectName;
      if (!projectMapAmounts[projectName]) projectMapAmounts[projectName] = { poAmount: 0, invoiceAmount: 0 };
      projectMapAmounts[projectName].invoiceAmount += inv.InvoiceAmount || 0;
    });
    const projectComparisonData = Object.keys(projectMapAmounts).map((projectName) => ({
      project: projectName,
      poAmount: projectMapAmounts[projectName].poAmount,
      invoiceAmount: projectMapAmounts[projectName].invoiceAmount,
    }));

    const currencyMap: Record<string, number> = {};
    filteredInvoiceItems.forEach((inv) => {
      const projectName = inv.ProjectName;
      const currencyCode = currencyMapToUse[projectName] || inv.Currency || "";
      if (!currencyCode) return;
      currencyMap[currencyCode] = (currencyMap[currencyCode] || 0) + (inv.InvoiceAmount || 0);
    });
    const currencyData = Object.keys(currencyMap).map((c) => ({ name: c, value: currencyMap[c] }));

    const agingData: AgingRow[] = filteredInvoiceItems.map((inv) => {
      const projectName = inv.ProjectName || "";
      const projectCurrencyCode = currencyMapToUse[projectName] || inv.Currency || "";
      return {
        projectName,
        poid: inv.PurchaseOrder || "",
        invoiceNumber: inv.InvoiceNumber || "",
        dueDate: inv.DueDate ? new Date(inv.DueDate).toLocaleDateString() : "",
        status: inv.Status || "",
        amount: inv.InvoiceAmount || 0,
        currency: projectCurrencyCode,
      };
    });

    const overdueAggregation = aggregateOverdueData(agingData);
    const overdueDataCount: Record<string, number[]> = {};
    const overdueDataAmount: Record<string, number[]> = {};
    Object.entries(overdueAggregation).forEach(([key, val]) => {
      overdueDataCount[key] = val.count;
      overdueDataAmount[key] = val.amount;
    });

    const currentStatusCount: Record<string, number> = {};
    filteredInvoiceItems.forEach((inv) => {
      const stat = inv.Status || "Unknown";
      currentStatusCount[stat] = (currentStatusCount[stat] || 0) + 1;
    });
    const currentStatusPieData = Object.keys(currentStatusCount).map((s) => ({
      name: s,
      value: currentStatusCount[s],
    }));

    const monthlyProjectComparisonData = getMonthlyInvoicePaidComparison(
      filteredInvoiceItems,
      filteredPOItems,
      actualStart,
      actualEnd
    );
    const poStatusStackedData = getPOStatusStackedData(
      filteredInvoiceItems,
      filteredPOItems,
      projectToBusinessMap,
      projectToBusinessUnitMap,
      business,
      businessUnit
    );

    setState((prev) => ({
      ...prev,
      loading: false,
      error: null,
      totalInvoices,
      pendingInvoices,
      paidInvoices,
      monthlyData,
      statusData: [
        { name: "Pending", value: pendingInvoices },
        { name: "Paid", value: paidInvoices },
        { name: "Others", value: totalInvoices - (pendingInvoices + paidInvoices) },
      ],
      projectComparisonData,
      currencyData,
      agingData,
      statusFilter,
      currentStatusFilter,
      period,
      startDate: actualStart,
      endDate: actualEnd,
      business,
      businessUnit,
      customer,
      project,
      monthlyProjectComparisonData,
      poStatusStackedData,
      currentStatusPieData,
      overdueDataCount,
      overdueDataAmount,
    }));
  }

  React.useEffect(() => {
    if (rawInvoiceItems.length > 0 && rawPoItems.length > 0) {
      processDashboardData(
        rawInvoiceItems,
        rawPoItems,
        projectToBusiness,
        projectToBusinessUnit,
        state.statusFilter,
        state.currentStatusFilter,
        state.period,
        state.startDate,
        state.endDate,
        state.business,
        state.businessUnit,
        state.customer,
        state.project,
        true
      );
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    state.statusFilter,
    state.currentStatusFilter,
    state.period,
    state.startDate,
    state.endDate,
    state.business,
    state.businessUnit,
    state.customer,
    state.project,
    projectToBusiness,
    projectToBusinessUnit,
    rawInvoiceItems,
    rawPoItems,
  ]);

  if (state.loading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ height: "80vh" }}>
        <Spinner label="Loading Dashboard..." />
      </Stack>
    );
  }

  if (state.error) {
    return <MessageBar messageBarType={MessageBarType.error}>{state.error}</MessageBar>;
  }

  return (
    <div style={{ padding: 20 }}>
      <Text variant="xxLarge" styles={{ root: { fontWeight: 600, marginBottom: 20 } }}>
        Invoice Dashboard
      </Text>

      <Stack horizontal tokens={{ childrenGap: 18 }} style={{ marginBottom: 18, flexWrap: "wrap" }}>
        <Dropdown
          label="Period"
          options={periodOptions}
          selectedKey={state.period}
          onChange={(_, o) => {
            const { start, end } = getPeriodRange(o?.key as string);
            setState((prev) => ({ ...prev, period: o?.key as string, startDate: start, endDate: end }));
          }}
          style={{ minWidth: 150 }}
        />
        <Dropdown
          label="Project"
          options={state.projectOptions}
          selectedKey={state.project}
          onChange={(_, o) => setState((prev) => ({ ...prev, project: o?.key as string }))}
          style={{ minWidth: 180 }}
        />
        <Dropdown
          label="Customer"
          options={state.customerOptions}
          selectedKey={state.customer}
          onChange={(_, o) => setState((prev) => ({ ...prev, customer: o?.key as string }))}
          style={{ minWidth: 160 }}
        />
        <Dropdown
          label="Current Status"
          options={state.currentStatusOptions}
          selectedKey={state.currentStatusFilter}
          onChange={(_, o) => setState((prev) => ({ ...prev, currentStatusFilter: o?.key as string }))}
          style={{ minWidth: 170 }}
        />
        <Dropdown
          label="Invoice Status"
          options={state.statusOptions}
          selectedKey={state.statusFilter}
          onChange={(_, o) => setState((prev) => ({ ...prev, statusFilter: o?.key as string }))}
          style={{ minWidth: 150 }}
        />
      </Stack>

      <Stack horizontal horizontalAlign="space-between" tokens={{ childrenGap: 20 }} wrap>
        <DashboardCard title="Total Invoices" value={state.totalInvoices} icon="ReceiptCheck" color="#0078d4" />
        <DashboardCard title="Pending" value={state.pendingInvoices} icon="Clock" color="#ffaa44" />
        <DashboardCard title="Paid" value={state.paidInvoices} icon="CheckMark" color="#107c10" />
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 24 }} wrap styles={{ root: { marginTop: 20, marginBottom: 32 } }}>
        <ChartContainer title="Invoices by Status">
          <ResponsiveContainer width="100%" height={320}>
            <PieChart>
              <Pie
                dataKey="value"
                data={state.currentStatusPieData}
                cx="50%"
                cy="50%"
                outerRadius={100}
                label={({ name, value }) => `${name}: ${value}`}
              >
                {(state.currentStatusPieData || []).map((entry, idx) => (
                  <Cell key={entry.name} fill={pieColors[idx % pieColors.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </ChartContainer>

        <ChartContainer title="Monthly PO vs Invoice by Project">
          <ResponsiveContainer width="100%" height="90%">
            <BarChart
              data={state.monthlyProjectComparisonData}
              margin={{ top: 10, right: 20, left: 0, bottom: 40 }}
              barCategoryGap="25%"
            >
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="month" angle={-25} textAnchor="end" interval={0} />
              <YAxis />
              {/* <Tooltip
                formatter={(v: any) => {
                  const code = state.project ? projectCurrency[state.project] : "";
                  return formatAmountWithCurrency(Number(v), code || "");
                }}
              /> */}
              <Tooltip
                formatter={(v: any, _name: any, entry: any) => {
                  const monthLabel = entry?.payload?.month as string;
                  // Try to derive a currency for this month from filtered PO items
                  const poForMonth = rawPoItems.find((po) => {
                    if (!po.Created) return false;
                    const d = new Date(po.Created);
                    const label = `${d.toLocaleString("default", { month: "short" })} ${d.getFullYear()}`;
                    return label === monthLabel;
                  });
                  const code =
                    (state.project && projectCurrency[state.project]) ||
                    poForMonth?.Currency ||
                    "";

                  return formatAmountWithCurrency(Number(v), code || undefined);
                }}
              />
              <Legend />
              <Bar dataKey="poAmount" name="PO Amount" fill="#0078d4">
                {/* <LabelList
                  dataKey="poAmount"
                  position="top"
                  formatter={(v: number) => formatAmountWithCurrency(v, "")}
                /> */}
                <LabelList
                  dataKey="poAmount"
                  position="top"
                  formatter={(v: number, _name: any, entry: any) => {
                    const monthLabel = entry?.month as string;
                    const poForMonth = rawPoItems.find((po) => {
                      if (!po.Created) return false;
                      const d = new Date(po.Created);
                      const label = `${d.toLocaleString("default", { month: "short" })} ${d.getFullYear()}`;
                      return label === monthLabel;
                    });
                    const code =
                      (state.project && projectCurrency[state.project]) ||
                      poForMonth?.Currency ||
                      "";
                    return formatAmountWithCurrency(v, code || undefined);
                  }}
                />
              </Bar>
              <Bar dataKey="invoiceAmount" name="Invoiced Amount" fill="#107c10">
                {/* <LabelList
                  dataKey="invoiceAmount"
                  position="top"
                  formatter={(v: number) => formatAmountWithCurrency(v, "")}
                /> */}
                <LabelList
                  dataKey="invoiceAmount"
                  position="top"
                  formatter={(v: number, _name: any, entry: any) => {
                    const monthLabel = entry?.month as string;
                    const invForMonth = rawInvoiceItems.find((inv) => {
                      if (!inv.Created) return false;
                      const d = new Date(inv.Created);
                      const label = `${d.toLocaleString("default", { month: "short" })} ${d.getFullYear()}`;
                      return label === monthLabel;
                    });
                    const code =
                      (state.project && projectCurrency[state.project]) ||
                      invForMonth?.Currency ||
                      "";
                    return formatAmountWithCurrency(v, code || undefined);
                  }}
                />
              </Bar>
              <Bar dataKey="paidAmount" name="Paid Amount" fill="#ffaa44">
                {/* <LabelList
                  dataKey="paidAmount"
                  position="top"
                  formatter={(v: number) => formatAmountWithCurrency(v, "")}
                /> */}
                <LabelList
                  dataKey="paidAmount"
                  position="top"
                  formatter={(v: number, _name: any, entry: any) => {
                    const monthLabel = entry?.month as string;
                    const invForMonth = rawInvoiceItems.find((inv) => {
                      if (!inv.Created) return false;
                      const d = new Date(inv.Created);
                      const label = `${d.toLocaleString("default", { month: "short" })} ${d.getFullYear()}`;
                      return label === monthLabel;
                    });
                    const code =
                      (state.project && projectCurrency[state.project]) ||
                      invForMonth?.Currency ||
                      "";
                    return formatAmountWithCurrency(v, code || undefined);
                  }}
                />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </ChartContainer>
      </Stack>

      <Stack
        styles={{
          root: { background: "#fff", borderRadius: 12, padding: 16, marginTop: 32, marginBottom: 24, boxShadow: "0 2px 8px rgba(0,0,0,0.1)" },
        }}
      >
        <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 12 } }}>
          Overdue Invoices Breakdown ({state.showOverdueBy === "count" ? "Count" : "Amount"})
        </Text>
        <Stack horizontal tokens={{ childrenGap: 14 }} styles={{ root: { marginBottom: 12 } }}>
          <Dropdown
            label="Show by"
            selectedKey={state.showOverdueBy}
            options={[
              { key: "count", text: "Count" },
              { key: "amount", text: "Amount" },
            ]}
            onChange={(_, option) =>
              setState((prev) => ({ ...prev, showOverdueBy: option?.key as "count" | "amount" }))
            }
            styles={{ root: { width: 140 } }}
          />
          <Dropdown
            label="Due Date Period"
            selectedKey={state.dueDateBucketFilter !== null ? state.dueDateBucketFilter : "all"}
            options={[
              { key: "all", text: "All Periods" },
              ...overdueBuckets.map((b, idx) => ({ key: idx, text: b.label })),
            ]}
            onChange={(_, option) =>
              setState((prev) => ({
                ...prev,
                dueDateBucketFilter: option?.key === "all" ? null : Number(option?.key),
              }))
            }
            styles={{ root: { width: 150 } }}
          />
        </Stack>
        <Stack.Item grow>
          <Text variant="medium" styles={{ root: { marginBottom: 6 } }}>
            Bar by Project
          </Text>
          <div style={{ height: 340, maxHeight: 340, overflowY: "auto", overflowX: "hidden" }}>
            <ResponsiveContainer width="100%" height={Math.max(60, getChartData(state).length * 44)}>
              <BarChart layout="vertical" data={getChartData(state)} margin={{ left: 90 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" />
                <YAxis type="category" dataKey="name" width={120} />
                <Tooltip
                  formatter={(v: any) =>
                    state.showOverdueBy === "amount"
                      ? formatAmountWithCurrency(Number(v), "")
                      : v
                  }
                />
                <Bar dataKey="value" fill="#0078d4" name="Overdue">
                  <LabelList
                    dataKey="value"
                    position="right"
                    formatter={(v: number) =>
                      state.showOverdueBy === "amount" ? formatAmountWithCurrency(v, "") : v
                    }
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
            {getChartData(state).length === 0 && (
              <Text variant="medium" styles={{ root: { color: "#605e5c", margin: 16 } }}>
                No overdue invoices for this selection
              </Text>
            )}
          </div>
        </Stack.Item>
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
        <ChartContainer title="PO Utilization by Project">
          <div style={{ height: 400, overflowY: "auto", overflowX: "hidden" }}>
            <ResponsiveContainer width="100%" height={state.poStatusStackedData.length * 40}>
              <BarChart
                layout="vertical"
                data={state.poStatusStackedData}
                margin={{ top: 5, right: 30, left: 220, bottom: 25 }}
              >
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" />
                <YAxis
                  dataKey="project"
                  type="category"
                  width={220}
                  tickFormatter={(name: any) =>
                    name && name.length > 30 ? name.substr(0, 28) + "..." : name
                  }
                  interval={0}
                  fontSize={12}
                />
                <Tooltip
                  formatter={(v: any, _name: any, entry: any) => {
                    const proj = entry?.payload?.project as string;
                    const code = projectCurrency[proj];
                    return formatAmountWithCurrency(Number(v), code);
                  }}
                />
                <Legend />
                <Bar dataKey="NotInvoiced" stackId="a" fill="#0078d4" name="Not Invoiced" />
                <Bar dataKey="Paid" stackId="a" fill="#107c10" name="Invoiced - Paid" />
                <Bar dataKey="Pending" stackId="a" fill="#f99923ff" name="Invoiced - Requested" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </ChartContainer>
      </Stack>

      <Stack
        styles={{
          root: {
            marginTop: 32,
            background: "#fff",
            borderRadius: 12,
            padding: 16,
            boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
          },
        }}
      >
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{ root: { marginBottom: 16 } }}
        >
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            Invoice Requests Table
          </Text>
          <PrimaryButton text="Export" onClick={exportTableToExcel} />
        </Stack>
        <div style={{ overflowX: "auto" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "separate",
              borderSpacing: 0,
              minWidth: 800,
              fontSize: 15,
              background: "#fafbfc",
            }}
          >
            <thead>
              <tr>
                <th style={{ ...thStyle }}>Invoice No:</th>
                <th style={{ ...thStyle }}>Project Name</th>
                <th style={{ ...thStyle }}>POID</th>
                <th style={{ ...thStyle }}>Due Date</th>
                <th style={{ ...thStyle }}>Invoice Status</th>
                <th style={{ ...thStyle, textAlign: "right" }}>Invoiced Amount</th>
              </tr>
            </thead>
            <tbody>
              {state.agingData.map((row, idx) => (
                <tr
                  key={idx}
                  style={{
                    background: idx % 2 === 0 ? "#fff" : "#f4f6fa",
                    transition: "background 0.2s",
                    borderBottom: "1px solid #eee",
                  }}
                >
                  <td style={tdStyle}>{row.invoiceNumber}</td>
                  <td style={tdStyle}>{row.projectName}</td>
                  <td style={tdStyle}>{row.poid}</td>
                  <td style={tdStyle}>{row.dueDate}</td>
                  <td
                    style={{
                      ...tdStyle,
                      color: row.status.toLowerCase().includes("pending")
                        ? "#ffaa44"
                        : row.status.toLowerCase().includes("received")
                          ? "#107c10"
                          : row.status.toLowerCase().includes("not")
                            ? "#605e5c"
                            : "#d83b01",
                      fontWeight: 500,
                    }}
                  >
                    {row.status}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>
                    {formatAmountWithCurrency(row.amount, row.currency)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Stack>
    </div>
  );
}

function DashboardCard({
  title,
  value,
  icon,
  color,
}: {
  title: string;
  value: number;
  icon: string;
  color: string;
}) {
  return (
    <div
      style={{
        flex: 1,
        backgroundColor: "#fff",
        borderRadius: 12,
        padding: 16,
        minWidth: 180,
        boxShadow: "0 2px 6px rgba(0,0,0,0.1)",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
      }}
    >
      <Text variant="mediumPlus" styles={{ root: { color: "#333" } }}>
        {title}
      </Text>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <Text variant="xxLarge" styles={{ root: { fontWeight: 600, color } }}>
          {value}
        </Text>
        <Icon iconName={icon} styles={{ root: { fontSize: 40, color, opacity: 0.8 } }} />
      </div>
    </div>
  );
}

function ChartContainer({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <Stack.Item
      grow={1}
      styles={{
        root: {
          minWidth: 400,
          height: 350,
          background: "#fff",
          borderRadius: 12,
          boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
          padding: 16,
        },
      }}
    >
      <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>
        {title}
      </Text>
      {children}
    </Stack.Item>
  );
}
