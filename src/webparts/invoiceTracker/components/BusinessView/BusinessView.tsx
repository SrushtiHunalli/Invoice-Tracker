/* eslint-disable @typescript-eslint/no-unused-vars */

import * as React from "react";
import { useEffect, useState } from "react";
import {
  Stack, Spinner, Text,
  DetailsList, IColumn,
  MessageBar, MessageBarType,
  Separator, Dropdown, IDropdownOption,
  Panel, PanelType, SearchBox, Label,
  PrimaryButton, Selection,
} from "office-ui-fabric-react";
import { SPFI } from "@pnp/sp";
// import { set } from "@microsoft/sp-lodash-subset";
interface BusinessViewProps {
  sp: SPFI;
  context: any;
  onNavigate?: (view: string) => void;
  projectsp: SPFI;
}
import DocumentViewer from "../DocumentViewer";
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
      return `[${date} ${time}]${user}${role} - ${title}: ${data}`;
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
  const [, setTotals] = useState<any>({});
  const [poPanel, setPoPanel] = useState<{
    open: boolean;
    po: InvoicePO | null;
    poItems: POItem[];
    invoiceRequests: InvoiceRequest[];
  }>({ open: false, po: null, poItems: [], invoiceRequests: [] });
  const [projects, setProjects] = useState<any[]>([]);
  const [, setDepartments] = useState<IDropdownOption[]>([]);
  const [searchText, setSearchText] = useState<string>("");
  const [invoiceRequestPanel, setInvoiceRequestPanel] = useState<{
    open: boolean;
    invoiceRequest: InvoiceRequest | null;
  }>({ open: false, invoiceRequest: null });
  const [userGroups, setUserGroups] = useState<string[]>([]);
  const [currentUser, setCurrentUser] = useState<string>("");
  const [projectToTeamMap, setProjectToTeamMap] = useState<{ [projectTitle: string]: string }>({});
  const [allowedTeams, setAllowedTeams] = useState<string[] | null>(null); // null means loading
  const [, setTeams] = useState<IDropdownOption[]>([]);
  const [selectedTeam, setSelectedTeam] = useState<string>("__all__");
  const [selectedDepartment, setSelectedDepartment] = useState<string>("__all__");
  const [selectedBusinessUnit, setSelectedBusinessUnit] = useState<string>("__all__");
  const [selectedBusiness, setSelectedBusiness] = useState<string>("__all__");
  const [selectedStatusFilter, setSelectedStatusFilter] = useState<string | null>(null);
  const [selectedCurrentStatusFilter, setSelectedCurrentStatusFilter] = useState<string | null>(null);
  const [attachmentViewer, setAttachmentViewer] = useState<{ isOpen: boolean; url: string; fileName: string }>({
    isOpen: false,
    url: '',
    fileName: '',
  });
  const [selectedPOItemTitle, setSelectedPOItemTitle] = useState<string | null>(null);
  const [businessOptions, setBusinessOptions] = useState<IDropdownOption[]>([]);
  const [businessUnitOptions, setBusinessUnitOptions] = useState<IDropdownOption[]>([]);
  const [departmentOptions, setDepartmentOptions] = useState<IDropdownOption[]>([]);
  const [teamOptions, setTeamOptions] = useState<IDropdownOption[]>([]);
  const [, setProjectToDepartmentMap] = useState<{ [project: string]: string }>({});
  const [projectToBusinessMap, setProjectToBusinessMap] = useState<{ [project: string]: string }>({});
  const [projectToBusinessUnitMap, setProjectToBusinessUnitMap] = useState<{ [project: string]: string }>({});


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
    // {
    //   key: "poitemtitle",
    //   name: "PO Item Title",
    //   fieldName: "POItem_x0020_Title",
    //   minWidth: 140,
    // },
    // {
    //   key: "poitemvalue",
    //   name: "PO Item Value",
    //   fieldName: "POItem_x0020_Value",
    //   minWidth: 120,
    //   onRender: (i) =>
    //     getCurrencySymbol(i.Currency) + (i.POItem_x0020_Value?.toLocaleString() ?? ""),
    // },
    {
      key: "POID",
      name: "Purchase Order",
      fieldName: "PurchaseOrder",
      minWidth: 110,
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

  useEffect(() => {
    async function loadUserGroups() {
      try {
        const groups = await sp.web.currentUser.groups();
        const groupNames = groups.map(g => g.Title.toLowerCase());
        setUserGroups(groupNames);
      } catch (e) {
        console.error("Failed to load user groups", e);
      }
    }
    loadUserGroups();
  }, [sp]);

  useEffect(() => {
    async function loadCurrentUser() {
      const user = await sp.web.currentUser();
      setCurrentUser(user.Title || user.LoginName || "");
    }
    loadCurrentUser();
  }, [sp]);

  useEffect(() => {
    async function loadProjectsAndTeams() {
      try {
        // Fetch projects with PM expanded to get Id and Title
        const projectsWithPM = await projectsp.web.lists.getByTitle("Projects")
          .items
          .select("Id", "Title", "Department", "PM/Id", "PM/Title", "PM/EMail")
          .expand("PM")
          .top(4999)();

        // Query Employees list, expand the Team lookup field, and select MailID and Team/Title
        const employees = await projectsp.web.lists.getByTitle("Employees")
          .items
          .select("MailID", "Team/Title")
          .expand("Team")
          .top(4999)();

        // Map PM MailID to Team title (lookup value)
        const pmMailIdToTeam = employees.reduce((acc: { [mailId: string]: string }, emp) => {
          if (emp.MailID && emp.Team && emp.Team.Title) {
            acc[emp.MailID.toLowerCase()] = emp.Team.Title;
          }
          return acc;
        }, {});

        // Build project to team map using direct PM email → Team title mapping
        const projectTeamMap = projectsWithPM.reduce((acc: { [project: string]: string }, project) => {
          const pmEmail = project.PM?.EMail?.toLowerCase();
          const team = pmEmail ? pmMailIdToTeam[pmEmail] : "";
          console.log(`Project: ${project.Title}, PM Email: ${pmEmail}, Team: ${team || '[none]'}`);
          if (project.Title) {
            acc[project.Title] = team;
          }
          return acc;
        }, {});

        setProjects(projectsWithPM);
        setProjectToTeamMap(projectTeamMap);

        // Identify allowed teams for current user by group role
        const normalizedUser = currentUser.toLowerCase();

        const allTCC = await projectsp.web.lists.getByTitle("Team Cost Center")
          .items
          .select("Title", "Manager/Id", "Manager/Title", "Manager/EMail", "Business", "Department", "BusinessUnit")
          .expand("Manager")
          .top(4999)();

        let allowedTeamsForUser: string[] = [];

        const addUnique = (arr: string[]) => {
          allowedTeamsForUser = Array.from(new Set([...allTCC, ...arr]));
        };

        if (userGroups.includes("business manager")) {
          const managedBusinesses = allTCC
            .filter(row => row.Manager && row.Manager.Title.toLowerCase() === normalizedUser)
            .map(row => row.Title)
            .filter(Boolean);

          const teams = allTCC
            .filter(row => managedBusinesses.includes(row.Business))
            .map(row => row.Title)
            .filter(Boolean);

          addUnique(teams);
        }

        // BUSINESS UNIT MANAGER
        if (userGroups.includes("business unit manager")) {
          const managedBusinessUnits = allTCC
            .filter(row => row.Manager && row.Manager.Title.toLowerCase() === normalizedUser)
            .map(row => row.Title)
            .filter(Boolean);

          const teams = allTCC
            .filter(row => managedBusinessUnits.includes(row.BusinessUnit))
            .map(row => row.Title)
            .filter(Boolean);

          addUnique(teams);
        }

        // DEPARTMENT MANAGER
        if (userGroups.includes("department manager")) {
          const managedDepartments = allTCC
            .filter(row => row.Manager && row.Manager.Title.toLowerCase() === normalizedUser)
            .map(row => row.Title)
            .filter(Boolean);

          const teams = allTCC
            .filter(row => managedDepartments.includes(row.Department))
            .map(row => row.Title)
            .filter(Boolean);

          addUnique(teams);
        }

        // TEAM MANAGER
        if (userGroups.includes("team manager")) {
          const teams = allTCC
            .filter(row => row.Manager && row.Manager.Title.toLowerCase() === normalizedUser)
            .map(row => row.Title)
            .filter(Boolean);

          addUnique(teams);
        }

        // DEFAULT (if no roles)
        if (allowedTeamsForUser.length === 0) {
          allowedTeamsForUser = [];
        }

        const filteredTCCByAllowed = allTCC.filter(row => allowedTeamsForUser.includes(row.Title));

        const businessSet = new Set<string>();
        const businessUnitSet = new Set<string>();
        const departmentSet = new Set<string>();
        const teamSet = new Set<string>();

        filteredTCCByAllowed.forEach(row => {
          if (row.Business) businessSet.add(row.Business);
          if (row.BusinessUnit) businessUnitSet.add(row.BusinessUnit);
          if (row.Department) departmentSet.add(row.Department);
          if (row.Title) teamSet.add(row.Title);
        });

        // Map projects to their Business, Business Unit, Department through PM team
        const projectToBusinessMap: { [project: string]: string } = {};
        const projectToBusinessUnitMap: { [project: string]: string } = {};
        const projectToDepartmentMap: { [project: string]: string } = {};

        filteredTCCByAllowed.forEach(row => {
          Object.entries(projectTeamMap).forEach(([projectTitle, team]) => {
            if (team === row.Title) {
              projectToBusinessMap[projectTitle] = row.Business || "";
              projectToBusinessUnitMap[projectTitle] = row.BusinessUnit || "";
              projectToDepartmentMap[projectTitle] = row.Department || "";
            }
          });
        });

        setBusinessOptions(
          [{ key: "__all__", text: "All" }, ...Array.from(businessSet).map(b => ({ key: b, text: b }))]
        );
        setBusinessUnitOptions(
          [{ key: "__all__", text: "All" }, ...Array.from(businessUnitSet).map(bu => ({ key: bu, text: bu }))]
        );
        setDepartmentOptions(
          [{ key: "__all__", text: "All" }, ...Array.from(departmentSet).map(dep => ({ key: dep, text: dep }))]
        );
        setTeamOptions(
          [{ key: "__all__", text: "All" }, ...Array.from(teamSet).map(dep => ({ key: dep, text: dep }))]
        )
        setProjectToBusinessMap(projectToBusinessMap);
        setProjectToBusinessUnitMap(projectToBusinessMap);
        setProjectToDepartmentMap(projectToDepartmentMap);

        console.log(projectToBusinessMap);
        console.log(projectToBusinessMap);
        console.log(projectToDepartmentMap);

        // Also set dropdown options for teams using allowedTeamsForUser as before
        setAllowedTeams(Array.from(new Set(allowedTeamsForUser)));
      } catch (err) {
        console.error("Error loading projects and roles", err);
        setError(err.message || String(err));
      } finally {
        setLoading(false);
      }
    }
    if (userGroups.length > 0 && currentUser) {
      loadProjectsAndTeams();
    }
  }, [userGroups, currentUser, projectsp]);

  useEffect(() => {
    async function loadProjects() {
      const projectItems = await projectsp.web
        .lists.getByTitle("Projects")
        .items.select("Id", "Title", "Department")
        .top(4999)();
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
          (r) => r.Status && r.Status.toLowerCase() === "overdue"
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

  const poItemSelection = React.useRef(new Selection({
    onSelectionChanged: () => {
      const selectedItems = poItemSelection.current.getSelection();
      if (selectedItems.length === 0) {
        setSelectedStatusFilter(null);
        setSelectedCurrentStatusFilter(null);
        setSelectedPOItemTitle(null);
      } else {
        setSelectedPOItemTitle((selectedItems[0] as POItem).POItem_x0020_Title);
      }
    }
  }));
  // Memoized filtered projects based on allowed teams
  const filteredProjectsByTeam = React.useMemo(() => {
    if (allowedTeams === null) return [];
    if (!allowedTeams.length) return [];
    return projects.filter(proj => {
      const team = projectToTeamMap[proj.Title];
      return allowedTeams.includes(team);
    });
  }, [projects, projectToTeamMap, allowedTeams]);

  const departmentFilteredProjects = React.useMemo(() => {
    if (!allowedTeams) return;
    let projs = filteredProjectsByTeam;

    // Filter by Business if selected and not '__all__'
    if (selectedBusiness && selectedBusiness !== "__all__")
      projs = projs.filter(p => {
        const business = projectToBusinessMap[p.Title];
        return business === selectedBusiness;
      });

    if (selectedBusinessUnit && selectedBusinessUnit !== "__all__")
      projs = projs.filter(p => {
        const businessUnit = projectToBusinessUnitMap[p.Title];
        return businessUnit === selectedBusinessUnit;
      });

    // Filter by Department if selected and not '__all__'
    if (selectedDepartment && selectedDepartment !== "__all__") {
      projs = projs.filter(p => p.Department === selectedDepartment);
    }

    // Filter by Team if selected and not '__all__'
    if (selectedTeam && selectedTeam !== "__all__") {
      projs = projs.filter(p => projectToTeamMap[p.Title] === selectedTeam);
    }

    return projs;
  }, [filteredProjectsByTeam, selectedBusiness, selectedBusinessUnit, selectedDepartment, selectedTeam, projectToBusinessMap, projectToBusinessUnitMap, projectToTeamMap]);


  const filteredPOList = React.useMemo(() => {
    console.log("Filtering PO List with allowedTeams:", allowedTeams);
    if (allowedTeams === null) {
      console.log("allowedTeams still null - returning empty PO list");
      return [];
    }
    if (!allowedTeams.length) {
      console.log("allowedTeams empty - returning empty PO list");
      return [];
    }
    const filtered = poList.filter(po => {
      const team = projectToTeamMap[po.ProjectName];
      const allowed = allowedTeams.includes(team);
      console.log("Allowed Teams:", allowedTeams);
      console.log(`PO ${po.POID} with ProjectName ${po.ProjectName} maps to team ${team}, allowed: ${allowed}`);
      return allowed;
    });
    console.log(`Filtered PO List count: ${filtered.length}`);
    return filtered;
  }, [poList, projectToTeamMap, allowedTeams]);

  // Create PO summary with invoiced and percentage values
  const poSummary = React.useMemo(() => {
    return filteredPOList.map(po => {
      const related = reqList.filter(
        r => r.PurchaseOrder === po.POID && r.Status && r.Status.toLowerCase() !== "cancelled"
      );
      const invoiced = related.reduce((s, r) => s + (+r.InvoiceAmount || 0), 0);
      const paid = related
        .filter(r => r.Status?.toLowerCase() === "payment received")
        .reduce((s, r) => s + (+r.InvoiceAmount || 0), 0);
      return {
        ...po,
        invoiced,
        paidAmount: paid,
        percentPaid: !po.POAmount ? 0 : (invoiced / +po.POAmount) * 100,
      };
    });
  }, [filteredPOList, reqList]);

  // Search filter on PO summary
  const searchFilteredPoSummary = React.useMemo(() => {
    return poSummary.filter(po => poMatchesSearch(po, searchText));
  }, [poSummary, searchText]);

  const filteredPoSummary = React.useMemo(() => {
    console.log('Calculating filteredPoSummary memo');

    if (allowedTeams === null) {
      console.log('allowedTeams is null, returning empty array');
      return [];
    }

    console.log('allowedTeams:', allowedTeams);
    console.log('searchFilteredPoSummary length before filtering:', searchFilteredPoSummary.length);
    console.log('departmentFilteredProjects length:', departmentFilteredProjects.length);

    const filteredByProject = searchFilteredPoSummary.filter(po => {
      const matchingProject = departmentFilteredProjects.find(p => {
        console.log('Comparing:', { projectTitle: p.Title, poProjectName: po.ProjectName });
        return p.Title === po.ProjectName;
      });

      const isMatch = !!matchingProject;
      if (!isMatch) {
        console.log(`PO with ProjectName "${po.ProjectName}" has no matching project`);
      }
      return isMatch;
    });

    console.log('After filtering by project, count:', filteredByProject.length);

    const filteredByParentPOID = filteredByProject.filter(po => {
      const hasNoParent = !po.ParentPOID;
      if (!hasNoParent) {
        console.log(`Excluding PO with ParentPOID: ${po.ParentPOID} for PO:`, po);
      }
      return hasNoParent;
    });

    console.log('Final filteredPoSummary count:', filteredByParentPOID.length);

    return filteredByParentPOID;

  }, [searchFilteredPoSummary, departmentFilteredProjects, allowedTeams]);


  const filteredPoSummaryWithStatusFilter = React.useMemo(() => {
    if (!filteredPoSummary || filteredPoSummary.length === 0) return [];

    if (!selectedStatusFilter && !selectedCurrentStatusFilter) {
      return filteredPoSummary;
    }

    return filteredPoSummary.filter(po => {
      // Get related invoice requests for this PO
      const relatedReqs = reqList.filter(req =>
        req.PurchaseOrder === po.POID &&
        (!selectedStatusFilter || req.Status === selectedStatusFilter) &&
        (!selectedCurrentStatusFilter || req.CurrentStatus === selectedCurrentStatusFilter) &&
        req.Status?.toLowerCase() !== "cancelled"
      );

      // Only include this PO if it has any related invoice requests passing the filter
      return relatedReqs.length > 0;
    });
  }, [filteredPoSummary, reqList, selectedStatusFilter, selectedCurrentStatusFilter]);

  const filteredInvoiceRequests = React.useMemo(() => {
    if (!filteredPoSummary || filteredPoSummary.length === 0) return [];

    const visiblePOIDs = new Set(filteredPoSummary.map(po => po.POID));

    let filtered = reqList.filter(req => req.PurchaseOrder && visiblePOIDs.has(req.PurchaseOrder));
    if (selectedCurrentStatusFilter) {
      filtered = filtered.filter(req => (req.CurrentStatus ?? "Unknown") === selectedCurrentStatusFilter);
    }

    if (selectedPOItemTitle) {
      filtered = filtered.filter(req => req.POItemx0020Title === selectedPOItemTitle);
    }
    return filtered;
  }, [filteredPoSummary, reqList, selectedStatusFilter, selectedCurrentStatusFilter]);

  const filteredInvoiceRequestsPanel = selectedPOItemTitle
    ? poPanel.invoiceRequests.filter(req => req.POItem_x0020_Title === selectedPOItemTitle)
    : poPanel.invoiceRequests;


  const filteredStatusMap = React.useMemo(() => {
    const statusCounts: { [status: string]: number } = {};
    filteredInvoiceRequests.forEach(req => {
      const key = req.Status || "Unknown";
      statusCounts[key] = (statusCounts[key] || 0) + 1;
    });
    return statusCounts;
  }, [filteredInvoiceRequests]);

  const filteredCurrentStatusMap = React.useMemo(() => {
    const currentStatusCounts: { [currentStatus: string]: number } = {};
    filteredInvoiceRequests.forEach(req => {
      const key = req.CurrentStatus || "Unknown";
      currentStatusCounts[key] = (currentStatusCounts[key] || 0) + 1;
    });
    return currentStatusCounts;
  }, [filteredInvoiceRequests]);

  // PO details screen open handler
  function openInvoiceRequestPanel(request: InvoiceRequest) {
    setInvoiceRequestPanel({ open: true, invoiceRequest: request });
  }

  const projectsInDetailList = React.useMemo(() => {
    if (!filteredPoSummary || filteredPoSummary.length === 0) {
      return [];
    }

    const poProjectNames = new Set(filteredPoSummary.map(po => po.ProjectName));
    return projects.filter(proj => poProjectNames.has(proj.Title));
  }, [filteredPoSummary, projects]);

  useEffect(() => {
    if (projectsInDetailList.length === 0) {
      setDepartments([{ key: "__all__", text: "All" }]);
      return;
    }

    const uniqueDepartments = Array.from(
      new Set(projectsInDetailList.map(p => p.Department).filter(Boolean))
    );

    const options: IDropdownOption[] = [
      { key: "__all__", text: "All" },
      ...uniqueDepartments.map(dep => ({ key: dep, text: dep })),
    ];

    setDepartments(options);
  }, [projectsInDetailList]);

  useEffect(() => {
    if (projects.length === 0) {
      setTeams([{ key: "__all__", text: "All" }]);
      return;
    }

    const uniqueTeams = Array.from(
      new Set(
        projectsInDetailList
          .map(p => projectToTeamMap[p.Title])
          .filter(Boolean)
      )
    );

    const options: IDropdownOption[] = [
      { key: "__all__", text: "All" },
      ...uniqueTeams.map(team => ({ key: team, text: team })),
    ];

    setTeams(options);
  }, [projectsInDetailList, projectToTeamMap]);

  function decodeHtml(html: string): string {
    const txt = document.createElement("textarea");
    txt.innerHTML = html;
    return txt.value;
  }
  // Get PO items function
  function getPOItems(po: InvoicePO, allPOs: InvoicePO[]): POItem[] {
    if (po.LineItemsJSON) {
      try {
        const items = decodeHtml(po.LineItemsJSON);
        const decoded = JSON.parse(items);

        if (Array.isArray(decoded)) {
          return decoded.map((li: any, idx: number) => ({
            Id: po.POID,
            POItem_x0020_Title: li.Title || `LineItem${idx + 1}`,
            POItem_x0020_Value: li.Value || "0",
            POComments: li.Comments,
            Currency: po.Currency,
          }));
        }
      } catch {

      }
    }
    // Check for child POs
    const childPOs = allPOs.filter(p => p.ParentPOID === po.POID);
    if (childPOs.length > 0) {
      return childPOs.map(child => ({
        POItem_x0020_Title: child.POID,
        POItem_x0020_Value: child.POAmount || 0,
        POComments: "",
        Currency: child.Currency || po.Currency,
      }));
    }
  }

  // PO panel open handler
  function openPoPanel(po: InvoicePO) {
    const poItems = getPOItems(po, poList);
    const poIdsForRequests = [po.POID];
    poList.forEach(p => {
      if (p.ParentPOID === po.POID) poIdsForRequests.push(p.POID);
    });
    const invoiceRequests = reqList.filter(r => poIdsForRequests.includes(r.PurchaseOrder));
    setPoPanel({ open: true, po, poItems, invoiceRequests });
  }

  // Search helper
  function poMatchesSearch(po: any, text: string): boolean {
    if (!text) return true;
    const lower = text.toLowerCase();
    return [
      po.POID,
      po.ProjectName,
      po.Currency,
      po.POAmount,
      po.invoiced,
      (po.percentPaid ?? 0) + "%",
    ].some(val => (val == null ? "" : val.toString().toLowerCase()).includes(lower));
  }

  // Columns definition
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
      onRender: (i) => getCurrencySymbol(i.Currency) + ((+i.invoiced || 0).toLocaleString()),
    },
    {
      key: "paidamount",
      name: "Paid Amount",
      fieldName: "paidAmount",
      minWidth: 120,
      maxWidth: 150,
      onRender: (item) =>
        getCurrencySymbol(item.Currency) + ((item.paidAmount ?? 0).toLocaleString()),
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

  if (loading || allowedTeams === null)
    return <Spinner label="Loading..." />;

  if (error)
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;

  return (
    <Stack tokens={{ childrenGap: 28 }} styles={{ root: { padding: 32, background: "#fafafa", minHeight: 600 } }}>
      <Separator />
      <Stack horizontal tokens={{ childrenGap: 24, padding: 0 }} styles={{ root: { marginBottom: 16 } }}>
        <div>
          <Label>Search</Label>
          <SearchBox
            placeholder="Search"
            value={searchText}
            onChange={(_, newValue) => setSearchText(newValue ?? "")}
            styles={{ root: { width: 300 } }}
          />
        </div>
        <div>
          <Dropdown
            label="Business"
            placeholder="Business"
            options={businessOptions}
            selectedKey={selectedBusiness}
            onChange={(_, option) => setSelectedBusiness(option?.key as string || "all")}
            styles={{ root: { width: 150 } }}
          />
        </div>
        <div>
          <Dropdown
            label="Business Unit"
            placeholder="Business Unit"
            options={businessUnitOptions}
            selectedKey={selectedBusinessUnit}
            onChange={(_, option) => setSelectedBusinessUnit(option?.key as string || "all")}
            styles={{ root: { width: 150 } }}
          />
        </div>
        <div>
          <Dropdown
            label="Department"
            placeholder="Department"
            options={departmentOptions}
            selectedKey={selectedDepartment}
            onChange={(_, option) => setSelectedDepartment(option?.key as string || "all")}
            styles={{ root: { width: 150 } }}
          />
        </div>
        <div>
          <Dropdown
            label="Team"
            placeholder="Team"
            options={teamOptions}
            selectedKey={selectedTeam}
            onChange={(_, option) => setSelectedTeam(option?.key as string || "all")}
            styles={{ root: { width: 150 } }}
          />
        </div>
      </Stack>
      <Stack>
        <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 6 } }}>
          Invoice Status
        </Text>
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          {Object.entries(filteredStatusMap ?? {}).map(([status, count]) => {
            const isSelected = selectedStatusFilter === status;
            return (
              <Stack
                key={status}
                onClick={() => {
                  setSelectedStatusFilter(isSelected ? null : status);
                  setSelectedCurrentStatusFilter(null); // Clear other filter if needed
                }}
                styles={{
                  root: {
                    minWidth: 100,
                    background: isSelected ? "#d0e7ff" : "#fff",
                    borderRadius: 6,
                    boxShadow: "0 2px 7px #f6f6f6",
                    padding: "10px 14px",
                    margin: "6px 0",
                    cursor: "pointer",
                    userSelect: "none",
                  },
                }}
              >
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 700 } }}>
                  {count}
                </Text>
                <div style={{ color: "#666" }}>{status}</div>
              </Stack>
            );
          })}
        </Stack>
      </Stack>
      <Stack>
        <Text variant="large" styles={{ root: { fontWeight: 600, marginTop: 28, marginBottom: 6 } }}>
          Current Status
        </Text>
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          {Object.entries(filteredCurrentStatusMap ?? {}).map(([currentStatus, count]) => {
            const isSelected = selectedCurrentStatusFilter === currentStatus;
            return (
              <Stack
                key={currentStatus}
                onClick={() => {
                  setSelectedCurrentStatusFilter(isSelected ? null : currentStatus);
                  setSelectedStatusFilter(null); // Clear other filter if needed
                }}
                styles={{
                  root: {
                    minWidth: 100,
                    background: isSelected ? "#d0e7ff" : "#fff",
                    borderRadius: 6,
                    boxShadow: "0 2px 7px #f6f6f6",
                    padding: "10px 14px",
                    margin: "6px 0",
                    cursor: "pointer",
                    userSelect: "none",
                  },
                }}
              >
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 700 } }}>
                  {count}
                </Text>
                <div style={{ color: "#666" }}>{currentStatus}</div>
              </Stack>
            );
          })}
        </Stack>
      </Stack>
      <Separator />
      <Text variant="large" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
        All Purchase Orders (Summary)
      </Text>
      <DetailsList
        items={filteredPoSummaryWithStatusFilter}
        columns={columns}
        compact
        isHeaderVisible
        styles={{ root: { background: "#fff", borderRadius: 8, marginTop: 4 } }}
        setKey="businessSummary"
        selectionMode={0}
        onActiveItemChanged={openPoPanel}
      />
      <Panel
        isOpen={poPanel.open}
        headerText="Purchase Order Summary"
        onDismiss={() =>
          setPoPanel({ open: false, po: null, poItems: [], invoiceRequests: [] })
        }
        closeButtonAriaLabel="Close"
        isLightDismiss
        type={PanelType.largeFixed}
      // styles={{ main: { maxWidth: 1000 } }}
      >
        {poPanel.po && (
          <Stack
            tokens={{ childrenGap: 24 }}
            styles={{
              root: {
                padding: 24,
                background: '#f5f9fc',
                borderRadius: 12,
                minHeight: 700,
              },
            }}
          >
            {/* --- Summary Card --- */}
            <div
              style={{
                background: '#ffffff',
                borderRadius: 10,
                padding: '22px 28px',
                boxShadow: '0 3px 10px rgba(0,0,0,0.05)',
                borderLeft: `6px solid ${primaryColor}`,
              }}
            >
              <div
                style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(3, 1fr)',
                  rowGap: '16px',
                  columnGap: '40px',
                }}
              >
                <div>
                  <Text><b>Purchase Order:</b> {poPanel.po.POID}</Text>
                </div>
                <div>
                  <Text><b>Invoiced Amount:</b> {getCurrencySymbol(poPanel.po.Currency)}{poPanel.po.invoiced?.toLocaleString()}</Text>
                </div>
                <div>
                  <Text><b>Paid Amount:</b> {getCurrencySymbol(poPanel.po.Currency)}{poPanel.po.paidAmount?.toLocaleString()}</Text>
                </div>

                <div>
                  <Text><b>Project Name:</b> {poPanel.po.ProjectName}</Text>
                </div>
                <div>
                  <Text><b>PO Amount:</b> {getCurrencySymbol(poPanel.po.Currency)}{poPanel.po.POAmount?.toLocaleString()}</Text>
                </div>
                <div>
                  <Text>
                    <b>Invoiced %:</b>{' '}
                    <span
                      style={{
                        color:
                          poPanel.po.percentPaid >= 100
                            ? '#28a745'
                            : poPanel.po.percentPaid >= 50
                              ? '#f5a623'
                              : '#0078d4',
                        fontWeight: 600,
                      }}
                    >
                      {poPanel.po.percentPaid?.toFixed(0)}%
                    </span>
                  </Text>
                </div>
              </div>
            </div>

            {/* --- PO Items Table --- */}
            <Stack>
              <Text
                variant="large"
                styles={{
                  root: {
                    fontWeight: 600,
                    color: primaryColor,
                    marginBottom: 8,
                    borderBottom: `2px solid ${primaryColor}`,
                    width: 'fit-content',
                    paddingBottom: 4,
                  },
                }}
              >
                PO Items
              </Text>
              <div
                style={{
                  background: '#fff',
                  borderRadius: 8,
                  padding: 14,
                  boxShadow: '0 2px 6px rgba(0,0,0,0.05)',
                }}
              >
                <DetailsList
                  items={poPanel.poItems}
                  columns={poItemColumns}
                  selectionMode={1}
                  onActiveItemChanged={(item) => {
                    setSelectedPOItemTitle(item.POItem_x0020_Title);
                  }}
                  selection={poItemSelection.current}
                  compact
                  isHeaderVisible
                  styles={{
                    root: {
                      background: 'transparent',
                      selectors: {
                        '.ms-DetailsRow:hover': {
                          background: '#f3f8fe !important',
                        },
                      },
                    },
                  }}
                />
              </div>
            </Stack>

            {/* --- Invoice Requests Table --- */}
            <Stack>
              <div style={{ display: 'flex', justifyContent: 'flex-end', width: '100%' }}>
                <PrimaryButton
                  text="Show all Invoice requests"
                  onClick={() => {
                    setSelectedStatusFilter(null);
                    setSelectedCurrentStatusFilter(null);
                    setSelectedPOItemTitle(null);
                  }}
                  styles={{ root: { marginLeft: 24, backgroundColor: primaryColor } }}
                  disabled={!selectedPOItemTitle}
                />
              </div>
              <Text
                variant="large"
                styles={{
                  root: {
                    fontWeight: 600,
                    color: primaryColor,
                    marginBottom: 8,
                    borderBottom: `2px solid ${primaryColor}`,
                    width: 'fit-content',
                    paddingBottom: 4,
                  },
                }}
              >
                Invoice Requests for PO {poPanel.po.POID}
              </Text>
              <div
                style={{
                  background: '#fff',
                  borderRadius: 8,
                  padding: 14,
                  boxShadow: '0 2px 6px rgba(0,0,0,0.05)',
                }}
              >
                {poPanel.invoiceRequests.length > 0 ? (
                  <DetailsList
                    items={filteredInvoiceRequestsPanel}
                    columns={invoiceColumns}
                    compact
                    isHeaderVisible
                    onActiveItemChanged={openInvoiceRequestPanel}
                    styles={{
                      root: {
                        background: 'transparent',
                        selectors: {
                          '.ms-DetailsRow:hover': {
                            background: '#f3f8fe !important',
                          },
                        },
                      },
                    }}
                  />
                ) : (
                  <Text styles={{ root: { color: '#888' } }}>
                    No invoice requests for this PO
                  </Text>
                )}
              </div>
            </Stack>
          </Stack>
        )
        }
        <Panel
          isOpen={invoiceRequestPanel.open}
          onDismiss={() => setInvoiceRequestPanel({ open: false, invoiceRequest: null })}
          headerText={`Invoice Request Details:`}
          closeButtonAriaLabel="Close"
          isLightDismiss
          type={PanelType.largeFixed}
        // styles={{ main: { maxWidth: 620 } }}
        >
          {invoiceRequestPanel.invoiceRequest && (
            <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: 16, background: "#f4f9fc", borderRadius: 10 } }}>

              {/* Main Details Card */}
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "155px 1fr 155px 1fr",
                  rowGap: 14,
                  columnGap: 22,
                }}
              >
                <div style={{ fontWeight: 600, color: primaryColor }}>Purchase Order:</div>
                <div>{invoiceRequestPanel.invoiceRequest.PurchaseOrder}</div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Project Name:</div>
                <div>{invoiceRequestPanel.invoiceRequest.ProjectName ?? "-"}</div>

                {/* <div style={{ fontWeight: 600, color: primaryColor }}>PO Item Title:</div>
                <div>{invoiceRequestPanel.invoiceRequest.POItem_x0020_Title ?? "-"}</div>
                <div style={{ fontWeight: 600, color: primaryColor }}>PO Item Value:</div>
                <div>
                  {getCurrencySymbol(invoiceRequestPanel.invoiceRequest.Currency)}
                  {invoiceRequestPanel.invoiceRequest.POItem_x0020_Value?.toLocaleString() ?? "-"}
                </div> */}

                <div style={{ fontWeight: 600, color: primaryColor }}>Invoiced Amount:</div>
                <div>
                  {getCurrencySymbol(invoiceRequestPanel.invoiceRequest.Currency)}
                  {invoiceRequestPanel.invoiceRequest.InvoiceAmount?.toLocaleString() ?? "-"}
                </div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Invoice Status:</div>
                <div>
                  <span
                    style={{
                      fontWeight: 700,
                      background: "#e5f1fa",
                      color: primaryColor,
                      borderRadius: 12,
                      padding: "2px 14px",
                      display: "inline-block",
                    }}
                  >
                    {invoiceRequestPanel.invoiceRequest.Status ?? "-"}
                  </span>
                </div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Current Status:</div>
                <div>{invoiceRequestPanel.invoiceRequest.CurrentStatus ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Due Date:</div>
                <div>
                  {invoiceRequestPanel.invoiceRequest.DueDate
                    ? new Date(invoiceRequestPanel.invoiceRequest.DueDate).toLocaleDateString()
                    : "-"}
                </div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Created:</div>
                <div>
                  {invoiceRequestPanel.invoiceRequest.Created
                    ? new Date(invoiceRequestPanel.invoiceRequest.Created).toLocaleDateString()
                    : "-"}
                </div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Created By:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Author?.Title ?? "-"}</div>

                <div style={{ fontWeight: 600, color: primaryColor }}>Modified:</div>
                <div>
                  {invoiceRequestPanel.invoiceRequest.Modified
                    ? new Date(invoiceRequestPanel.invoiceRequest.Modified).toLocaleDateString()
                    : "-"}
                </div>
                <div style={{ fontWeight: 600, color: primaryColor }}>Modified By:</div>
                <div>{invoiceRequestPanel.invoiceRequest.Editor?.Title ?? "-"}</div>
              </div>

              {/* PM Comments Section */}
              {invoiceRequestPanel.invoiceRequest.PMCommentsHistory && (
                <div style={{
                  background: "#fcfcfd",
                  borderRadius: 7,
                  padding: 14,
                  marginBottom: 10,
                  border: "1px solid #edf3fa"
                }}>
                  <Text variant="medium" styles={{ root: { fontWeight: 600, color: primaryColor, marginBottom: 4 } }}>Requestor Comments</Text>
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
                        <a
                          href="#"
                          onClick={e => {
                            e.preventDefault();
                            setAttachmentViewer({
                              isOpen: true,
                              url: file.ServerRelativeUrl,
                              fileName: file.FileName,
                            });
                          }}
                          style={{
                            color: primaryColor,
                            textDecoration: 'underline',
                            fontWeight: 500,
                            fontSize: 15,
                            cursor: 'pointer',
                          }}
                        >
                          {file.FileName}
                        </a>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <Text styles={{ root: { color: "#888", marginTop: 3 } }}>No attachments</Text>
                )}
              </div>
              <Panel
                isOpen={attachmentViewer.isOpen}
                onDismiss={() => setAttachmentViewer({ isOpen: false, url: '', fileName: '' })}
                headerText="Attachment Preview"
                type={PanelType.large}
                closeButtonAriaLabel="Close"
              >
                <div style={{ height: "100%", width: "100%" }}>
                <DocumentViewer
                  url={attachmentViewer.url}
                  fileName={attachmentViewer.fileName}
                  isOpen={attachmentViewer.isOpen}
                  onDismiss={() => setAttachmentViewer({ isOpen: false, url: '', fileName: '' })}
                />
                </div>
              </Panel>
            </Stack>
          )}
        </Panel>
      </Panel >
    </Stack >
  );
}
