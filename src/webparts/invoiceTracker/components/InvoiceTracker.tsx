import * as React from "react";
import { INavLink, INavLinkGroup, Nav, Persona, PersonaSize, ProgressIndicator, Icon } from "office-ui-fabric-react";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import Dashboard from "./Dashboard";
import CreateView from "./Create View/CreateView";
import Settings from "./Settings/Settings";
import MyRequests from "./MyRequests/MyRequests";
import FinanceView from "./Finance View/FinanceView";
import Home from "./Home/Home";
import ManageMembers from "./ManageMembers/ManageMembers";
import styles from "./InvoiceTracker.module.scss";
import Logo from "../assets/Logo.png";
import BusinessView from "./BusinessView/BusinessView";
import Help from "./Help/Help";
export interface IInvoiceTrackerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: any;
  userDisplayName: string;
  onConfigChange?: (settings: any) => void;
  selectedSites: string[];
  projectSiteUrl?: string;
  getCurrentPageUrl?: () => string;
  pageConfig?: Record<string, boolean>;
}
export const InvoiceTrackerContext: any = React.createContext(React.useContext);
interface IInvoiceTrackerState {
  loading: boolean;
  progress: number;
  selectedTab: string;
  navLinks: INavLinkGroup[];
  userRoles: string[];
  isAdminUser: boolean;
  pendingRequests: number;
  paymentPending: number;
  filter?: {
    [key: string]: any;
  };
  clarificationCount?: number;
  userGroups: string[];
  isNavCollapsed: boolean;
  pageConfig: Record<string, boolean>;
}
const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";
const navStyles = {
  root: {
    overflowY: "auto",
    flexGrow: 1
  },
  link: {
    color: primaryColor, // normal state text
    selectors: {
      ".ms-Icon": { color: primaryColor }
    }
  },
  linkText: {
    color: primaryColor
  },
  linkSelected: {
    color: primaryColor, // selected text color
    borderLeft: '4px solid #7d0c71', // selected left border
    background: "#f9f5fa" // lighter purple background (optional)
  },
  linkHovered: {
    color: primaryColor,
    background: "#f9f5fa" // optional: lighter hover background
  },
  chevronButton: {
    color: primaryColor
  }
};

export default class InvoiceTracker extends React.Component<IInvoiceTrackerProps, IInvoiceTrackerState> {
  public sp: SPFI;
  private totalSteps = 4;
  private completedSteps = 0;
  public projectSp: SPFI;

  constructor(props: IInvoiceTrackerProps) {
    super(props);
    this.state = {
      loading: true,
      progress: 0,
      selectedTab: "home",
      navLinks: [],
      userRoles: [],
      isAdminUser: false,
      pendingRequests: 0,
      paymentPending: 0,
      userGroups: [],
      isNavCollapsed: false,
      pageConfig: props.pageConfig || {},
    };
    this.sp = spfi().using(SPFx(this.props.context));
    this.setCanvasParentStyles();
    this.projectSp = this.props.projectSiteUrl
      ? spfi(this.props.projectSiteUrl).using(SPFx(this.props.context))
      : this.sp;   // fallback to current if not set

  }

  private updateProgress() {
    this.completedSteps++;
    const newProgress = (this.completedSteps / this.totalSteps) * 100;
    this.setState({ progress: newProgress });
  }

  private toggleNavCollapse = () => {
    this.setState(prev => ({ isNavCollapsed: !prev.isNavCollapsed }));
  };

  public async componentDidMount() {

    // window.addEventListener("resize", this.handleResize);
    // this.handleResize();

    window.addEventListener("popstate", this.onPopState);
    await this.fetchUserGroups();

    // const hash = window.location.hash;
    // const initialTab = hash ? hash.replace("#", "") : "home";
    // window.history.replaceState({ selectedTab: initialTab }, "", `#${initialTab}`);

    // this.setState({ selectedTab: initialTab });
    const fragment = window.location.hash.substring(1); // 'myrequests?selectedInvoice=13'
    const [tab, query] = fragment.split("?");
    const params = new URLSearchParams(query || "");
    const selectedInvoice = params.get("selectedInvoice");

    window.history.replaceState(
      { selectedTab: tab || "home", filter: { selectedInvoice } },
      "",
      window.location.href
    );

    this.setState({
      selectedTab: tab || "home",
      filter: selectedInvoice ? { selectedInvoice } : undefined,
    });


    const canvas = document.querySelector('.CanvasSection');
    if (canvas && canvas.parentElement) {
      canvas.parentElement.style.width = "100%";
      canvas.parentElement.style.minWidth = "100%";
      canvas.parentElement.style.maxWidth = "100%";
      canvas.parentElement.style.position = "fixed";
      canvas.parentElement.style.top = "0"
      canvas.parentElement.style.left = "0";
      canvas.parentElement.style.height = "100%";
      canvas.parentElement.style.zIndex = "1000";
      canvas.parentElement.style.margin = "0";
      canvas.parentElement.style.background = "#fff";
    }

    this.setState({ loading: true, progress: 10 });
    const roles = await this.getUserRoles();
    const isAdmin = roles.includes("admin");
    const isFinance = roles.includes("Finance");
    const isPM = roles.includes("PM") || roles.includes("Project Manager");
    const isDM = roles.includes("DM") || roles.includes("Delivery Manager");
    const isDH = roles.includes("DH") || roles.includes("Department Head");
    const isBusinessM = roles.includes("Business");
    const isBusinessUnitM = roles.includes("Business Unit");
    const isDepartmentM = roles.includes("Department");
    const isTeamM = roles.includes("Team");
    function decodeHtmlEntities(str: string): string {
      const txt = document.createElement('textarea');
      txt.innerHTML = str;
      return txt.value;
    }

    try {
      const items = await this.sp.web.lists.getByTitle("InvoiceConfiguration").items
        .filter(`Title eq 'PageConfig'`)
        .top(1)();

      if (items.length > 0 && items[0].Value) {
        const decodedValue = decodeHtmlEntities(items[0].Value);
        const config = JSON.parse(decodedValue);
        this.setState({ pageConfig: config });
        this.applyPageSettings(config);
      } else {
        this.applyPageSettings({});
      }
    } catch (e) {
      console.error("Error loading pageConfig:", e);
      this.applyPageSettings({});
    }

    this.setState({ userRoles: roles, isAdminUser: isAdmin });
    this.updateProgress();

    // Use new role info to build nav links:
    const navLinks = this.getNavLinks(isAdmin, isFinance, isPM, isDM, isDH, isBusinessM, isBusinessUnitM, isDepartmentM, isTeamM, this.state.isNavCollapsed);
    this.setState({ navLinks, loading: false, progress: 100 });
    this.updateProgress();
    await this.ensureGroups();

    await this.ensureInvoiceConfigList();
    this.updateProgress();

    await this.ensureLists();
    await this.ensureInvoicePOList();
    this.updateProgress();

    await this.ensureConfigList();
    this.updateProgress();

    await this.loadConfiguration();
    this.updateProgress();

    this.setState({ navLinks, loading: false, progress: 100 });

    this.setCanvasParentStyles();
  }

  public componentWillUnmount() {
    window.removeEventListener("popstate", this.onPopState);
    // window.removeEventListener("resize", this.handleResize);
  }

  componentDidUpdate(prevProps: IInvoiceTrackerProps, prevState: IInvoiceTrackerState) {
    if (prevState.pageConfig !== this.state.pageConfig) {
      this.applyPageSettings(this.state.pageConfig);
    }
  }


  applyPageSettings = (config: Record<string, boolean>) => {
    const pageSettings = [
      { stateVariable: "hideCommandBar", selectors: ["#spCommandBar"] },
      { stateVariable: "hideSideAppBar", selectors: ["#sp-appBar"] },
      { stateVariable: "hidePageTitle", selectors: ["[id*='PageTitle']"] },
      { stateVariable: "hideSiteHeader", selectors: ["#spSiteHeader", "#spLeftNav"] },
      { stateVariable: "hideCommentsWrapper", selectors: ["#CommentsWrapper"] },
      { stateVariable: "hideO365BrandNavbar", selectors: ["#SuiteNavWrapper"] },
      { stateVariable: "hideSharepointHubNavbar", selectors: [".ms-HubNav"] },
    ];

    pageSettings.forEach(ps => {
      const hide = config[ps.stateVariable];
      ps.selectors.forEach(selector => {
        const el = document.querySelector(selector);
        if (el && el instanceof HTMLElement) {
          el.style.setProperty("display", hide ? "none" : "", "important");
        }
      });
    });
  };



  // private handleResize = () => {
  //   this.setState({
  //     windowHeight: window.innerHeight,
  //     windowWidth: window.innerWidth,
  //   });
  // };

  handleConfigChange = (newConfig: Record<string, boolean>) => {
    this.setState({ pageConfig: newConfig });
    // Optionally, persist config to SharePoint here.
  };


  private async getUserRoles(): Promise<string[]> {
    try {
      const groups = await this.sp.web.currentUser.groups();
      return groups.map(g => g.Title);
    } catch {
      return ["User"];
    }
  }
  private async fetchUserGroups() {
    try {
      const groups = await this.sp.web.currentUser.groups();
      // Filter out groups whose titles end with Members, Owners, or Visitors
      const filteredGroupTitles = groups
        .map(g => g.Title)
        .filter(title =>
          !(
            title.endsWith('Members') ||
            title.endsWith('Owners') ||
            title.endsWith('Visitors')
          )
        );
      this.setState({ userGroups: filteredGroupTitles });
    } catch {
      this.setState({ userGroups: [] });
    }
  }

  private onPopState = (event: PopStateEvent) => {
    const state = event.state as { selectedTab?: string; filter?: any } | null;
    const newTab = state?.selectedTab || 'home';

    if (this.state.selectedTab !== newTab) {
      this.setState({ selectedTab: newTab, filter: state?.filter });
    }
  };

  private async ensureInvoiceConfigList() {
    try {
      await this.sp.web.lists.getByTitle("InvoiceConfiguration").select("Id")();
    } catch (error: any) {
      if (error.status === 404) {
        const added = await this.sp.web.lists.add("InvoiceConfiguration", "Configuration for Invoice Tracker", 100);
        const list = this.sp.web.lists.getById(added.Id)
        await list.fields.addMultilineText("Value");
      }
    }
    // const value = { "hideCommandBar": true, "hideSideAppBar": true, "hidePageTitle": true, "hideSiteHeader": true, "hideCommentsWrapper": true, "hideO365BrandNavbar": true, "hideSharepointHubNavbar": true }
    const valueObj = {
      hideCommandBar: true,
      hideSideAppBar: true,
      hidePageTitle: true,
      hideSiteHeader: true,
      hideCommentsWrapper: true,
      hideO365BrandNavbar: true,
      hideSharepointHubNavbar: true
    };

    const valueString = JSON.stringify(valueObj);
    await this.sp.web.lists.getByTitle("InvoiceConfiguration").items.add({
      Title: "PageConfig",
      Value: valueString,
    })
    await this.sp.web.lists.getByTitle("InvoiceConfiguration").items.add({
      Title: "FinanceEmail",
      Value: "",
    })
  }

  private async ensureLists() {
    try {
      await this.sp.web.lists.getByTitle("Invoice Requests").select("Id")();
    } catch (error: any) {
      if (error.status === 404) {
        const addResult = await this.sp.web.lists.add("Invoice Requests", "Stores invoice requests", 100); // 100 = Custom List
        const list = this.sp.web.lists.getById(addResult.Id);

        // Add fields
        await list.fields.addText("PurchaseOrder", { MaxLength: 255 });
        await list.fields.addText("ProjectName", { MaxLength: 255 });
        await list.fields.addText("POItem Title", { MaxLength: 255 });    // SharePoint internal name for space: _x0020_
        await list.fields.addNumber("POItem Value");
        await list.fields.addNumber("InvoiceAmount");
        await list.fields.addText("Customer Contact", { MaxLength: 255 });
        await list.fields.addMultilineText("Comments");

        await list.fields.addText("Status", { MaxLength: 255 });
        await list.fields.addText("InvoiceNumber", { MaxLength: 255 });
        await list.fields.addMultilineText("FinanceComments");

        await list.fields.addText("PMStatus", { MaxLength: 255 });
        await list.fields.addText("FinanceStatus", { MaxLength: 255 });
        await list.fields.addMultilineText("PMCommentsHistory");
        await list.fields.addMultilineText("FinanceCommentsHistory");

        await list.fields.addNumber("POAmount");
        await list.fields.addText("CurrentStatus", { MaxLength: 255 });
        await list.fields.addText("Currency")
        await list.fields.addDateTime("DueDate")
        await list.fields.getByInternalNameOrTitle("Title").update({ Title: "Title", Required: false, Hidden: true });
      }
    }
  }

  private async ensureInvoicePOList() {
    try {
      // Try to get the list by title
      await this.sp.web.lists.getByTitle("InvoicePO").select("Id")();
    } catch (error: any) {
      if (error.status === 404) {
        // List not found, so create it
        const addResult = await this.sp.web.lists.add("InvoicePO", "Stores Purchase Order records", 100); // 100 = Custom List
        const list = this.sp.web.lists.getById(addResult.Id);

        // Add required fields
        await list.fields.addText("POID", { MaxLength: 255 });
        await list.fields.addText("ParentPOID", { MaxLength: 255 });
        await list.fields.addNumber("POAmount");
        await list.fields.addMultilineText("LineItemsJSON");
        await list.fields.addText("ProjectName", { MaxLength: 255 });
        await list.fields.addText("Currency");
        await list.fields.addMultilineText("POComments");
        await list.fields.addText("Customer");
        await list.fields.getByInternalNameOrTitle("Title").update({ Title: "Title", Required: false, Hidden: true });
      } else {
        throw error; // rethrow if not a 'not found' error
      }
    }
  }

  private async ensureConfigList() {
    try {
      // Try to get the list by title
      await this.sp.web.lists.getByTitle("InvoiceConfiguration")();
    } catch (error: any) {
      if (error.status === 404) {
        //   // List not found, so create it
        //   const addResult = await this.sp.web.lists.add("InvoiceConfiguration", "Stores Configuration", 100); // 100 = Custom List
        //   const list = this.sp.web.lists.getById(addResult.Id);

        //   await list.fields.addMultilineText("Value");
        // } else {
        //   throw error; // rethrow if not a 'not found' error
        // }
        const defaultConfig = {
          hideCommandBar: false,
          hideSideAppBar: false,
          hidePageTitle: false,
          hideSiteHeader: false,
          hideCommentsWrapper: false,
          hideO365BrandNavbar: false,
          hideSharepointHubNavbar: false
        };

        const list = this.sp.web.lists.getByTitle("InvoiceConfiguration");
        const pageConfigItems = await list.items.filter(`Title eq 'PageConfig'`).top(1)();
        const financeEmailItems = await list.items.filter(`Title eq 'FinanceEmail'`).top(1)();

        if (pageConfigItems.length === 0) {
          await list.items.add({ Title: "PageConfig", Value: JSON.stringify(defaultConfig) });
        } else {
          await list.items.getById(pageConfigItems[0].Id).update({ Value: JSON.stringify(defaultConfig) });
        }
        if (financeEmailItems.length === 0) {
          await list.items.add({ Title: "FinanceEmail", Value: "" });
        } else {
          await list.items.getById(financeEmailItems[0].Id).update({ Value: "" });
        }
      }
    }
  }
  private async loadConfiguration() {
  }

  private getNavLinks(
    isAdmin: boolean,
    isFinance: boolean,
    isPM: boolean,
    isDM: boolean,
    isDH: boolean,
    isBusinessM: boolean,
    isBusinessUnitM: boolean,
    isDepartmentM: boolean,
    isTeamM: boolean,
    navCollapsed: boolean
  ): INavLinkGroup[] {
    const links: INavLink[] = [];

    const isPMorDMorDH = isPM || isDM || isDH;
    const isBuBUDepTeam = isBusinessM || isBusinessUnitM || isDepartmentM || isTeamM

    if (isPMorDMorDH || isFinance || isAdmin) {
      links.push({ key: "home", name: "Home", iconProps: { iconName: "Home" }, url: "" });
    }
    if (isPMorDMorDH || isAdmin) {
      links.push({ key: "myrequests", name: "My Requests", iconProps: { iconName: "ViewDashboard" }, url: "" });
      links.push({ key: "Createview", name: "Create Invoice Request", iconProps: { iconName: "People" }, url: "" });
    }
    if (isFinance || isAdmin) {
      links.push({ key: "updaterequests", name: "Update Invoice Request", iconProps: { iconName: "Money" }, url: "" });
    }
    if (isBuBUDepTeam || isAdmin) {
      links.push({ key: "businessview", name: "Business View", iconProps: { iconName: "Financial" }, url: "" });
    }
    if (isAdmin) {
      links.push({ key: "settings", name: "Settings", iconProps: { iconName: "Settings" }, url: "" });
      links.push({ key: "managemembers", name: "Manage Members", iconProps: { iconName: "SecurityGroup" }, url: "" });
    }
    if (isPMorDMorDH || isFinance || isAdmin) {
      links.push({ key: "help", name: "Help", iconProps: { iconName: "Help" }, url: "" });
    }
    if (navCollapsed) {
      return [{
        links: links.map(l => ({
          ...l,
          name: "", // Remove name to display only icon
          onRenderNavLink: (link: any) => (
            <div title={link.originalName || link.name} style={{ display: 'flex', justifyContent: 'center' }}>
              <i className={`ms-Icon ms-Icon--${link.iconProps?.iconName}`} aria-hidden="true" style={{ color: primaryColor }} />
            </div>
          )
        }))
      }];
    }
    return [{ links }];
  }


  private async ensureGroups() {
    const groupNames = [
      "admin",
      "PM",      // Project Manager
      "DM",      // Delivery Manager
      "DH",      // Delivery Head
      "Finance",  // Finance users
      "Business Manager",  // Business users
      "Business Unit Manager",
      "Department Manager",
      "Team Manager"
    ];
    for (const groupName of groupNames) {
      try {
        // If group exists, this will succeed
        await this.sp.web.siteGroups.getByName(groupName)();
      } catch (error: any) {
        if (error && error.status === 404) {
          // Group does not exist, create
          await this.sp.web.siteGroups.add({ Title: groupName, Description: `Members of the ${groupName} group.` });
        }
      }
    }
  }


  // private onNavClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
  //   ev?.preventDefault();
  //   if (item) {
  //     this.setState({ selectedTab: item.key });
  //   }
  // };

  private onNavClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
    ev?.preventDefault();
    if (item) {
      this.handleNavigate(item.key);
    }
  };

  private setCanvasParentStyles() {
    const canvasSection = document.querySelector(".CanvasSection");
    if (canvasSection) {
      const parentElement = canvasSection.parentElement;
      if (parentElement) {
        parentElement.style.setProperty("width", "100%", "important");
        parentElement.style.setProperty("min-width", "100%", "important");
        parentElement.style.setProperty("max-width", "100%", "important");
      }
    }
  }

  // private handleNavigate = (pageKey: string, params?: any) => {
  //   if (params?.initialFilters) {
  //     this.setState({ selectedTab: pageKey, filter: params.initialFilters });
  //   } else {
  //     this.setState({ selectedTab: pageKey, filter: undefined });
  //   }
  //   const stateData = { selectedTab: pageKey, filter: params?.initialFilters };
  //   window.history.pushState(stateData, "", `#${pageKey}`);
  // };

  private handleNavigate = (pageKey: string, params?: any) => {
    if (this.state.selectedTab === pageKey && JSON.stringify(this.state.filter) === JSON.stringify(params?.initialFilters)) {
      return; // already on this tab
    }
    this.setState({ selectedTab: pageKey, filter: params?.initialFilters });

    const baseUrl = this.getCurrentPageUrl();
    let newUrl = `${baseUrl}#${pageKey}`;

    if (params?.initialFilters?.selectedInvoice) {
      newUrl += `?selectedInvoice=${params.initialFilters.selectedInvoice}`;
    }

    window.history.pushState({ selectedTab: pageKey, filter: params?.initialFilters }, "", newUrl);
  };

  private getCurrentPageUrl(): string {
    const pathSegments = window.location.pathname.split('/').filter(Boolean);
    const pageName = pathSegments[pathSegments.length - 1];
    const sitePath = pathSegments.slice(0, -1).join('/');
    return `${window.location.origin}/${sitePath}/${pageName}`;
  }

  private renderContent() {
    const { selectedTab, userRoles, isAdminUser } = this.state;
    const isFinance = userRoles.includes("Finance");
    const isPM = userRoles.includes("PM") || userRoles.includes("Project Manager");
    const isDM = userRoles.includes("DM") || userRoles.includes("Delivery Manager");
    const isDH = userRoles.includes("DH") || userRoles.includes("Department Head");
    const isBusinessM = userRoles.includes("Business");
    const isBusinessUnitM = userRoles.includes("Business Unit");
    const isDepartmentM = userRoles.includes("Department");
    const isTeamM = userRoles.includes("Team");
    const isBuBUDepTeam = isBusinessM || isBusinessUnitM || isDepartmentM || isTeamM
    const isPMorDMorDH = isPM || isDM || isDH;

    switch (selectedTab) {
      case "home":
        if (isPMorDMorDH || isFinance || isAdminUser)
          return <Home sp={this.sp} context={this.props.context} onNavigate={this.handleNavigate} />;
        break;
      case "myrequests":
        if (isPMorDMorDH || isAdminUser)
          return <MyRequests sp={this.sp} projectsp={this.projectSp} context={this.props.context} initialFilters={this.state.filter} onNavigate={this.handleNavigate} getCurrentPageUrl={this.getCurrentPageUrl} />;
        break;
      case "Createview":
        if (isPMorDMorDH || isAdminUser)
          return <CreateView sp={this.sp} projectsp={this.projectSp} context={this.props.context} />;
        break;
      case "updaterequests":
        if (isFinance || isAdminUser)
          return <FinanceView sp={this.sp} projectsp={this.projectSp} context={this.props.context} initialFilters={this.state.filter} onNavigate={this.handleNavigate} />;
        break;
      case "businessview":
        if (isBuBUDepTeam || isAdminUser)
          return <BusinessView sp={this.sp} projectsp={this.projectSp} context={this.props.context} onNavigate={this.handleNavigate} />;
        break;
      case "settings":
        if (isAdminUser)
          return <Settings sp={this.sp} context={this.props.context} pageConfig={this.props.pageConfig} onConfigChange={this.handleConfigChange} />;
        break;
      case "managemembers":
        if (isAdminUser)
          return <ManageMembers context={this.props.context} />;
        break;
      case "help":
        if (isPMorDMorDH || isFinance || isAdminUser) {
          return <Help isOpen={true} onDismiss={() => this.setState({ selectedTab: "home" })} />;
        }
        break;
      default:
        return <Dashboard />;
    }
    // If user not authorized for selected tab
    return <div style={{ padding: 40, textAlign: "center", color: "#b00" }}>
      You do not have access to this section.
    </div>;
  }

  public render() {

    if (this.state.loading) {
      return (
        <div className={styles.invoiceTracker} style={{ height: 100, overflow: 'auto' }}>
          <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <ProgressIndicator
              label="Configuring Invoice Tracker..."
              percentComplete={this.state.progress / 100}
              styles={{ root: { width: "50%" } }}
            />
          </div>
        </div>
      );
    }

    return (
      <div className={styles.invoiceTracker}>
        {/* <div className={styles.sidebar}>
          <div className={styles.sidebarHeader}>Invoice Tracker</div>
          <div className={styles.flexGrow}>
            <div style={{ display: "flex", justifyContent: this.state.isNavCollapsed ? "center" : "flex-end", alignItems: "center", marginBottom: 12 }}>
              {this.state.isNavCollapsed ? (
                <i
                  className="ms-Icon ms-Icon--ChevronRight"
                  style={{ cursor: "pointer", fontSize: 20, padding: 4 }}
                  aria-label="Expand"
                  onClick={this.toggleNavCollapse}
                />
              ) : (
                <i
                  className="ms-Icon ms-Icon--ChevronLeft"
                  style={{ cursor: "pointer", fontSize: 20, padding: 4 }}
                  aria-label="Collapse"
                  onClick={this.toggleNavCollapse}
                />
              )}
            </div>

            <Nav
              selectedKey={this.state.selectedTab}
              onLinkClick={this.onNavClick}
              groups={this.state.navLinks}
              styles={{ root: { overflowY: "auto", flexGrow: 1 } }}
            />
          </div>
          <div className={styles.sidebarFooter}>
            <Persona text={this.props.userDisplayName} size={PersonaSize.size24} />
            <div style={{ marginTop: 8, fontSize: 12, color: "#666" }}>
              {`(${this.state.userGroups.join(", ")})`}
            </div>
          </div>
        </div> */}
        <div className={styles.sidebar} style={{ width: this.state.isNavCollapsed ? 56 : 220, transition: "width 0.2s" }}>
          <div className={styles.sidebarHeader}>
            <div style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: this.state.isNavCollapsed ? "center" : "flex-start",
              height: 56, // Control height as per visual needs
              width: '100%'
            }}>
              <img
                src={Logo}
                alt="Invoice Tracker Logo"
                style={{
                  height: 40,
                  width: 40,
                  marginRight: this.state.isNavCollapsed ? 0 : 12,
                  display: 'block',
                  transition: 'margin 0.2s'
                }}
              />
              {!this.state.isNavCollapsed && (
                <span style={{
                  fontWeight: 'bold',
                  fontSize: 17,
                  lineHeight: '32px'
                }}>
                  Invoice Tracker
                </span>
              )}
            </div>
          </div>

          <div style={{ display: "flex", flexDirection: "column", alignItems: "self-start", paddingTop: 8 }}>
            {/* Chevron changes with state */}
            {this.state.isNavCollapsed ? (
              <Icon iconName="ChevronRight" style={{ fontSize: 20, marginBottom: 18, cursor: "pointer", color: primaryColor }} onClick={this.toggleNavCollapse} />
            ) : (
              <Icon iconName="ChevronLeft" style={{ fontSize: 20, marginBottom: 18, cursor: "pointer", color: primaryColor }} onClick={this.toggleNavCollapse} />
            )}
          </div>

          <div className={styles.flexGrow}>
            <Nav
              selectedKey={this.state.selectedTab}
              onLinkClick={this.onNavClick}
              groups={this.getNavLinks(
                this.state.isAdminUser,
                this.state.userRoles.includes("Finance"),
                this.state.userRoles.includes("PM") || this.state.userRoles.includes("Project Manager"),
                this.state.userRoles.includes("DM") || this.state.userRoles.includes("Delivery Manager"),
                this.state.userRoles.includes("DH") || this.state.userRoles.includes("Department Head"),
                this.state.userRoles.includes("Business"),
                this.state.userRoles.includes("Business Unit"),
                this.state.userRoles.includes("Department"),
                this.state.userRoles.includes("Team"),
                this.state.isNavCollapsed
              )}
              // styles={{ root: { overflowY: "auto", flexGrow: 1, color: primaryColor } }}
              styles={navStyles}
            />
          </div>
          {!this.state.isNavCollapsed && (
            <div className={styles.sidebarFooter}>
              <Persona text={this.props.userDisplayName} size={PersonaSize.size24} />
              {/* <div style={{ marginTop: 8, fontSize: 12, color: "#666" }}>
                {`(${this.state.userGroups.join(", ")})`}
              </div> */}
              <div style={{ marginTop: 8, fontSize: 12, color: "#666", display: "flex", justifyContent: "center", width: "100%", textAlign: "center" }}>
                {`(${this.state.userGroups.join(", ")})`}
              </div>
            </div>
          )}
        </div>

        <div className={styles.workspace}>{this.renderContent()}</div>
      </div>
    );
  }
}
