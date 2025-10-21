export interface IInvoiceTrackerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: any;
  userDisplayName: string;
  selectedSites: string[];
  initialFilters?: {
    [key: string]: any; // or a stricter shape e.g. { CurrentStatus?: string }
  };
  onNavigate?: (pageKey: string, filter?: any) => void;
  projectSiteUrl?: string;
  getCurrentPageUrl?: () => string;
  pageConfig?: Record<string, boolean>;
}





