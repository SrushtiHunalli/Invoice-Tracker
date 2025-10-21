import * as React from "react";
import {
  Spinner,
  Toggle,
  Text,
  Pivot,
  PivotItem,
  Stack,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  IBasePickerSuggestionsProps,
  IPersonaProps,
} from '@fluentui/react';
import { SPFI, PrincipalType, PrincipalSource } from '@pnp/sp';
import { useState, useEffect } from "react";
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';

const pageSettings = [
  {
    label: "Top Command Bar",
    stateVariable: "hideCommandBar",
    sharepointElement: "#spCommandBar",
    sharepointElements: ["#spCommandBar"],
    tooltip: "Hides the Command Bar (containing New, Share, Edit, etc...)"
  },
  {
    label: "Side App Bar",
    stateVariable: "hideSideAppBar",
    sharepointElement: "#sp-appBar",
    sharepointElements: ["#sp-appBar"],
    tooltip: "Hides the SharePoint Side Navigation Bar"
  },
  {
    label: "Page Title",
    stateVariable: "hidePageTitle",
    sharepointElement: "[id*='PageTitle']",
    sharepointElements: ["[id*='PageTitle']"],
    tooltip: "Hides the Page Title"
  },
  {
    label: "Site Navigation Bar",
    stateVariable: "hideSiteHeader",
    sharepointElement: "#spSiteHeader",
    sharepointElements: ["#spSiteHeader", "#spLeftNav"],
    tooltip: "Hides the SharePoint Site Navigation Bar"
  },
  {
    label: "Comments Section",
    stateVariable: "hideCommentsWrapper",
    sharepointElement: "#CommentsWrapper",
    sharepointElements: ["#CommentsWrapper"],
    tooltip: "Hides the Like/Comment Section"
  },
  {
    label: "O365 Brand Navigation Bar",
    stateVariable: "hideO365BrandNavbar",
    sharepointElement: "#SuiteNavWrapper",
    sharepointElements: ["#SuiteNavWrapper"],
    tooltip: "Hides the O365 Navigation Bar"
  },
  {
    label: "SharePoint Hub Navigation Bar",
    stateVariable: "hideSharepointHubNavbar",
    sharepointElement: ".ms-HubNav",
    sharepointElements: [".ms-HubNav"],
    tooltip: "Hides the SharePoint Hub Navigation"
  },
];


const SETTINGS_LIST = "InvoiceConfiguration";
const PAGE_CONFIG_ITEM_TITLE = "PageConfig";

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  noResultsFoundText: 'No results found',
  mostRecentlyUsedHeaderText: 'Recent',
};
const toggleElement = (selector: string, hide: boolean, retries = 5) => {
  const element = document.querySelector(selector) as HTMLElement | null;
  if (!element) {
    if (retries > 0) {
      setTimeout(() => toggleElement(selector, hide, retries - 1), 500); // retry after 0.5s
    }
    return;
  }
  element.style.setProperty("display", hide ? "none" : "", "important");
};


interface SettingsProps {
  sp: SPFI;
  onSettingsUpdate?: (settings: Record<string, boolean>) => void;
  context?: any;
  pageConfig?: Record<string, boolean>;
  onConfigChange?: (settings: Record<string, boolean>) => void;
}

const Settings: React.FC<SettingsProps> = ({ sp, onSettingsUpdate, pageConfig, context, onConfigChange }) => {
  // const [loading, setLoading] = useState(true);
  const [savingConfig, setSavingConfig] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [financePeople, setFinancePeople] = useState<IPersonaProps[]>([]);
  const [settings, setSettings] = React.useState<Record<string, boolean>>(
    pageSettings.reduce((acc, item) => {
      acc[item.stateVariable] = false;
      return acc;
    }, {} as Record<string, boolean>)
  );
  // const [settings, setSettings] = useState<Record<string, boolean> | null>(null);
  const [loading, setLoading] = useState(true);
  // React.useEffect(() => {
  //   if (pageConfig) {
  //     setSettings(prev => ({ ...prev, ...pageConfig }));
  //   }
  // }, [pageConfig]);

  // Load page settings from InvoiceConfiguration list
  useEffect(() => {
    function decodeHtmlEntities(str: string): string {
      const txt = document.createElement('textarea');
      txt.innerHTML = str;
      return txt.value;
    }
    async function loadPageSettings() {
      try {
        const items = await sp.web.lists.getByTitle(SETTINGS_LIST).items.filter(`Title eq '${PAGE_CONFIG_ITEM_TITLE}'`).top(1)();
        if (items.length > 0 && items[0].Value) {
          const decoded = decodeHtmlEntities(items[0].Value);
          const settingValues = JSON.parse(decoded);
          setSettings(settingValues);
        } else {
          setSettings(
            pageSettings.reduce((acc, item) => {
              acc[item.stateVariable] = false;
              return acc;
            }, {} as Record<string, boolean>)
          );
        }
      } catch {
        setSettings(
          pageSettings.reduce((acc, item) => {
            acc[item.stateVariable] = false;
            return acc;
          }, {} as Record<string, boolean>)
        );
      }
      setLoading(false);
    }
    loadPageSettings();
  }, [sp]);

  // Apply page toggles on settings change
  useEffect(() => {
    pageSettings.forEach(ps => {
      const hide = settings[ps.stateVariable];
      const element = document.querySelector(ps.sharepointElement);
      if (element instanceof HTMLElement) {
        element.style.setProperty("display", hide ? "none" : "", "important");
      }
      toggleElement(ps.sharepointElement, settings[ps.stateVariable]);
    });
  }, [settings]);

  // Load Finance Emails from InvoiceConfiguration list
  useEffect(() => {
    async function loadFinanceEmails() {
      try {
        const items = await sp.web.lists.getByTitle(SETTINGS_LIST).items.filter(`Title eq 'FinanceEmail'`).top(1)();
        if (items.length > 0) {
          const rawData = items[0].Value || items[0].Data;
          // Allow both , and ; as separators
          const emails = typeof rawData === "string" ?
            rawData.split(/[;,]/).map(e => e.trim()).filter(e => e) : [];
          const resolved = await resolvePeopleByEmails(sp, emails);
          setFinancePeople(resolved);
        }
      } catch {
        setFinancePeople([]);
      }
    }
    loadFinanceEmails();
  }, [sp]);

  async function resolvePeopleByEmails(sp: SPFI, emails: string[]): Promise<IPersonaProps[]> {
    const resolvedPeople: IPersonaProps[] = [];
    for (const email of emails) {
      if (!email) continue;
      // Search for exact match user profile by email
      const results = await sp.utility.searchPrincipals(
        email, PrincipalType.User, PrincipalSource.All, null, 1
      );
      const user = results?.[0];
      if (user) {
        resolvedPeople.push({
          text: user.DisplayName || user.LoginName || email,
          secondaryText: user.Email || email,
          id: (user.Email.toString() || user.PrincipalId?.toString() || user.LoginName || email)
        });
      } else {
        // Fallback: just use the email
        resolvedPeople.push({
          text: email,
          secondaryText: email,
          id: email
        });
      }
    }
    return resolvedPeople;
  }
  // Persist Finance Emails
  const saveFinanceEmails = async () => {
    setSavingConfig(true);
    setError(null);
    try {
      const emailsString = financePeople.map(p => p.secondaryText || '').filter(e => e).join(',');
      const items = await sp.web.lists.getByTitle(SETTINGS_LIST).items.filter(`Title eq 'FinanceEmail'`).top(1)();
      if (items.length === 0) {
        await sp.web.lists.getByTitle(SETTINGS_LIST).items.add({ Title: 'FinanceEmail', Value: emailsString });
      } else {
        await sp.web.lists.getByTitle(SETTINGS_LIST).items.getById(items[0].Id).update({ Value: emailsString });
      }
      setSavingConfig(false);
      alert('Finance Emails saved successfully');
    } catch (e) {
      setError('Error saving emails');
      setSavingConfig(false);
    }
  };

  // On toggle change: update UI and save immediately
  const onToggleChange = async (key: string, checked?: boolean) => {
    if (checked === undefined) return;
    const updated = { ...settings, [key]: checked };
    if (onConfigChange) {
      onConfigChange(updated);
    }
    setSettings(updated);
    if (onSettingsUpdate) onSettingsUpdate(updated);
    try {
      setSavingConfig(true);
      const items = await sp.web.lists.getByTitle(SETTINGS_LIST).items.filter(`Title eq '${PAGE_CONFIG_ITEM_TITLE}'`).top(1)();
      const configValue = JSON.stringify(updated);
      if (items.length === 0) {
        await sp.web.lists.getByTitle(SETTINGS_LIST).items.add({ Title: PAGE_CONFIG_ITEM_TITLE, Value: configValue });
      } else {
        await sp.web.lists.getByTitle(SETTINGS_LIST).items.getById(items[0].Id).update({ Value: configValue });
      }
      setSavingConfig(false);
    } catch (e) {
      setSavingConfig(false);
      setError("Error saving settings");
    }
  };

  if (loading || !settings) return <Spinner label="Loading settings..." />;

  return (
    <div style={{ maxWidth: 480, padding: 12 }}>
      <Text variant="large" styles={{ root: { fontWeight: 600, fontSize: 20, marginBottom: 22 } }}>Settings</Text>
      <Pivot>
        <PivotItem headerText="General" itemKey="general">
          <Stack tokens={{ childrenGap: 22 }} styles={{ root: { marginTop: 14, marginBottom: 12 } }}>
            <NormalPeoplePicker
              onResolveSuggestions={async (filterText, selectedItems) => {
                if (!filterText) return [];
                const results = await sp.utility.searchPrincipals(
                  filterText,
                  PrincipalType.User,
                  PrincipalSource.All,
                  null,
                  5
                );
                return results
                  .map((user: any) => ({
                    text: user.DisplayName || user.LoginName,
                    secondaryText: user.Email || "",
                    id: user.Id?.toString() || user.PrincipalId?.toString() || user.LoginName
                  }))
                  .filter(p => !selectedItems?.some(si => si.id === p.id));
              }}
              onChange={items => setFinancePeople(items || [])}
              selectedItems={financePeople}
              pickerSuggestionsProps={suggestionProps}
              resolveDelay={300}
              inputProps={{ placeholder: 'Select Finance person(s)' }}
            />
            <PrimaryButton
              text="Save"
              onClick={saveFinanceEmails}
              disabled={savingConfig}
              styles={{ root: { marginTop: 18, width: "100%", maxWidth: 340, fontWeight: 600, fontSize: 16 } }}
            />
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Page" itemKey="page">
          <Stack tokens={{ childrenGap: 16, padding: 8 }}>
            {pageSettings.map(ps => (
              <Stack
                key={ps.stateVariable}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 12 }}
                styles={{ root: { justifyContent: "space-between" } }}
              >
                <Text title={ps.tooltip} style={{ minWidth: 170, fontWeight: 500 }}>
                  {ps.label.replace(/^Hide\s?/, '')}
                </Text>
                <Toggle
                  checked={settings[ps.stateVariable]}
                  onChange={(_e, checked) => onToggleChange(ps.stateVariable, checked)}
                  onText="Hide"
                  offText="Show"
                />
              </Stack>
            ))}
          </Stack>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default Settings;
