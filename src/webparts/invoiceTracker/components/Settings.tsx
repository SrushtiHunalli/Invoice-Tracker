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
  { label: 'Command Bar', tooltip: 'Show or hide the command bar.', stateVariable: 'hideCommandBar', sharepointElement: '.CommandBarWrapper' },
  { label: 'Sidebar', tooltip: 'Show or hide the sidebar.', stateVariable: 'hideSidebar', sharepointElement: '.sidebar' },
  { label: 'Footer', tooltip: 'Show or hide the footer.', stateVariable: 'hideFooter', sharepointElement: '.footer' }
];

interface SettingsProps {
  sp: SPFI;
  onSettingsUpdate?: (settings: Record<string, boolean>) => void;
  context?: any;
}

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Suggested People',
  noResultsFoundText: 'No results found',
  mostRecentlyUsedHeaderText: 'Recent',
};

const Settings: React.FC<SettingsProps> = ({ sp, onSettingsUpdate }) => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [settings, setSettings] = useState<Record<string, boolean>>({
    hideCommandBar: false,
    hideSidebar: false,
    hideFooter: false,
  });
  const [financePeople, setFinancePeople] = useState<IPersonaProps[]>([]);

  // Load Finance Emails from InvoiceConfiguration list
  // useEffect(() => {
  //   async function loadFinanceEmails() {
  //     try {
  //       const items = await sp.web.lists
  //         .getByTitle('InvoiceConfiguration')
  //         .items
  //         .filter(`Title eq 'Finance Email'`)
  //         .top(1)();

  //       if (items.length > 0) {
  //         // If Data column is Person or Group, fetch user emails from those users
  //         // But for simplicity, assume Data has stored emails separated by semicolon as fallback
  //         const rawData = items[0].Data;
  //         if (typeof rawData === "string") {
  //           const emails = rawData.split(";").map((e: any) => e.trim()).filter((e: any) => e);
  //           const personas = emails.map((email: string) => ({ text: email, secondaryText: email }));
  //           setFinancePeople(personas);
  //         } else if (Array.isArray(rawData)) {
  //           // If Data is actually array of users (complex type), map accordingly
  //           const personas = rawData.map((user: any) => ({ text: user.Title || user.Email || user.LoginName, secondaryText: user.Email || user.LoginName }));
  //           setFinancePeople(personas);
  //         }
  //       }
  //       setLoading(false);
  //     } catch {
  //       setError('Error loading Finance Emails');
  //       setLoading(false);
  //     }
  //   }
  //   loadFinanceEmails();
  // }, [sp]);
  useEffect(() => {
    async function loadFinanceEmails() {
      try {
        const items = await sp.web.lists
          .getByTitle('InvoiceConfiguration')
          .items
          .filter(`Title eq 'Finance Email'`)
          .top(1)();

        if (items.length > 0) {
          const rawData = items[0].Data; // semicolon-separated emails string
          if (typeof rawData === "string") {
            const emails = rawData.split(";").map(e => e.trim()).filter(e => e);
            const personas = emails.map(email => ({ text: email, secondaryText: email }));
            setFinancePeople(personas);
          } else {
            // fallback empty or other type cases
            setFinancePeople([{ text: "srushti.hunalli@sacha.solutions", secondaryText: "srushti.hunalli@sacha.solutions" }]);
          }
        }
        setLoading(false);
      } catch {
        setError('Error loading Finance Emails');
        setLoading(false);
      }
    }
    loadFinanceEmails();
  }, [sp]);


  const saveFinanceEmails = async () => {
    setLoading(true);
    setError(null);

    try {
      const emailsString = financePeople.map(p => p.secondaryText || '').filter(e => e).join(';');

      // Check if 'Finance Email' item exists
      const items = await sp.web.lists
        .getByTitle('InvoiceConfiguration')
        .items
        .filter(`Title eq 'Finance Email'`)
        .top(1)();

      if (items.length === 0) {
        // Create new item with Data column as string of emails
        await sp.web.lists.getByTitle('InvoiceConfiguration').items.add({
          Title: 'Finance Email',
          Value: emailsString
        });
      } else {
        // Update existing item
        await sp.web.lists.getByTitle('InvoiceConfiguration').items.getById(items[0].Id).update({
          Value: emailsString
        });
      }

      setLoading(false);
      alert('Finance Emails saved successfully');

    } catch (e) {
      setError('Error saving emails');
      setLoading(false);
      console.error(e);
    }
  };


  const toggleElement = (selector: string, hide: boolean) => {
    const element = document.querySelector(selector);
    if (element && element instanceof HTMLElement) {
      element.style.display = hide ? 'none' : '';
    }
  };

  const onToggleChange = (key: string, checked?: boolean) => {
    if (checked === undefined) return;
    const updated = { ...settings, [key]: checked };
    setSettings(updated);
    if (onSettingsUpdate) onSettingsUpdate(updated);
    const setting = pageSettings.find(s => s.stateVariable === key);
    if (setting) toggleElement(setting.sharepointElement, checked!);
  };

  if (loading) return <Spinner label="Loading settings..." />;

  return (
    <div style={{ maxWidth: 450, padding: 10 }}>
      <Text variant="large" styles={{ root: { fontWeight: 600, fontSize: 20, marginBottom: 22 } }}>
        Settings
      </Text>
      <Pivot>
        <PivotItem headerText="General" itemKey="general">
          <Stack tokens={{ childrenGap: 22 }} styles={{ root: { marginTop: 14, marginBottom: 12 } }}>
            {/* <NormalPeoplePicker
              onResolveSuggestions={async (filterText: string, selectedItems?: IPersonaProps[]) => {
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
                    text: user.DisplayName || user.DisplayText || user.Title,
                    secondaryText: user.Email || user.EMail || '',
                    id: user.Key,
                  }))
                  .filter(p => !selectedItems?.some(si => si.id === p.id));
              }}
              onChange={(items?: IPersonaProps[]) => setFinancePeople(items || [])}
              selectedItems={financePeople}
              pickerSuggestionsProps={suggestionProps}
              resolveDelay={300}
              inputProps={{ placeholder: 'Select Finance person(s)' }}
            /> */}

            <NormalPeoplePicker
              onResolveSuggestions={async (filterText: string, selectedItems?: IPersonaProps[]) => {
                if (!filterText) return [];
                const results = await sp.utility.searchPrincipals(
                  filterText,
                  PrincipalType.User,
                  PrincipalSource.All,
                  null,
                  5
                );
                return results.map((user: any) => ({
                  text: user.DisplayName || user.LoginName,
                  secondaryText: user.Email || "",
                  id: user.Id?.toString() || user.PrincipalId?.toString() || user.LoginName
                }))
                  .filter(p => !selectedItems?.some(si => si.id === p.id));
              }}
              onChange={(items?: IPersonaProps[]) => setFinancePeople(items || [])}
              selectedItems={financePeople}
              pickerSuggestionsProps={suggestionProps}
              resolveDelay={300}
              inputProps={{ placeholder: 'Select Finance person(s)' }}
            />


            <PrimaryButton
              text="Save"
              onClick={saveFinanceEmails}
              disabled={loading}
              styles={{ root: { marginTop: 18, width: "100%", maxWidth: 340, fontWeight: 600, fontSize: 16 } }}
            />

            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
          </Stack>
        </PivotItem>

        <PivotItem headerText="Page" itemKey="page">
          <Stack tokens={{ childrenGap: 16, padding: 8 }}>
            {pageSettings.map(ps => (
              <Toggle
                key={ps.stateVariable}
                label={ps.label}
                checked={settings[ps.stateVariable]}
                onChange={(_e, checked) => onToggleChange(ps.stateVariable, checked)}
                onText="Hide"
                offText="Show"
              />
            ))}
          </Stack>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default Settings;
