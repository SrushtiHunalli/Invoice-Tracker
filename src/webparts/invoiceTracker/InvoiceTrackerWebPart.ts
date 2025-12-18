// // InvoiceTrackerWebPart.ts

import * as React from "react";
import * as ReactDom from "react-dom";
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from "@microsoft/sp-webpart-base";
import { IInvoiceTrackerProps } from "./components/IInvoiceTrackerProps";
import InvoiceTracker from "./components/InvoiceTracker";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users";

import {
  PropertyFieldPeoplePicker,
  PrincipalType
} from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

// Import SitePicker control
import { PropertyFieldSitePicker } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface IInvoiceTrackerWebPartProps {
  description: string;
  selectedSites: any;
  projectsSite: string;
  pageConfig?: any;
  viewAsUser?: string;
}

export default class InvoiceTrackerWebPart extends BaseClientSideWebPart<IInvoiceTrackerWebPartProps> {
  private _isOwner: boolean = false;
  public async onInit(): Promise<void> {
    await super.onInit();

    const sp = spfi().using(SPFx(this.context));

    try {
      const ownerGroup = await sp.web.associatedOwnerGroup();
      const ownerUsers = await sp.web.siteGroups.getById(ownerGroup.Id).users();
      const me = await sp.web.currentUser();
      this._isOwner = ownerUsers.some(u => u.Id === me.Id);
    } catch (e) {
      this._isOwner = false;
    }

    if (this.context.propertyPane) {
      this.context.propertyPane.refresh();
    }

    function decodeHtmlEntities(str: string): string {
      const txt = document.createElement('textarea');
      txt.innerHTML = str;
      return txt.value;
    }
    try {
      const items = await sp.web.lists.getByTitle("InvoiceConfiguration").items
        .filter(`Title eq 'PageConfig'`)
        .top(1)();

      if (items.length > 0 && items[0].Value) {
        const decodedValue = decodeHtmlEntities(items[0].Value);
        this.properties.pageConfig = JSON.parse(decodedValue);
      } else {
        this.properties.pageConfig = {};
      }
    } catch (e) {
      console.error("Error loading pageConfig:", e);
      this.properties.pageConfig = {};
    }
  }
  public render(): void {
    const projectSiteUrl = Array.isArray(this.properties.selectedSites) && this.properties.selectedSites.length > 0
      ? (typeof this.properties.selectedSites[0] === "object" && "url" in this.properties.selectedSites[0]
        ? (this.properties.selectedSites[0] as any).url
        : this.properties.selectedSites[0])
      : undefined;

    const effectiveLogin =
      this.properties.viewAsUser || this.context.pageContext.user.loginName;

    const element: React.ReactElement<IInvoiceTrackerProps> = React.createElement(
      InvoiceTracker,
      {
        context: this.context,
        isDarkTheme: false,
        environmentMessage: "",
        hasTeamsContext: false,
        userDisplayName: this.context.pageContext.user.displayName,
        description: this.properties.description,
        selectedSites: this.properties.selectedSites,
        projectSiteUrl,
        pageConfig: this.properties.pageConfig,
        effectiveUserLogin: effectiveLogin
      }
    )
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === "viewAsUser") {
      const picked = Array.isArray(newValue) && newValue.length ? newValue[0] : null;
      this.properties.viewAsUser = picked ? picked.login : undefined;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  // Configure property pane
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Map selectedSites property to initialSites format expected by SitePicker
    const initialSites = (Array.isArray(this.properties.selectedSites) && this.properties.selectedSites.length > 0)
      ? this.properties.selectedSites.map(site => typeof site === "string" ? { url: site } : site)
      : [{ url: this.context.pageContext.web.absoluteUrl }]; // fallback to current site if none selected

    const fields: any[] = [
      PropertyFieldSitePicker("selectedSites", {
        label: "Select Project Site",
        initialSites,
        context: this.context as any,
        multiSelect: false,
        deferredValidationTime: 0,
        properties: this.properties,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        key: "sitePickerFieldId"
      })
    ];
    console.log("isOwner?", this._isOwner);

    if (this._isOwner) {
      fields.push(
        PropertyFieldPeoplePicker("viewAsUser", {
          label: "View as user",
          context: this.context as any,
          properties: this.properties,
          initialData: [],
          allowDuplicate: false,
          principalType: [PrincipalType.Users],
          multiSelect: false,
          onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
          key: "viewAsUserPicker"
        })
      );
    }
    return {
      pages: [
        {
          groups: [
            {
              groupName: "General",
              groupFields: fields
            }
          ]
        }
      ]
    };
  }

}
