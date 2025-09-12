// // InvoiceTrackerWebPart.ts

import * as React from "react";
import * as ReactDom from "react-dom";
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from "@microsoft/sp-webpart-base";
import { IInvoiceTrackerProps } from "./components/IInvoiceTrackerProps";
import InvoiceTracker from "./components/InvoiceTracker";

// Import SitePicker control
import { PropertyFieldSitePicker } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface IInvoiceTrackerWebPartProps {
  description: string;
  selectedSites: any;
  projectsSite: string;
}

export default class InvoiceTrackerWebPart extends BaseClientSideWebPart<IInvoiceTrackerWebPartProps> {

  public render(): void {
    const projectSiteUrl = Array.isArray(this.properties.selectedSites) && this.properties.selectedSites.length > 0
      ? (typeof this.properties.selectedSites[0] === "object" && "url" in this.properties.selectedSites[0]
        ? (this.properties.selectedSites[0] as any).url
        : this.properties.selectedSites[0])
      : undefined;

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
        projectSiteUrl
      }
    )
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // Configure property pane
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  // Map selectedSites property to initialSites format expected by SitePicker
  const initialSites = (Array.isArray(this.properties.selectedSites) && this.properties.selectedSites.length > 0)
    ? this.properties.selectedSites.map(site => typeof site === "string" ? { url: site } : site)
    : [{ url: this.context.pageContext.web.absoluteUrl }]; // fallback to current site if none selected

  return {
    pages: [
      {
        // header: { description: "Invoice Tracker Settings" },
        groups: [
          {
            groupName: "General",
            groupFields: [
              PropertyFieldSitePicker("selectedSites", {
                label: "Select Project Site",
                initialSites: initialSites,       // <-- Use saved selections here
                context: this.context as any,
                multiSelect: false,
                deferredValidationTime: 0,
                properties: this.properties,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                key: "sitePickerFieldId"
              })  
            ]
          }
        ]
      }
    ]
  };
}

}
