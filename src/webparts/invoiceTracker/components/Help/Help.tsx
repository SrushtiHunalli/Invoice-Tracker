import { useContext } from 'react';
import * as React from "react";
import { Panel, PanelType } from 'office-ui-fabric-react';
import { InvoiceTrackerContext } from '../InvoiceTracker'; // create/use your context similar to TaskContext

const links = [
  { title: "About Us", url: "https://www.invoice-tracker.example.com/about" },
  { title: "Privacy Policy", url: "https://www.invoice-tracker.example.com/privacy" },
  { title: "Support", url: "https://www.invoice-tracker.example.com/support" },
  { title: "Usage guide", url: "https://www.invoice-tracker.example.com/usage-guide" },
];

const EDITION: string = "Standard";

const Help: React.FC<{ isOpen: boolean; onDismiss(): void; }> = ({ isOpen, onDismiss }) => {
  const context: any = useContext(InvoiceTrackerContext);
  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      headerText='Help'
      type={PanelType.custom}
      customWidth="350px"
    >
      {/* logo */}
      <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", marginBottom: "1.15em" }}>
        <div style={{ marginLeft: "-0.5em" }} >
          <img src={require('../../assets/ChandraWorks.png')} alt='Invoice Tracker' height={75} />
        </div>
        <div style={{ margin: "-1em 0 0 -0.5em", display: "flex", gap: "1em", alignItems: "center" }} >
          <img src={require('../../assets/Logo.png')} alt='Invoice Tracker Icon' height={45} />
          <p style={{ margin: "0 0 0 -0.5em", fontWeight: 600, userSelect: "none", fontSize: '1.25em', color: context.isDarkTheme ? "#fff" : "#000" }}>
            Invoice Tracker
          </p>
        </div>
      </div>

      <hr />

      {/* version and type */}
      <p style={{ fontWeight: 500, color: context.isDarkTheme ? "#fff" : "#000" }}>
        App Type: {`${EDITION}`.charAt(0).toUpperCase() + `${EDITION}`.slice(1)}
      </p>
      <p style={{ fontWeight: 500, color: context.isDarkTheme ? "#fff" : "#000" }}>Version: 1.0.0</p>

      {/* links */}
      {
        links.map((l, idx) => (
          <p key={idx}>
            <a
              style={{ color: context.isDarkTheme ? "#fff" : "#000", fontWeight: 500 }}
              href={l.url}
              target='_blank'
              rel='noopener noreferrer'
            >
              {l.title}
            </a>
          </p>
        ))
      }
    </Panel>
  );
};

export default Help;
