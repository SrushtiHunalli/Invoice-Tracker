import * as React from "react";
import { useEffect, useState } from "react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { SPFI } from "@pnp/sp";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { Icon } from "@fluentui/react";

interface HomeProps {
  sp: SPFI;
  context: any;
  onNavigate: (pageKey: string, filter?: any) => void;
  primaryColor?: string; // pass or derive from theme for webpart color
}

function DashboardCard({
  label,
  count,
  ariaLabel,
  onClick,
  iconName,
  accentColor,
  cardWidth,
  cardHeight,
  cardPadding
}: {
  label: string;
  count: number;
  ariaLabel: string;
  onClick: () => void;
  iconName: string;
  accentColor: string;
  cardWidth: number | string;
  cardHeight: number | string;
  cardPadding: number | string;
}) {
  return (
    <div
      role="button"
      tabIndex={0}
      aria-label={ariaLabel}
      onClick={onClick}
      onKeyDown={(ev) => {
        if (ev.key === "Enter" || ev.key === " ") onClick();
      }}
      style={{
        minWidth: cardWidth,
        height: cardHeight,
        background: "#fff",
        borderRadius: 14,
        boxShadow: "0 2px 15px rgba(0,30,60,0.06)",
        padding: cardPadding,
        margin: "auto",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        transition: "box-shadow .15s, transform .14s",
        outline: 0,
        border: "1.5px solid #eee",
        cursor: "pointer",
      }}
      onMouseEnter={(e) => {
        (e.currentTarget as HTMLElement).style.boxShadow = `0 6px 32px ${accentColor}22`;
        (e.currentTarget as HTMLElement).style.transform = "translateY(-2px)";
      }}
      onMouseLeave={(e) => {
        (e.currentTarget as HTMLElement).style.boxShadow = "0 2px 15px rgba(0,30,60,0.06)";
        (e.currentTarget as HTMLElement).style.transform = "none";
      }}
    >
      <Icon iconName={iconName} styles={{ root: { fontSize: 36, color: accentColor, marginBottom: 8 } }} />
      <div style={{ fontSize: 19, fontWeight: 600, color: "#222", marginBottom: 2, textAlign: "center" }}>
        {label}
      </div>
      <div style={{ fontSize: 32, fontWeight: 700, color: "#218838", margin: "6px 0 4px" }}>
        {count}
      </div>
      <div
        style={{
          fontSize: 15,
          color: "#444",
          fontWeight: 500,
          borderTop: "1px solid #e6e6e6",
          paddingTop: 9,
          marginTop: "auto",
          width: "90%",
          textAlign: "center"
        }}
      >
        View <span style={{ marginLeft: 10, fontSize: 18, fontWeight: 700 }}>→</span>
      </div>
    </div>
  );
}

export default function Home({ sp, context, onNavigate, primaryColor }: HomeProps) {
  // const theme = getTheme();
  const spTheme = (window as any).__themeState__?.theme;
  const accentColor = spTheme?.themePrimary || "#0078d4";
  // const neutralColor = spTheme?.neutralPrimary || "#444";

  // const accentColor = primaryColor || theme.palette.themePrimary || "#2564cf";

  const [, setAllRequests] = useState<any[]>([]);
  const [counts, setCounts] = useState({ pending: 0, paymentPending: 0, clarification: 0, paymentReceived: 0 });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [userGroups, setUserGroups] = useState<string[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);

  // Responsive card size states
  const [cardWidth, setCardWidth] = useState<number | string>(220);
  const [cardHeight, setCardHeight] = useState<number | string>(152);
  const [cardPadding, setCardPadding] = useState<number | string>(24);

  // Adjust card size based on window width for responsiveness
  useEffect(() => {
    function handleResize() {
      const width = window.innerWidth;
      if (width <= 500) {
        setCardWidth(120);
        setCardHeight(92);
        setCardPadding(8);
      } else if (width <= 750) {
        setCardWidth(150);
        setCardHeight(120);
        setCardPadding(14);
      } else {
        setCardWidth(220);
        setCardHeight(152);
        setCardPadding(24);
      }
    }

    handleResize(); // initial
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    async function fetchUserGroups() {
      try {
        const spCurrent = spfi().using(SPFx(context));
        const groups = await spCurrent.web.currentUser.groups();
        setUserGroups(groups.map((g) => g.Title));
      } catch {
        setUserGroups([]);
      } finally {
        setLoadingGroups(false);
      }
    }
    fetchUserGroups();
  }, [context, sp]);

  useEffect(() => {
    async function fetchAllRequests() {
      setLoading(true);
      setError(null);
      try {
        const spCurrent = spfi().using(SPFx(context));
        const requests = await spCurrent.web.lists
          .getByTitle("Invoice Requests")
          .items.select("Id", "Status", "CurrentStatus", "FinanceStatus")
          .top(5000)();

        setAllRequests(requests);

        setCounts({
          pending: requests.filter(r => r.FinanceStatus === "Pending").length,
          paymentPending: requests.filter((r) => r.Status === "Pending Payment").length,
          clarification: requests.filter(
            r => r.CurrentStatus === "Finance asked Clarification"
          ).length,
          paymentReceived: requests.filter((r) => r.Status === "Payment Received").length,
        });
      } catch {
        setError("Failed to load invoice requests");
        setCounts({ pending: 0, paymentPending: 0, clarification: 0, paymentReceived: 0 });
      } finally {
        setLoading(false);
      }
    }
    fetchAllRequests();
  }, [context, sp]);

  if (loading || loadingGroups)
    return (
      <div style={{ padding: 20 }}>
        <Spinner label="Loading ..." />
      </div>
    );

  if (error)
    return (
      <div style={{ padding: 20, color: "red" }}>
        <strong>Error: {error}</strong>
      </div>
    );

  const roles = userGroups.map((g) => g.toLowerCase());
  const isAdmin = roles.includes("admin");
  const isFinance = roles.includes("finance");
  const isPm = roles.includes("pm");
  const isDm = roles.includes("dm");
  const isDh = roles.includes("dh");
  const isPmDmDh = isPm || isDm || isDh;
  const showAll = isAdmin || (isPmDmDh && isFinance);
  const showPmOnly = !isFinance && isPmDmDh;
  const showFinanceOnly = isFinance && !isPmDmDh;

  function onCardClick(filterKey: keyof typeof counts) {
    const filterMap = {
      pending: { FinanceStatus: "Pending" },
      paymentPending: { Status: "Pending Payment" },
      clarification: { CurrentStatus: "Finance asked Clarification" },
      paymentReceived: { Status: "Payment Received" },
    };
    if (showAll || showPmOnly) onNavigate("myrequests", { initialFilters: filterMap[filterKey] });
    else if (showFinanceOnly) onNavigate("financeview", { initialFilters: filterMap[filterKey] });
    else alert("Access denied");
  }

  return (
    <div
      style={{
        width: "100%",
        minHeight: "50%",
        background: "linear-gradient(120deg,#f6f8fa 80%,#eaf2fc 100%)",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
      }}
    >
      {/* Responsive max-width grid, always centered */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))",
          gap: 36,
          justifyContent: "center",
          width: "100%",
          maxWidth: 1110,
          margin: "56px 0 0",
        }}
      >
        {(showAll || showPmOnly || showFinanceOnly) && (
          <>
            <DashboardCard
              label="Pending Requests"
              count={counts.pending}
              ariaLabel={`Pending Requests: ${counts.pending}`}
              onClick={() => onCardClick("pending")}
              iconName="Edit"
              accentColor={accentColor}
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
            <DashboardCard
              label="Pending Payment"
              count={counts.paymentPending}
              ariaLabel={`Pending Payment: ${counts.paymentPending}`}
              onClick={() => onCardClick("paymentPending")}
              iconName="Clock"
              accentColor={accentColor}
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
            <DashboardCard
              label="Clarification"
              count={counts.clarification}
              ariaLabel={`Clarification Requests: ${counts.clarification}`}
              onClick={() => onCardClick("clarification")}
              iconName="WarningSolid"
              accentColor={accentColor}
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
            <DashboardCard
              label="Payment Received"
              count={counts.paymentReceived}
              ariaLabel={`Payment Received: ${counts.paymentReceived}`}
              onClick={() => onCardClick("paymentReceived")}
              iconName="CheckMark"
              accentColor={accentColor}
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
          </>
        )}
      </div>
      {(showAll || showPmOnly) && (
        <div
          role="button"
          tabIndex={0}
          onClick={() => onNavigate("Createview")}
          onKeyDown={(e) => {
            if (e.key === "Enter" || e.key === " ") onNavigate("Createview");
          }}
          style={{
            margin: "54px auto 0",
            cursor: "pointer",
            background: accentColor,
            borderRadius: 32,
            padding: "18px 42px",
            display: "flex",
            alignItems: "center",
            color: "#fff",
            fontSize: 21,
            boxShadow: `0 8px 32px ${accentColor}40`,
            fontWeight: 700,
            border: 0,
            outline: 0,
            transition: "background .17s,box-shadow .14s",
          }}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLElement).style.background = "#174fb6";
            (e.currentTarget as HTMLElement).style.boxShadow = `0 10px 34px ${accentColor}33`;
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLElement).style.background = accentColor;
            (e.currentTarget as HTMLElement).style.boxShadow = `0 8px 32px ${accentColor}40`;
          }}
        >
          <span style={{ marginRight: 12 }}>Create Invoice Request</span>
          <Icon iconName="Forward" styles={{ root: { fontSize: 25, fontWeight: 700 } }} />
        </div>
      )}
    </div>
  );
}
