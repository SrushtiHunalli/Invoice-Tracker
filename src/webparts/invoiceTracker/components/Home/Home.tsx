import * as React from "react";
import { useEffect, useState } from "react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { SPFI } from "@pnp/sp";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { Icon, Dialog, DialogType, DialogFooter, PrimaryButton } from "@fluentui/react";

interface HomeProps {
  sp: SPFI;
  context: any;
  onNavigate: (pageKey: string, filter?: any) => void;
  primaryColor?: string;
}

function DashboardCard({
  label,
  count,
  ariaLabel,
  onClick,
  iconName,
  cardWidth,
  cardHeight,
  cardPadding,
  color = primaryColor
}: {
  label: string;
  count: number;
  ariaLabel: string;
  onClick: () => void;
  iconName: string;
  cardWidth: number | string;
  cardHeight: number | string;
  cardPadding: number | string;
  color?: string;
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
        borderRadius: 16,
        boxShadow: "0 4px 20px rgba(0,0,0,0.08)",
        padding: cardPadding,
        margin: "auto",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "flex-start",
        transition: "all .2s cubic-bezier(0.4, 0, 0.2, 1)",
        outline: 0,
        border: "2px solid transparent",
        cursor: "pointer",
        position: "relative",
      }}
      onMouseEnter={(e) => {
        const target = e.currentTarget as HTMLElement;
        target.style.boxShadow = `0 8px 32px ${color}20`;
        target.style.transform = "translateY(-4px)";
        target.style.borderColor = color;
      }}
      onMouseLeave={(e) => {
        const target = e.currentTarget as HTMLElement;
        target.style.boxShadow = "0 4px 20px rgba(0,0,0,0.08)";
        target.style.transform = "translateY(0)";
        target.style.borderColor = "transparent";
      }}
    >
      <div style={{
        width: 40, height: 40,  // ✅ SMALLER ICON CONTAINER
        background: `${color}12`,
        borderRadius: 10,  // ✅ SMALLER RADIUS
        display: "flex", alignItems: "center", justifyContent: "center",
        marginBottom: 8  // ✅ TIGHTER SPACING
      }}>
        <Icon iconName={iconName} styles={{
          root: { fontSize: 20, color, fontWeight: "bold" }  // ✅ SMALLER ICON
        }} />
      </div>
      <div style={{
        fontSize: 12, fontWeight: 600,  // ✅ SMALLER TEXT
        color: "#555", marginBottom: 2,  // ✅ TIGHTER SPACING
        textAlign: "center",
        letterSpacing: 0.3
      }}>
        {label}
      </div>
      <div style={{
        fontSize: 24, fontWeight: 800,  // ✅ SLIGHTLY SMALLER NUMBER
        color, margin: "2px 0 8px",  // ✅ TIGHTER SPACING
        fontFeatureSettings: "'zero' 1"
      }}>
        {count}
      </div>
      <div style={{
        fontSize: 11,  // ✅ SMALLER FOOTER TEXT
        color: "#888",
        fontWeight: 500,
        borderTop: "1px solid #f0f0f0",
        paddingTop: 8,  // ✅ TIGHTER PADDING
        marginTop: "auto",
        width: "90%",
        textAlign: "center",
      }}>
        View <span style={{
          marginLeft: 4,  // ✅ TIGHTER SPACING
          fontSize: 14,  // ✅ SMALLER ARROW
          fontWeight: 700,
          color
        }}>→</span>
      </div>
    </div>
  );
}

const spTheme = (window as any).__themeState__?.theme;
const primaryColor = spTheme?.themePrimary || "#0078d4";

export default function Home({ sp, context, onNavigate }: HomeProps) {
  const spTheme = (window as any).__themeState__?.theme;
  const primaryColor = spTheme?.themePrimary || "#0078d4";

  const [, setAllRequests] = useState<any[]>([]);
  const [counts, setCounts] = useState({ pending: 0, paymentPending: 0, clarification: 0, paymentReceived: 0, overdue: 0, cancelled: 0 });
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [userGroups, setUserGroups] = useState<string[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);
  const [isAccessDeniedDialogVisible, setIsAccessDeniedDialogVisible] = React.useState(false);

  // ✅ COMPACT RESPONSIVE SIZES
  const [cardWidth, setCardWidth] = useState<number | string>(120);  // ✅ SMALLER
  const [cardHeight, setCardHeight] = useState<number | string>(110);  // ✅ SMALLER
  const [cardPadding, setCardPadding] = useState<number | string>(16);  // ✅ SMALLER

  // Adjust card size based on window width for responsiveness
  useEffect(() => {
    function handleResize() {
      const width = window.innerWidth;
      if (width <= 500) {
        setCardWidth(100);      // ✅ EXTRA SMALL MOBILE
        setCardHeight(92);      // ✅ EXTRA SMALL MOBILE
        setCardPadding(12);     // ✅ EXTRA SMALL MOBILE
      } else if (width <= 750) {
        setCardWidth(120);      // ✅ SMALL TABLET
        setCardHeight(100);     // ✅ SMALL TABLET
        setCardPadding(14);     // ✅ SMALL TABLET
      } else {
        setCardWidth(120);      // ✅ COMPACT DESKTOP
        setCardHeight(110);     // ✅ COMPACT DESKTOP
        setCardPadding(16);     // ✅ COMPACT DESKTOP
      }
    }

    handleResize();
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
          .items.select("Id", "Status", "CurrentStatus", "FinanceStatus", "DueDate")
          .top(5000)();

        setAllRequests(requests);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        setCounts({
          pending: requests.filter(r => r.Status === "Invoice Requested").length,
          paymentPending: requests.filter((r) => r.Status === "Pending Payment").length,
          clarification: requests.filter(
            r => r.CurrentStatus === "Finance asked Clarification"
          ).length,
          paymentReceived: requests.filter((r) => r.Status === "Payment Received").length,
          overdue: requests.filter(r =>
            r.Status === "Overdue"
          ).length,
          cancelled: requests.filter((r) => r.Status === "Cancelled").length,
        });
      } catch {
        setError("Failed to load invoice requests");
        setCounts({ pending: 0, paymentPending: 0, clarification: 0, paymentReceived: 0, overdue: 0, cancelled: 0 });
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
      pending: { Status: "Invoice Requested" },
      paymentPending: { Status: "Pending Payment" },
      clarification: { CurrentStatus: "Finance asked Clarification" },
      paymentReceived: { Status: "Payment Received" },
      overdue: { Status: "Overdue" },
      cancelled: { Status: "Cancelled" },
    };
    if (showAll || showPmOnly) onNavigate("myrequests", { initialFilters: filterMap[filterKey] });
    else if (showFinanceOnly) onNavigate("updaterequests", { initialFilters: filterMap[filterKey] });
    else setIsAccessDeniedDialogVisible(true);
  }

  return (
    <div
      style={{
        width: "100%",
        height: "100%",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
      }}
    >
      {/* ✅ TIGHTER GRID SPACING */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: window.innerWidth <= 500
            ? "repeat(auto-fit, minmax(140px, 1fr))"
            : `repeat(auto-fit, minmax(20px, 1fr))`,
          gap: window.innerWidth <= 500 ? 12 : 16,  // ✅ TIGHTER GAP
          justifyContent: "center",
          width: "100%",
          maxWidth: 1300,
          margin: "32px auto",  // ✅ LESS TOP MARGIN
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
              color="#0681e6ff"
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
              color="#ffb900"
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
              color="#ff8c00"
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
          </>
        )}
      </div>

      {/* ✅ TIGHTER SECOND GRID SPACING */}
      <div
        style={{
          display: "grid",
          gridTemplateColumns: window.innerWidth <= 500
            ? "repeat(auto-fit, minmax(140px, 1fr))"
            : `repeat(auto-fit, minmax(20px, 1fr))`,
          gap: window.innerWidth <= 500 ? 12 : 16,  // ✅ TIGHTER GAP
          justifyContent: "center",
          width: "100%",
          maxWidth: 1300,
          margin: "24px auto",  // ✅ LESS MARGIN
        }}
      >
        {(showAll || showPmOnly || showFinanceOnly) && (
          <>
            <DashboardCard
              label="Overdue"
              count={counts.overdue}
              ariaLabel={`Overdue Requests: ${counts.overdue}`}
              onClick={() => onCardClick("overdue")}
              iconName="ClockSolid"
              color="#808080"
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
            <DashboardCard
              label="Cancelled"
              count={counts.cancelled}
              ariaLabel={`Cancelled Requests: ${counts.cancelled}`}
              onClick={() => onCardClick("cancelled")}
              iconName="Cancel"
              color="#d13438"
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
              color="#107c10"
              cardWidth={cardWidth}
              cardHeight={cardHeight}
              cardPadding={cardPadding}
            />
          </>
        )}

        <Dialog
          hidden={!isAccessDeniedDialogVisible}
          onDismiss={() => setIsAccessDeniedDialogVisible(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Access Denied",
            subText: "You do not have permission to view this content.",
          }}
          modalProps={{
            isBlocking: false,
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={() => setIsAccessDeniedDialogVisible(false)} text="OK" />
          </DialogFooter>
        </Dialog>
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
            margin: "40px auto 0",  // ✅ LESS BOTTOM MARGIN
            cursor: "pointer",
            background: primaryColor,
            borderRadius: 32,
            padding: "18px 42px",
            display: "flex",
            alignItems: "center",
            color: "#fff",
            fontSize: 21,
            boxShadow: `0 8px 32px ${primaryColor}40`,
            fontWeight: 700,
            border: 0,
            outline: 0,
            transition: "background .17s,box-shadow .14s",
          }}
          onMouseEnter={(e) => {
            (e.currentTarget as HTMLElement).style.background = "#424242e9";
            (e.currentTarget as HTMLElement).style.boxShadow = `0 10px 34px ${primaryColor}33`;
          }}
          onMouseLeave={(e) => {
            (e.currentTarget as HTMLElement).style.background = primaryColor;
            (e.currentTarget as HTMLElement).style.boxShadow = `0 8px 32px ${primaryColor}40`;
          }}
        >
          <span style={{ marginRight: 12 }}>Create Invoice Request</span>
          <Icon iconName="Forward" styles={{ root: { fontSize: 25, fontWeight: 700 } }} />
        </div>
      )}
    </div>
  );
}
