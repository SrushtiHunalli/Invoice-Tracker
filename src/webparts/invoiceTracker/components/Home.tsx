import * as React from "react"; 
import { useEffect, useState } from "react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { SPFI } from "@pnp/sp";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";

interface HomeProps {
  sp: SPFI;
  context: any;
  onNavigate: (pageKey: string, filter?: any) => void;
}

function DashboardCard({
  label,
  count,
  ariaLabel,
  onClick,
}: {
  label: string;
  count: number;
  ariaLabel: string;
  onClick: () => void;
}) {
  return (
    <div
      role="button"
      tabIndex={0}
      aria-label={ariaLabel}
      onClick={onClick}
      onKeyDown={ev => {
        if (ev.key === "Enter" || ev.key === " ") onClick();
      }}
      style={{
        cursor: "pointer",
        userSelect: "none",
        background: "#f3f3f3",
        borderRadius: 12,
        width: 295,
        height: 155,
        border: "1.5px solid #ccc",
        boxShadow: "0 4px 10px rgba(0,0,0,0.1)",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "space-between",
        padding: 20,
      }}
    >
      <div style={{ fontSize: "1.25rem", fontWeight: 500, color: "#605e5c" }}>
        {label}
      </div>
      <div style={{ fontSize: "2.2rem", fontWeight: 700, color: "#323130" }}>
        {count}
      </div>
      <div
        style={{
          width: "100%",
          borderTop: "1.5px solid #ccc",
          color: "#605e5c",
          fontSize: "1.1rem",
          fontWeight: 550,
          padding: "15px 0 10px",
          borderRadius: "0 0 12px 12px",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        View
        <span style={{ marginLeft: 14, fontSize: "1.2rem", fontWeight: 600 }}>
          →
        </span>
      </div>
    </div>
  );
}

export default function Home({ sp, context, onNavigate }: HomeProps) {
  const [, setAllRequests] = useState<any[]>([]);
  const [counts, setCounts] = useState<{
    pending: number;
    paymentPending: number;
    clarification: number;
    paymentReceived: number;
  }>({
    pending: 0,
    paymentPending: 0,
    clarification: 0,
    paymentReceived: 0,
  });

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [userGroups, setUserGroups] = useState<string[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);

  // Fetch user groups once
  useEffect(() => {
    async function fetchUserGroups() {
      try {
        const spCurrent = spfi().using(SPFx(context));
        const groups = await spCurrent.web.currentUser.groups();
        setUserGroups(groups.map(g => g.Title));
      } catch {
        setUserGroups([]);
      } finally {
        setLoadingGroups(false);
      }
    }
    fetchUserGroups();
  }, [context, sp]);

  // Fetch all requests once and calculate counts
  useEffect(() => {
    async function fetchAllRequests() {
      setLoading(true);
      setError(null);
      try {
        const spCurrent = spfi().using(SPFx(context));
        const requests = await spCurrent.web.lists
          .getByTitle("Invoice Requests")
          .items
          .select("Id", "Status", "FinanceStatus")
          .top(5000)();

        setAllRequests(requests);

        // Compute counts locally
        const pending = requests.filter(r => r.FinanceStatus === "Pending").length;
        const paymentPending = requests.filter(r => r.Status === "Pending Payment").length;
        const clarification = requests.filter(r => r.FinanceStatus === "Clarification").length;
        const paymentReceived = requests.filter(r => r.Status === "Payment Received").length;

        setCounts({
          pending,
          paymentPending,
          clarification,
          paymentReceived,
        });
      } catch {
        setError("Failed to load invoice requests");
        setCounts({
          pending: 0,
          paymentPending: 0,
          clarification: 0,
          paymentReceived: 0,
        });
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

  // Roles from groups
  const roles = userGroups.map(g => g.toLowerCase());
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
    const filterMap: { [key: string]: any } = {
      pending: { FinanceStatus: "Pending" },
      paymentPending: { Status: "Pending Payment" },
      clarification: { FinanceStatus: "Clarification" },
      paymentReceived: { Status: "Payment Received" },
    };

    if (showAll || showPmOnly) {
      onNavigate("myrequests", { initialFilters: filterMap[filterKey] });
    } else if (showFinanceOnly) {
      onNavigate("financeview", { initialFilters: filterMap[filterKey] });
    } else {
      alert("Access denied");
    }
  }

  return (
    <div
      style={{
        display: "flex",
        flexWrap: "wrap",
        justifyContent: "center",
        gap: 54,
        padding: 40,
        background: "#f3f3f3",
        minHeight: "80vh",
      }}
    >
      {(showAll || showPmOnly || showFinanceOnly) && (
        <>
          <DashboardCard
            label="Pending Requests"
            count={counts.pending}
            ariaLabel={`Pending Requests: ${counts.pending}`}
            onClick={() => onCardClick("pending")}
          />
          <DashboardCard
            label="Pending Payment"
            count={counts.paymentPending}
            ariaLabel={`Pending Payment: ${counts.paymentPending}`}
            onClick={() => onCardClick("paymentPending")}
          />
          <DashboardCard
            label="Clarification"
            count={counts.clarification}
            ariaLabel={`Clarification Requests: ${counts.clarification}`}
            onClick={() => onCardClick("clarification")}
          />
          <DashboardCard
            label="Payment Received"
            count={counts.paymentReceived}
            ariaLabel={`Payment Received: ${counts.paymentReceived}`}
            onClick={() => onCardClick("paymentReceived")}
          />
        </>
      )}

      {(showAll || showPmOnly) && (
        <div
          role="button"
          tabIndex={0}
          onClick={() => onNavigate("Createview")}
          onKeyDown={e => {
            if (e.key === "Enter" || e.key === " ") onNavigate("Createview");
          }}
          style={{
            cursor: "pointer",
            userSelect: "none",
            background: "#eaeaea",
            borderRadius: 12,
            width: 430,
            height: 70,
            border: "1.5px solid #ccc",
            boxShadow: "0 4px 10px rgba(0,0,0,0.1)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: 20,
            fontWeight: 600,
            color: "#323130",
            marginTop: "2rem",
          }}
        >
          Create Invoice Request
          <span style={{ marginLeft: 16, fontSize: 22, fontWeight: 700 }}>→</span>
        </div>
      )}
    </div>
  );
}
