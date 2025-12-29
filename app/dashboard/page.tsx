"use client";

export const dynamic = "force-dynamic";

import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  addDays,
  addYears,
  endOfWeek,
  format,
  isValid,
  parseISO,
  startOfWeek,
  startOfYear,
  subWeeks,
} from "date-fns";
import {
  ResponsiveContainer,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  BarChart,
  Bar,
  LabelList,
} from "recharts";
import { getSupabase } from "@/lib/supabaseClient";
import { Button, Card } from "@/components/ui";

type Row = Record<string, any>;
type SortKey =
  | "client"
  | "created_at"
  | "BOP_Date"
  | "BOP_Status"
  | "Followup_Date"
  | "status"
  | "CalledOn"
  | "Issued";
type SortDir = "asc" | "desc";

type ProgressSortKey =
  | "client_name"
  | "last_call_date"
  | "call_attempts"
  | "last_bop_date"
  | "bop_attempts"
  | "last_followup_date"
  | "followup_attempts";

const ALL_PAGE_SIZE = 20;
const PROGRESS_PAGE_SIZE = 20;

const READONLY_LIST_COLS = new Set([
  "interest_type",
  "business_opportunities",
  "wealth_solutions",
  "preferred_days",
]);

const DATE_TIME_KEYS = new Set([
  "BOP_Date",
  "CalledOn",
  "Followup_Date",
  "FollowUp_Date",
  "Issued",
]);

const LABEL_OVERRIDES: Record<string, string> = {
  client_name: "Client Name",
  last_call_date: "Last Call On",
  call_attempts: "No of Calls",
  last_bop_date: "Last BOP Call On",
  bop_attempts: "No of BOP Calls",
  last_followup_date: "Last FollowUp On",
  followup_attempts: "No of FollowUp Calls",
  created_at: "Created Date",
  interest_type: "Interest Type",
  business_opportunities: "Business Opportunities",
  wealth_solutions: "Wealth Solutions",
  preferred_days: "Preferred Days",
  preferred_time: "Preferred Time",
  referred_by: "Referred By",
  CalledOn: "Called On",
  BOP_Date: "BOP Date",
  BOP_Status: "BOP Status",
  Followup_Date: "Follow-Up Date",
  FollowUp_Status: "Follow-Up Status",
};

function labelFor(key: string) {
  if (LABEL_OVERRIDES[key]) return LABEL_OVERRIDES[key];
  const s = key
    .replace(/_/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .trim();
  const acronyms = new Set(["BOP", "ID", "API", "URL", "CAN"]);
  return s
    .split(/\s+/)
    .map((w) =>
      acronyms.has(w.toUpperCase())
        ? w.toUpperCase()
        : w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()
    )
    .join(" ");
}

function clientName(r: Row) {
  return `${r.first_name || ""} ${r.last_name || ""}`.trim();
}


function matchesGlobalSearch(r: any, needle: string) {
  if (!needle) return true;
  const n = needle.toLowerCase();
  const full = `${r?.first_name || ""} ${r?.last_name || ""}`.trim();
  const fields = [
    r?.client_name,
    full,
    r?.first_name,
    r?.last_name,
    r?.phone,
    r?.email,
  ];
  return fields.some((v) => String(v ?? "").toLowerCase().includes(n));
}

function toLocalInput(value: any) {
  if (!value) return "";
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return "";
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(
    d.getHours()
  )}:${pad(d.getMinutes())}`;
}

function fromLocalInput(value: string) {
  if (!value?.trim()) return null;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return null;
  return d.toISOString();
}

function asListItems(value: any): string[] {
  if (value == null) return [];
  if (Array.isArray(value)) return value.map((v) => String(v)).filter(Boolean);
  const s = String(value).trim();
  if (!s) return [];
  if (s.includes(",")) return s.split(",").map((x) => x.trim()).filter(Boolean);
  return [s];
}

function toggleSort(cur: { key: SortKey; dir: SortDir }, k: SortKey) {
  if (cur.key !== k) return { key: k, dir: "asc" as SortDir };
  return { key: k, dir: cur.dir === "asc" ? ("desc" as SortDir) : ("asc" as SortDir) };
}

function toggleProgressSort(
  cur: { key: ProgressSortKey; dir: SortDir },
  k: ProgressSortKey
) {
  if (cur.key !== k) return { key: k, dir: "asc" as SortDir };
  return { key: k, dir: cur.dir === "asc" ? ("desc" as SortDir) : ("asc" as SortDir) };
}

/** -------- Column Resize Helper (used by all tables) -------- */
function useColumnResizer() {
  const [widths, setWidths] = useState<Record<string, number>>({});
  const resizeRef = useRef<{
    colId: string;
    startX: number;
    startW: number;
    minW: number;
  } | null>(null);

  const startResize = (e: React.MouseEvent, colId: string, curWidth: number, minW = 70) => {
    e.preventDefault();
    e.stopPropagation();
    resizeRef.current = { colId, startX: e.clientX, startW: curWidth, minW };

    const onMove = (ev: MouseEvent) => {
      if (!resizeRef.current) return;
      const dx = ev.clientX - resizeRef.current.startX;
      const next = Math.max(resizeRef.current.minW, resizeRef.current.startW + dx);
      setWidths((prev) => ({ ...prev, [resizeRef.current!.colId]: next }));
    };

    const onUp = () => {
      resizeRef.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
    };

    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  };

  return { widths, setWidths, startResize };
}

export default function Dashboard() {
  const [error, setError] = useState<string | null>(null);

  // Trends
  const [weekly, setWeekly] = useState<{ weekEnd: string; prospects: number; bops: number }[]>([]);
  const [monthly, setMonthly] = useState<{ month: string; prospects: number; bops: number }[]>([]);
  const [trendLoading, setTrendLoading] = useState(false);

  // Upcoming
  const [rangeStart, setRangeStart] = useState(format(new Date(), "yyyy-MM-dd"));
  const [rangeEnd, setRangeEnd] = useState(format(addDays(new Date(), 30), "yyyy-MM-dd"));
  const [upcoming, setUpcoming] = useState<Row[]>([]);
  const [upcomingLoading, setUpcomingLoading] = useState(false);
  const [sortUpcoming, setSortUpcoming] = useState<{ key: SortKey; dir: SortDir }>({
    key: "BOP_Date",
    dir: "asc",
  });
  const [upcomingVisible, setUpcomingVisible] = useState(false);

  // Client Progress Summary
  const [progressRows, setProgressRows] = useState<Row[]>([]);
  const [progressLoading, setProgressLoading] = useState(false);
  const [progressFilter, setProgressFilter] = useState("");
  const [progressSort, setProgressSort] = useState<{ key: ProgressSortKey; dir: SortDir }>({
    key: "client_name",
    dir: "asc",
  });
  const [progressPage, setProgressPage] = useState(0);
  const [progressVisible, setProgressVisible] = useState(true);

  // Search + All Records
  const [q, setQ] = useState("");
  const [filterClient, setFilterClient] = useState("");
  const [filterInterestType, setFilterInterestType] = useState("");
  const [filterBusinessOpp, setFilterBusinessOpp] = useState("");
  const [filterWealthSolutions, setFilterWealthSolutions] = useState("");
  const [filterBopStatus, setFilterBopStatus] = useState("");
  const [filterFollowUpStatus, setFilterFollowUpStatus] = useState("");

  const [records, setRecords] = useState<Row[]>([]);
  const [allRows, setAllRows] = useState<Row[]>([]);
  const [allLoaded, setAllLoaded] = useState(false);
  const [allFetching, setAllFetching] = useState(false);
  const allRowsRef = useRef<Row[]>([]);

  useEffect(() => {
    allRowsRef.current = allRows;
  }, [allRows]);
  const [total, setTotal] = useState(0);
  const [page, setPage] = useState(0);
  const [pageJump, setPageJump] = useState("1");
  const [loading, setLoading] = useState(true);
  const [savingId, setSavingId] = useState<string | null>(null);
  const [sortAll, setSortAll] = useState<{ key: SortKey; dir: SortDir }>({
    key: "created_at",
    dir: "desc",
  });

  useEffect(() => {
    (async () => {
      try {
        const supabase = getSupabase();
        const { data } = await supabase.auth.getSession();
        if (!data.session) {
          window.location.href = "/";
          return;
        }
        await Promise.all([fetchTrends(), fetchProgressSummary(), fetchAllRecords()]);
        await loadPage(0);
      } catch (e: any) {
        setError(e?.message || "Failed to initialize");
      } finally {
        setLoading(false);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    loadPage(0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sortAll.key, sortAll.dir]);

  useEffect(() => {
    if (!allLoaded) return;
    const t = setTimeout(() => {
      setProgressPage(0);
      loadPage(0);
    }, 250);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q, allLoaded]);



  useEffect(() => {
    if (upcoming.length) fetchUpcoming();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sortUpcoming.key, sortUpcoming.dir]);

  function applySort(query: any, sort: { key: SortKey; dir: SortDir }) {
    const ascending = sort.dir === "asc";
    if (sort.key === "client")
      return query.order("first_name", { ascending }).order("last_name", { ascending });
    return query.order(sort.key, { ascending });
  }

  async function logout() {
    try {
      const supabase = getSupabase();
      await supabase.auth.signOut();
    } finally {
      window.location.href = "/";
    }
  }

  async function fetchTrends() {
    setTrendLoading(true);
    setError(null);
    try {
      const supabase = getSupabase();
      const start = startOfWeek(subWeeks(new Date(), 4), { weekStartsOn: 1 });

      const { data: createdRows, error: createdErr } = await supabase
        .from("client_registrations")
        .select("created_at, BOP_Date")
        .gte("created_at", start.toISOString())
        .order("created_at", { ascending: true })
        .limit(100000);

      if (createdErr) throw createdErr;

      const weekEnds: string[] = [];
      const weekCount = new Map<string, number>();
      const bopWeekCount = new Map<string, number>();

      for (let i = 4; i >= 0; i--) {
        const wkStart = startOfWeek(subWeeks(new Date(), i), { weekStartsOn: 1 });
        const wkEnd = endOfWeek(wkStart, { weekStartsOn: 1 });
        const key = format(wkEnd, "yyyy-MM-dd");
        weekEnds.push(key);
        weekCount.set(key, 0);
        bopWeekCount.set(key, 0);
      }

      for (const r of createdRows || []) {
        const created = parseISO(String((r as any).created_at));
        if (isValid(created)) {
          const wkEnd = endOfWeek(created, { weekStartsOn: 1 });
          const key = format(wkEnd, "yyyy-MM-dd");
          if (weekCount.has(key)) weekCount.set(key, (weekCount.get(key) || 0) + 1);
        }

        const bopRaw = (r as any).BOP_Date;
        if (bopRaw) {
          const bop = parseISO(String(bopRaw));
          if (isValid(bop)) {
            const wkEnd2 = endOfWeek(bop, { weekStartsOn: 1 });
            const key2 = format(wkEnd2, "yyyy-MM-dd");
            if (bopWeekCount.has(key2)) bopWeekCount.set(key2, (bopWeekCount.get(key2) || 0) + 1);
          }
        }
      }

      setWeekly(
        weekEnds.map((weekEnd) => ({
          weekEnd,
          prospects: weekCount.get(weekEnd) || 0,
          bops: bopWeekCount.get(weekEnd) || 0,
        }))
      );

      const yearStart = startOfYear(new Date());
      const nextYear = addYears(yearStart, 1);

      const { data: yearRows, error: yearErr } = await supabase
        .from("client_registrations")
        .select("created_at, BOP_Date")
        .gte("created_at", yearStart.toISOString())
        .lt("created_at", nextYear.toISOString())
        .order("created_at", { ascending: true })
        .limit(200000);

      if (yearErr) throw yearErr;

      const y = yearStart.getFullYear();
      const monthCount = new Map<string, number>();
      const bopMonthCount = new Map<string, number>();

      for (let m = 1; m <= 12; m++) {
        const k = `${y}-${String(m).padStart(2, "0")}`;
        monthCount.set(k, 0);
        bopMonthCount.set(k, 0);
      }

      for (const r of yearRows || []) {
        const created = parseISO(String((r as any).created_at));
        if (isValid(created)) {
          const key = format(created, "yyyy-MM");
          if (monthCount.has(key)) monthCount.set(key, (monthCount.get(key) || 0) + 1);
        }

        const bopRaw = (r as any).BOP_Date;
        if (bopRaw) {
          const bop = parseISO(String(bopRaw));
          if (isValid(bop)) {
            const key2 = format(bop, "yyyy-MM");
            if (bopMonthCount.has(key2)) bopMonthCount.set(key2, (bopMonthCount.get(key2) || 0) + 1);
          }
        }
      }

      setMonthly(
        Array.from(monthCount.keys()).map((month) => ({
          month,
          prospects: monthCount.get(month) || 0,
          bops: bopMonthCount.get(month) || 0,
        }))
      );
    } catch (e: any) {
      setError(e?.message || "Failed to load trends");
    } finally {
      setTrendLoading(false);
    }
  }

  
  async function fetchUpcoming() {
    setUpcomingLoading(true);
    setError(null);
    try {
      const supabase = getSupabase();

      // Use date-string boundaries (local midnight) to avoid timezone shifts.
      const startTs = `${rangeStart}T00:00:00`;
      const endDay = addDays(parseISO(rangeEnd), 1);
      const endTs = `${format(endDay, "yyyy-MM-dd")}T00:00:00`;

      let query = supabase
        .from("client_registrations")
        .select("*")
        .gte("BOP_Date", startTs)
        .lt("BOP_Date", endTs)
        .limit(5000);

      query = applySort(query, sortUpcoming);

      const { data, error } = await query;
      if (error) throw error;

      setUpcoming(data || []);
      setUpcomingVisible(true);
    } catch (e: any) {
      setError(e?.message || "Failed to load upcoming meetings");
    } finally {
      setUpcomingLoading(false);
    }
  }


  async function fetchProgressSummary() {
    setProgressLoading(true);
    setError(null);
    try {
      const supabase = getSupabase();
      const { data, error } = await supabase
        .from("v_client_progress_summary")
        .select(
          "clientid, first_name, last_name, phone, email, last_call_date, call_attempts, last_bop_date, bop_attempts, last_followup_date, followup_attempts"
        )
        .order("clientid", { ascending: false })
        .limit(10000);

      if (error) throw error;

      const rows = (data || []).map((r: any) => ({
        clientid: r.clientid,
        client_name: `${r.first_name || ""} ${r.last_name || ""}`.trim(),
        first_name: r.first_name,
        last_name: r.last_name,
        phone: r.phone,
        email: r.email,
        last_call_date: r.last_call_date,
        call_attempts: r.call_attempts,
        last_bop_date: r.last_bop_date,
        bop_attempts: r.bop_attempts,
        last_followup_date: r.last_followup_date,
        followup_attempts: r.followup_attempts,
      }));

      setProgressRows(rows);
      setProgressPage(0);
    } catch (e: any) {
      setError(e?.message || "Failed to load Client Progress Summary");
    } finally {
      setProgressLoading(false);
    }
  }

  
  async function fetchAllRecords() {
    if (allLoaded || allFetching) return;
    setAllFetching(true);
    setError(null);
    try {
      const supabase = getSupabase();
      const pageSize = 5000;
      let from = 0;
      const acc: Row[] = [];

      while (true) {
        const { data, error } = await supabase
          .from("client_registrations")
          .select("*")
          .order("created_at", { ascending: false })
          .range(from, from + pageSize - 1);

        if (error) throw error;

        const batch = (data || []) as Row[];
        acc.push(...batch);

        if (batch.length < pageSize) break;

        from += pageSize;
        if (from >= 200000) break; // safety cap
      }

      allRowsRef.current = acc;
      setAllRows(acc);
      setAllLoaded(true);
    } catch (e: any) {
      setError(e?.message || "Failed to load all records");
    } finally {
      setAllFetching(false);
    }
  }


  async function loadPage(nextPage: number) {
    setError(null);
    setLoading(true);
    try {
      if (!allLoaded && allRowsRef.current.length === 0) {
        await fetchAllRecords();
      }

      const needle = q.trim().toLowerCase();
      const fc = filterClient.trim().toLowerCase();
      const fi = filterInterestType.trim().toLowerCase();
      const fbo = filterBusinessOpp.trim().toLowerCase();
      const fws = filterWealthSolutions.trim().toLowerCase();
      const fb = filterBopStatus.trim().toLowerCase();
      const ffu = filterFollowUpStatus.trim().toLowerCase();

      const filtered = (allRowsRef.current || []).filter((row: any) => {
        if (needle && !matchesGlobalSearch(row, needle)) return false;

        if (fc) {
          const nm = `${row.first_name || ""} ${row.last_name || ""}`.trim().toLowerCase();
          if (!nm.includes(fc)) return false;
        }

        if (fi) {
          if (String(row.interest_type || "").toLowerCase() !== fi) return false;
        }

        if (fb) {
          if (String(row.BOP_Status || "").toLowerCase() !== fb) return false;
        }

        if (fbo) {
          const opp = asListItems(row.business_opportunities).join(",").toLowerCase();
          if (!opp.includes(fbo)) return false;
        }

        if (fws) {
          const ws = asListItems(row.wealth_solutions).join(",").toLowerCase();
          if (!ws.includes(fws)) return false;
        }

        if (ffu) {
          const fu = String(row.FollowUp_Status ?? row.Followup_Status ?? "").toLowerCase();
          if (!fu.includes(ffu)) return false;
        }

        return true;
      });

      const ascending = sortAll.dir === "asc";
      const sortKey = sortAll.key;

      const asTime = (v: any) => {
        if (!v) return Number.NaN;
        const d = new Date(String(v));
        const t = d.getTime();
        return Number.isNaN(t) ? Number.NaN : t;
      };

      const asText = (v: any) => String(v ?? "").toLowerCase();

      const sorted = [...filtered].sort((a: any, b: any) => {
        const dir = ascending ? 1 : -1;

        if (sortKey === "client") {
          return dir * asText(clientName(a)).localeCompare(asText(clientName(b)));
        }

        const va = a?.[sortKey];
        const vb = b?.[sortKey];

        const isDateKey =
          sortKey === "created_at" ||
          sortKey === "BOP_Date" ||
          sortKey === "Followup_Date" ||
          sortKey === "CalledOn" ||
          sortKey === "Issued";

        if (isDateKey) {
          const ta = asTime(va);
          const tb = asTime(vb);
          if (Number.isNaN(ta) && Number.isNaN(tb)) return 0;
          if (Number.isNaN(ta)) return 1;
          if (Number.isNaN(tb)) return -1;
          return dir * (ta - tb);
        }

        if (typeof va === "number" && typeof vb === "number") return dir * (va - vb);

        return dir * asText(va).localeCompare(asText(vb));
      });

      setTotal(sorted.length);

      const totalPagesLocal = Math.max(1, Math.ceil(sorted.length / ALL_PAGE_SIZE));
      const safePage = Math.min(Math.max(0, nextPage), totalPagesLocal - 1);

      const from = safePage * ALL_PAGE_SIZE;
      const pageRows = sorted.slice(from, from + ALL_PAGE_SIZE);

      setRecords(pageRows);
      setPage(safePage);
      setPageJump(String(safePage + 1));
    } catch (e: any) {
      setError(e?.message || "Failed to load records");
    } finally {
      setLoading(false);
    }
  }


  async function updateCell(id: string, key: string, rawValue: string) {
    setSavingId(id);
    setError(null);
    try {
      const supabase = getSupabase();
      const payload: any = {};

      const isDateTime = DATE_TIME_KEYS.has(key);
      payload[key] = isDateTime ? fromLocalInput(rawValue) : rawValue?.trim() ? rawValue : null;

      const { error } = await supabase.from("client_registrations").update(payload).eq("id", id);
      if (error) throw error;

      // Patch local state so the UI immediately shows the saved value
      const patch = (prev: Row[]) =>
        prev.map((r) => (String(r.id) === String(id) ? { ...r, [key]: payload[key] } : r));

      setRecords(patch);
      setAllRows(patch);
      setUpcoming(patch);
    } catch (e: any) {
      setError(e?.message || "Update failed");
      throw e;
    } finally {
      setSavingId(null);
    }
  }

  const totalPages = Math.max(1, Math.ceil((total || 0) / ALL_PAGE_SIZE));
  const canPrev = page > 0;
  const canNext = (page + 1) * ALL_PAGE_SIZE < total;

  const exportUpcomingXlsx = () => {
    const ws = XLSX.utils.json_to_sheet(upcoming);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Upcoming_BOP");
    XLSX.writeFile(wb, `Upcoming_BOP_${rangeStart}_to_${rangeEnd}.xlsx`);
  };

  const sortHelp = (
    <div className="text-xs text-slate-600">
      Click headers to sort: <b>Client Name</b>, <b>Created Date</b>, <b>BOP Date</b>, <b>BOP Status</b>,{" "}
      <b>Follow-Up Date</b>, <b>Status</b>.
    </div>
  );

  const extraClientCol = useMemo(
    () => [{ label: "Client Name", sortable: "client" as SortKey, render: (r: Row) => clientName(r) }],
    []
  );

  // -------- Progress Summary (filter/sort/paginate client-side) --------
  const progressFilteredSorted = useMemo(() => {
    const needle = q.trim().toLowerCase();

    const filtered = (progressRows || []).filter((r: any) => {
      if (!needle) return true;
      return matchesGlobalSearch(r, needle);
    });

    const dirMul = progressSort.dir === "asc" ? 1 : -1;

    const asNum = (v: any) => {
      const n = Number(v);
      return Number.isFinite(n) ? n : 0;
    };

    const asTime = (v: any) => {
      if (!v) return 0;
      const d = new Date(v);
      const t = d.getTime();
      return Number.isFinite(t) ? t : 0;
    };

    filtered.sort((a, b) => {
      const k = progressSort.key;
      if (k === "client_name") {
        return String(a.client_name || "").localeCompare(String(b.client_name || "")) * dirMul;
      }
      if (k === "call_attempts" || k === "bop_attempts" || k === "followup_attempts") {
        return (asNum(a[k]) - asNum(b[k])) * dirMul;
      }
      // date keys
      return (asTime(a[k]) - asTime(b[k])) * dirMul;
    });

    return filtered;
  }, [progressRows, q, progressSort]);

  const upcomingFiltered = useMemo(() => {
    const needle = q.trim().toLowerCase();
    if (!needle) return upcoming;
    return (upcoming || []).filter((r: any) => matchesGlobalSearch(r, needle));
  }, [upcoming, q]);

  const progressTotalPages = Math.max(1, Math.ceil(progressFilteredSorted.length / PROGRESS_PAGE_SIZE));
  const progressPageSafe = Math.min(progressTotalPages - 1, Math.max(0, progressPage));
  const progressSlice = progressFilteredSorted.slice(
    progressPageSafe * PROGRESS_PAGE_SIZE,
    progressPageSafe * PROGRESS_PAGE_SIZE + PROGRESS_PAGE_SIZE
  );

  return (
    <div className="min-h-screen">
      <div className="max-w-[1600px] mx-auto p-6 space-y-6">
        <header className="flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <img src="/can-logo.png" className="h-10 w-auto" alt="CAN Financial Solutions" />
            <div>
              <div className="text-2xl font-bold text-slate-800">CAN Financial Solutions Clients Report</div>
              <div className="text-sm text-slate-500">Excel-style tables, editable follow-ups, and trends</div>
            </div>
          </div>
          <Button variant="secondary" onClick={logout}>
            Logout
          </Button>
        </header>

        {error && <div className="rounded-xl border border-red-200 bg-red-50 p-4 text-red-700">{error}</div>}

        {/* Trends */}
        <Card title="Trends">
          <div className="flex items-center justify-end mb-3">
            <Button variant="secondary" onClick={fetchTrends}>
              Refresh
            </Button>
          </div>

          <div className="grid lg:grid-cols-2 gap-6">
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-2">Weekly (Last 5 Weeks)</div>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={weekly}>
                    <XAxis dataKey="weekEnd" tick={{ fontSize: 11 }} />
                    <YAxis allowDecimals={false} />
                    <Tooltip />
                    <Line
                      type="monotone"
                      dataKey="prospects"
                      stroke="#2563eb"
                      strokeWidth={2}
                      dot={{ r: 3 }}
                      activeDot={{ r: 5 }}
                    >
                      <LabelList dataKey="prospects" position="top" fill="#0f172a" />
                    </Line>
                    <Line
                      type="monotone"
                      dataKey="bops"
                      stroke="#f97316"
                      strokeWidth={2}
                      dot={{ r: 3 }}
                      activeDot={{ r: 5 }}
                    >
                      <LabelList dataKey="bops" position="top" fill="#0f172a" />
                    </Line>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div>
              <div className="text-xs font-semibold text-slate-600 mb-2">Monthly (Current Year)</div>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={monthly}>
                    <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                    <YAxis allowDecimals={false} />
                    <Tooltip />
                    <Bar dataKey="prospects" fill="#22c55e">
                      <LabelList dataKey="prospects" position="top" fill="#0f172a" />
                    </Bar>
                    <Bar dataKey="bops" fill="#a855f7">
                      <LabelList dataKey="bops" position="top" fill="#0f172a" />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>

          {trendLoading && <div className="mt-2 text-xs text-slate-500">Loading…</div>}
        </Card>

        {/* Upcoming Range */}
        <Card title="Upcoming BOP Date Range">
          <div className="grid md:grid-cols-5 gap-3 items-end">
            <label className="block md:col-span-2">
              <div className="text-xs font-semibold text-slate-600 mb-1">Start</div>
              <input
                type="date"
                className="w-full border border-slate-300 px-3 py-2"
                value={rangeStart}
                onChange={(e) => setRangeStart(e.target.value)}
              />
            </label>
            <label className="block md:col-span-2">
              <div className="text-xs font-semibold text-slate-600 mb-1">End</div>
              <input
                type="date"
                className="w-full border border-slate-300 px-3 py-2"
                value={rangeEnd}
                onChange={(e) => setRangeEnd(e.target.value)}
              />
            </label>
            <div className="flex gap-2 md:col-span-1">
              <Button onClick={fetchUpcoming} disabled={upcomingLoading}>
                {upcomingLoading ? "Loading…" : "Load"}
              </Button>
              <Button variant="secondary" onClick={exportUpcomingXlsx} disabled={upcoming.length === 0}>
                Export XLSX
              </Button>
            </div>
          </div>

          <div className="mt-3">
            <Button
              variant="secondary"
              onClick={() => setUpcomingVisible((v) => !v)}
              disabled={upcomingLoading}
            >
              {upcomingVisible ? "Hide Upcoming Table" : "Show Upcoming Table"}
            </Button>
            <span className="ml-3 text-xs text-slate-500">
              After you press Load, you can show/hide the Upcoming table.
            </span>
          </div>
        </Card>

        {/* Upcoming Table */}
        {upcomingVisible && upcomingFiltered.length > 0 && (
          <Card title="Upcoming BOP Meetings (Editable)">
            <div className="flex items-center justify-between mb-2">
              <div className="text-sm text-slate-600">Table supports vertical + horizontal scrolling.</div>
              {sortHelp}
            </div>

            <ExcelTableEditable
              rows={upcomingFiltered}
              savingId={savingId}
              onUpdate={updateCell}
              preferredOrder={["BOP_Date", "created_at", "BOP_Status", "Followup_Date", "status"]}
              extraLeftCols={[
                { label: "Client Name", sortable: "client", render: (r) => clientName(r) },
              ]}
              maxHeightClass="max-h-[420px]"
              sortState={sortUpcoming}
              onSortChange={(k) => setSortUpcoming((cur) => toggleSort(cur, k))}
              stickyLeftCount={1}
            />
          </Card>
        )}

        {/* Search */}
        <Card title="Search">
          <div className="flex flex-col md:flex-row gap-2 md:items-center">
            <input
              className="w-full border border-slate-300 px-4 py-3"
              placeholder="Search by first name, last name, or phone"
              value={q}
              onChange={(e) => setQ(e.target.value)}
            />
            <Button onClick={() => loadPage(0)}>Go</Button>
            <div className="text-sm text-slate-600 md:ml-auto">
              {total.toLocaleString()} records • showing {ALL_PAGE_SIZE} per page
            </div>
          </div>

          <div className="mt-3 grid md:grid-cols-3 lg:grid-cols-6 gap-2">
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">Client Name</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterClient}
                onChange={(e) => setFilterClient(e.target.value)}
                placeholder="Contains…"
              />
            </div>
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">Interest Type</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterInterestType}
                onChange={(e) => setFilterInterestType(e.target.value)}
                placeholder="e.g., client"
              />
            </div>
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">Business Opportunities</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterBusinessOpp}
                onChange={(e) => setFilterBusinessOpp(e.target.value)}
                placeholder="Contains…"
              />
            </div>
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">Wealth Solutions</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterWealthSolutions}
                onChange={(e) => setFilterWealthSolutions(e.target.value)}
                placeholder="Contains…"
              />
            </div>
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">BOP Status</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterBopStatus}
                onChange={(e) => setFilterBopStatus(e.target.value)}
                placeholder="e.g., scheduled"
              />
            </div>
            <div>
              <div className="text-xs font-semibold text-slate-600 mb-1">Follow-Up Status</div>
              <input
                className="w-full border border-slate-300 px-3 py-2 text-sm"
                value={filterFollowUpStatus}
                onChange={(e) => setFilterFollowUpStatus(e.target.value)}
                placeholder="e.g., pending"
              />
            </div>
          </div>

          <div className="mt-2 text-xs text-slate-600">
            Tip: Enter filters and click <b>Go</b> to apply. Page size is {ALL_PAGE_SIZE}.
          </div>
        </Card>

        {/* All Records */}
        <Card title="All Records (Editable)">
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3 mb-2">
            <div className="text-sm text-slate-600">
              Page <b>{page + 1}</b> of <b>{totalPages}</b>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {sortHelp}
              <div className="flex items-center gap-2 border border-slate-300 px-3 py-2 bg-white">
                <span className="text-xs font-semibold text-slate-600">Search</span>
                <input
                  className="w-56 border border-slate-300 px-2 py-1 text-sm"
                  value={q}
                  onChange={(e) => setQ(e.target.value)}
                  placeholder="Name / phone / email"
                />
              </div>
              <div className="flex items-center gap-2 border border-slate-300 px-3 py-2 bg-white">
                <span className="text-xs font-semibold text-slate-600">Go to page</span>
                <input
                  type="number"
                  min={1}
                  max={totalPages}
                  className="w-20 border border-slate-300 px-2 py-1 text-sm"
                  value={pageJump}
                  onChange={(e) => setPageJump(e.target.value)}
                />
                <Button
                  variant="secondary"
                  onClick={() => {
                    const n = Number(pageJump);
                    if (!Number.isFinite(n)) return;
                    const p = Math.min(totalPages, Math.max(1, Math.floor(n)));
                    loadPage(p - 1);
                  }}
                  disabled={loading || totalPages <= 1}
                >
                  Go
                </Button>
              </div>
              <Button variant="secondary" onClick={() => loadPage(Math.max(0, page - 1))} disabled={!canPrev || loading}>
                Previous
              </Button>
              <Button variant="secondary" onClick={() => loadPage(page + 1)} disabled={!canNext || loading}>
                Next
              </Button>
            </div>
          </div>

          {loading ? (
            <div className="text-slate-600">Loading…</div>
          ) : (
            <ExcelTableEditable
              rows={records}
              savingId={savingId}
              onUpdate={updateCell}
              extraLeftCols={extraClientCol}
              maxHeightClass="max-h-[560px]"
              sortState={sortAll}
              onSortChange={(k) => setSortAll((cur) => toggleSort(cur, k))}
              stickyLeftCount={1}
            />
          )}
        </Card>

        {/* Client Progress Summary */}
        <Card title="Client Progress Summary">
          <div className="flex flex-col md:flex-row md:items-center gap-2 mb-2">
            <input
              className="w-full border border-slate-300 px-4 py-3"
              placeholder="Filter by client name..."
              value={q}
              onChange={(e) => {
                setQ(e.target.value);
                setProgressPage(0);
              }}
            />
            <Button variant="secondary" onClick={fetchProgressSummary} disabled={progressLoading}>
              {progressLoading ? "Loading…" : "Refresh"}
            </Button>
            <Button variant="secondary" onClick={() => setProgressVisible((v) => !v)}>
              {progressVisible ? "Hide Table" : "Show Table"}
            </Button>

            <div className="md:ml-auto flex items-center gap-2">
              <Button
                variant="secondary"
                onClick={() => setProgressPage((p) => Math.max(0, p - 1))}
                disabled={!progressVisible || progressPageSafe <= 0}
              >
                Previous
              </Button>
              <Button
                variant="secondary"
                onClick={() => setProgressPage((p) => Math.min(progressTotalPages - 1, p + 1))}
                disabled={!progressVisible || progressPageSafe >= progressTotalPages - 1}
              >
                Next
              </Button>
            </div>
          </div>

          <div className="text-xs text-slate-600 mb-2">Click headers to sort.</div>

          {progressVisible && (
            <ProgressSummaryTable
              rows={progressSlice}
              sortState={progressSort}
              onSortChange={(k) => setProgressSort((cur) => toggleProgressSort(cur, k))}
            />
          )}

          {progressVisible && (
            <div className="mt-2 text-xs text-slate-600">
              Page <b>{progressPageSafe + 1}</b> of <b>{progressTotalPages}</b> • showing {PROGRESS_PAGE_SIZE} per page
            </div>
          )}
        </Card>
      </div>
    </div>
  );
}

/** ---------------- Progress Summary Table (Resizable + Sticky Client Name) ---------------- */
function ProgressSummaryTable({
  rows,
  sortState,
  onSortChange,
}: {
  rows: Row[];
  sortState: { key: ProgressSortKey; dir: SortDir };
  onSortChange: (k: ProgressSortKey) => void;
}) {
  const { widths, startResize } = useColumnResizer();

  const cols = useMemo(
    () => [
      { id: "client_name", label: "Client Name", key: "client_name" as ProgressSortKey, defaultW: 170 },
      { id: "first_name", label: "First Name", defaultW: 95 },
      { id: "last_name", label: "Last Name", defaultW: 90 },
      { id: "phone", label: "Phone", defaultW: 105 },
      { id: "email", label: "Email", defaultW: 220 },
      { id: "last_call_date", label: "Last Call On", key: "last_call_date" as ProgressSortKey, defaultW: 190 },
      { id: "call_attempts", label: "No of Calls", key: "call_attempts" as ProgressSortKey, defaultW: 90 },
      { id: "last_bop_date", label: "Last BOP Call On", key: "last_bop_date" as ProgressSortKey, defaultW: 200 },
      { id: "bop_attempts", label: "No of BOP Calls", key: "bop_attempts" as ProgressSortKey, defaultW: 110 },
      { id: "last_followup_date", label: "Last FollowUp On", key: "last_followup_date" as ProgressSortKey, defaultW: 200 },
      { id: "followup_attempts", label: "No of FollowUp Calls", key: "followup_attempts" as ProgressSortKey, defaultW: 140 },
    ],
    []
  );

  const getW = (id: string, def: number) => widths[id] ?? def;

  const stickyLeftPx = (colIndex: number) => {
    // only first col sticky
    if (colIndex <= 0) return 0;
    return 0;
  };

  const sortIcon = (k?: ProgressSortKey) => {
    if (!k) return null;
    if (sortState.key !== k) return <span className="ml-1 text-slate-400">↕</span>;
    return <span className="ml-1 text-slate-700">{sortState.dir === "asc" ? "↑" : "↓"}</span>;
  };

  const minWidth = cols.reduce((sum, c) => sum + getW(c.id, c.defaultW), 0);

  const fmtDate = (v: any) => {
    if (!v) return "";
    const d = new Date(v);
    const t = d.getTime();
    if (!Number.isFinite(t)) return "";
    return d.toLocaleString();
  };

  const fmtZeroBlank = (v: any) => {
    const n = Number(v);
    if (!Number.isFinite(n) || n === 0) return "";
    return String(n);
  };

  return (
    <div className="overflow-auto border border-slate-500 bg-white max-h-[520px]">
      <table className="w-full table-fixed border-collapse" style={{ minWidth }}>
        <thead className="sticky top-0 bg-slate-100 z-20">
          <tr className="text-left text-xs font-semibold text-slate-700">
            {cols.map((c, idx) => {
              const w = getW(c.id, c.defaultW);
              const isSticky = idx === 0;
              const style: React.CSSProperties = {
                width: w,
                minWidth: w,
                maxWidth: w,
                position: isSticky ? "sticky" : undefined,
                left: isSticky ? stickyLeftPx(idx) : undefined,
                top: 0,
                zIndex: isSticky ? 40 : 20,
                background: isSticky ? "#f1f5f9" : undefined,
              };

              return (
                <th key={c.id} className="border border-slate-500 px-2 py-2 whitespace-nowrap relative" style={style}>
                  {c.key ? (
                    <button
                      className="inline-flex items-center hover:underline"
                      onClick={() => onSortChange(c.key!)}
                      type="button"
                    >
                      {c.label}
                      {sortIcon(c.key)}
                    </button>
                  ) : (
                    c.label
                  )}

                  <div
                    className="absolute top-0 right-0 h-full w-2 cursor-col-resize select-none"
                    onMouseDown={(e) => startResize(e, c.id, w)}
                  >
                    <div className="mx-auto h-full w-px bg-slate-300" />
                  </div>
                </th>
              );
            })}
          </tr>
        </thead>

        <tbody>
          {rows.map((r, ridx) => (
            <tr key={String(r.clientid ?? ridx)} className="hover:bg-slate-50">
              {cols.map((c, idx) => {
                const w = getW(c.id, c.defaultW);
                const isSticky = idx === 0;
                const style: React.CSSProperties = {
                  width: w,
                  minWidth: w,
                  maxWidth: w,
                  position: isSticky ? "sticky" : undefined,
                  left: isSticky ? stickyLeftPx(idx) : undefined,
                  zIndex: isSticky ? 10 : 1,
                  background: isSticky ? "#ffffff" : undefined,
                };

                let v = "";
                if (c.id === "client_name") v = String(r.client_name || "");
                else if (c.id === "first_name") v = String(r.first_name || "");
                else if (c.id === "last_name") v = String(r.last_name || "");
                else if (c.id === "phone") v = String(r.phone || "");
                else if (c.id === "email") v = String(r.email || "");
                else if (c.id === "last_call_date") v = fmtDate(r.last_call_date);
                else if (c.id === "call_attempts") v = fmtZeroBlank(r.call_attempts);
                else if (c.id === "last_bop_date") v = fmtDate(r.last_bop_date);
                else if (c.id === "bop_attempts") v = fmtZeroBlank(r.bop_attempts);
                else if (c.id === "last_followup_date") v = fmtDate(r.last_followup_date);
                else if (c.id === "followup_attempts") v = fmtZeroBlank(r.followup_attempts);

                return (
                  <td
                    key={c.id}
                    className={`border border-slate-300 px-2 py-2 whitespace-nowrap ${
                      isSticky ? "font-semibold text-slate-800" : ""
                    }`}
                    style={style}
                  >
                    {v}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** ---------------- Editable Excel-style table (Resizable + Sticky first column + Date saves reliably) ---------------- */
function ExcelTableEditable({
  rows,
  savingId,
  onUpdate,
  extraLeftCols,
  maxHeightClass,
  sortState,
  onSortChange,
  preferredOrder,
  stickyLeftCount = 1,
}: {
  rows: Row[];
  savingId: string | null;
  onUpdate: (id: string, key: string, value: string) => Promise<void>;
  extraLeftCols: { label: string; render: (r: Row) => string; sortable?: SortKey }[];
  maxHeightClass: string;
  sortState: { key: SortKey; dir: SortDir };
  onSortChange: (key: SortKey) => void;
  preferredOrder?: string[];
  stickyLeftCount?: number;
}) {
  const { widths, startResize } = useColumnResizer();
  const [openCell, setOpenCell] = useState<string | null>(null);
  const [drafts, setDrafts] = useState<Record<string, string>>({});

  const sortIcon = (k?: SortKey) => {
    if (!k) return null;
    if (sortState.key !== k) return <span className="ml-1 text-slate-400">↕</span>;
    return <span className="ml-1 text-slate-700">{sortState.dir === "asc" ? "↑" : "↓"}</span>;
  };

  const keys = useMemo(() => {
    if (!rows.length) return [] as string[];
    const baseKeys = Object.keys(rows[0]).filter((k) => k !== "id");
    if (!preferredOrder || !preferredOrder.length) return baseKeys;
    const set = new Set(baseKeys);
    const ordered: string[] = [];
    for (const k of preferredOrder) if (set.has(k)) ordered.push(k);
    for (const k of baseKeys) if (!ordered.includes(k)) ordered.push(k);
    return ordered;
  }, [rows, preferredOrder]);

  // Column models (extra cols + keys)
  const columns = useMemo(() => {
    const extra = extraLeftCols.map((c, i) => ({
      id: `extra:${i}`,
      label: c.label,
      sortable: c.sortable,
      kind: "extra" as const,
      defaultW: c.label.toLowerCase().includes("client") ? 180 : 150,
    }));

    const main = keys.map((k) => {
      const label = labelFor(k);
      const isDateTime = DATE_TIME_KEYS.has(k);
      const defaultW =
        k === "created_at" ? 120 : isDateTime ? 210 : k.toLowerCase().includes("email") ? 240 : 160;

      const sortable =
        k === "created_at"
          ? ("created_at" as SortKey)
          : k === "BOP_Date"
          ? ("BOP_Date" as SortKey)
          : k === "BOP_Status"
          ? ("BOP_Status" as SortKey)
          : k === "Followup_Date"
          ? ("Followup_Date" as SortKey)
          : k === "status"
          ? ("status" as SortKey)
          : k === "CalledOn"
          ? ("CalledOn" as SortKey)
          : k === "Issued"
          ? ("Issued" as SortKey)
          : undefined;

      return {
        id: `col:${k}`,
        key: k,
        label,
        sortable,
        kind: "data" as const,
        defaultW,
      };
    });

    return [...extra, ...main];
  }, [extraLeftCols, keys]);

  const getW = (id: string, def: number) => widths[id] ?? def;

  const stickyLeftPx = (colIndex: number) => {
    let left = 0;
    for (let i = 0; i < colIndex; i++) {
      const c = columns[i];
      left += getW(c.id, (c as any).defaultW || 160);
    }
    return left;
  };

  const minWidth = columns.reduce((sum, c: any) => sum + getW(c.id, c.defaultW || 160), 0);

  const getCellValueForInput = (r: Row, k: string) => {
    const isDateTime = DATE_TIME_KEYS.has(k);
    const val = r[k];
    if (isDateTime) return toLocalInput(val);
    return val ?? "";
  };

  const handleBlur = async (rowId: string, key: string, cellId: string) => {
    const v = drafts[cellId] ?? "";
    try {
      await onUpdate(String(rowId), key, v);
    } finally {
      // After save attempt, keep UI stable: clear draft so it renders from updated row value
      setDrafts((prev) => {
        const next = { ...prev };
        delete next[cellId];
        return next;
      });
    }
  };

  return (
    <div className={`overflow-auto border border-slate-500 bg-white ${maxHeightClass}`}>
      <table className="w-full table-fixed border-collapse" style={{ minWidth }}>
        <thead className="sticky top-0 bg-slate-100 z-20">
          <tr className="text-left text-xs font-semibold text-slate-700">
            {columns.map((c: any, colIndex: number) => {
              const w = getW(c.id, c.defaultW || 160);
              const isSticky = colIndex < stickyLeftCount;
              const isTopLeft = isSticky; // header row is always top sticky
              const style: React.CSSProperties = {
                width: w,
                minWidth: w,
                maxWidth: w,
                position: isSticky ? "sticky" : undefined,
                left: isSticky ? stickyLeftPx(colIndex) : undefined,
                top: 0,
                zIndex: isTopLeft ? 50 : 20,
                background: isSticky ? "#f1f5f9" : undefined,
              };

              const headerLabel = c.kind === "extra" ? c.label : c.label;

              return (
                <th
                  key={c.id}
                  className="border border-slate-500 px-2 py-2 whitespace-nowrap relative"
                  style={style}
                >
                  {c.sortable ? (
                    <button
                      className="inline-flex items-center hover:underline"
                      onClick={() => onSortChange(c.sortable)}
                      type="button"
                    >
                      {headerLabel}
                      {sortIcon(c.sortable)}
                    </button>
                  ) : (
                    headerLabel
                  )}

                  {/* Resize handle for EVERY column */}
                  <div
                    className="absolute top-0 right-0 h-full w-2 cursor-col-resize select-none"
                    onMouseDown={(e) => startResize(e, c.id, w)}
                  >
                    <div className="mx-auto h-full w-px bg-slate-300" />
                  </div>
                </th>
              );
            })}
          </tr>
        </thead>

        <tbody>
          {rows.map((r, ridx) => (
            <tr key={String(r.id ?? ridx)} className="hover:bg-slate-50">
              {columns.map((c: any, colIndex: number) => {
                const w = getW(c.id, c.defaultW || 160);
                const isSticky = colIndex < stickyLeftCount;
                const style: React.CSSProperties = {
                  width: w,
                  minWidth: w,
                  maxWidth: w,
                  position: isSticky ? "sticky" : undefined,
                  left: isSticky ? stickyLeftPx(colIndex) : undefined,
                  zIndex: isSticky ? 10 : 1,
                  background: isSticky ? "#ffffff" : undefined,
                };

                // EXTRA COLUMNS (non-editable)
                if (c.kind === "extra") {
                  const idx = Number(String(c.id).split(":")[1] || "0");
                  const colDef = extraLeftCols[idx];
                  const v = colDef?.render ? colDef.render(r) : "";
                  return (
                    <td
                      key={c.id}
                      className={`border border-slate-300 px-2 py-2 whitespace-nowrap font-semibold text-slate-800`}
                      style={style}
                    >
                      {v}
                    </td>
                  );
                }

                const k = c.key as string;

                // created_at shown as date (not editable)
                if (k === "created_at") {
                  const d = new Date(r.created_at);
                  const v = Number.isNaN(d.getTime()) ? "" : d.toLocaleDateString();
                  return (
                    <td key={c.id} className="border border-slate-300 px-2 py-2 whitespace-nowrap" style={style}>
                      {v}
                    </td>
                  );
                }

                // read-only list cols with dropdown
                if (READONLY_LIST_COLS.has(k)) {
                  const cellId = `${r.id}:${k}`;
                  const items = asListItems(r[k]);
                  const display = items.join(", ");
                  return (
                    <td key={c.id} className="border border-slate-300 px-2 py-2 align-top" style={style}>
                      <div className="relative">
                        <button
                          type="button"
                          className="w-full text-left text-slate-800 whitespace-normal break-words"
                          onClick={() => setOpenCell((cur) => (cur === cellId ? null : cellId))}
                        >
                          {display || "—"}
                        </button>

                        {openCell === cellId && (
                          <div className="absolute left-0 top-full mt-1 w-72 max-w-[70vw] bg-white border border-slate-500 shadow-lg z-30">
                            <div className="px-2 py-1 text-xs font-semibold text-slate-700 bg-slate-100 border-b border-slate-300">
                              {labelFor(k)}
                            </div>
                            <ul className="max-h-48 overflow-auto">
                              {(items.length ? items : ["(empty)"]).map((x, i) => (
                                <li key={i} className="px-2 py-1 text-sm border-b border-slate-100">
                                  {x}
                                </li>
                              ))}
                            </ul>
                            <div className="p-2">
                              <Button variant="secondary" onClick={() => setOpenCell(null)}>
                                Close
                              </Button>
                            </div>
                          </div>
                        )}
                      </div>
                    </td>
                  );
                }

                // EDITABLE CELLS (Controlled inputs so selected dates always stay visible)
                const cellId = `${r.id}:${k}`;
                const isDateTime = DATE_TIME_KEYS.has(k);

                const value =
                  drafts[cellId] !== undefined ? drafts[cellId] : String(getCellValueForInput(r, k));

                return (
                  <td key={c.id} className="border border-slate-300 px-2 py-2" style={style}>
                    <input
                      type={isDateTime ? "datetime-local" : "text"}
                      className="w-full bg-transparent border-0 outline-none text-sm"
                      value={value}
                      onChange={(e) => setDrafts((prev) => ({ ...prev, [cellId]: e.target.value }))}
                      onBlur={() => handleBlur(String(r.id), k, cellId)}
                      disabled={savingId != null && String(savingId) === String(r.id)}
                    />
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
