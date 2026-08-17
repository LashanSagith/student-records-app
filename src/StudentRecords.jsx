import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { useMsal } from "@azure/msal-react";
import {
  Search, UserPlus, Upload, Download, X, Pencil, Trash2, ChevronRight,
  Users, GraduationCap, UserX, FileSpreadsheet, SlidersHorizontal, ChevronDown, GripVertical,
  History, Cloud, CloudOff, RefreshCw, AlertCircle, LogOut, User as UserIcon,
} from "lucide-react";
import { loadFromSharePoint, saveToSharePoint } from "./graphService";

/* ----------------------------- design tokens ----------------------------- */
const THEME = {
  bg: "#E9EBEE",
  surface: "#FFFFFF",
  ink: "#616771",
  inkMuted: "#90949C",
  inkFaint: "#B7BAC0",
  border: "#DADCE0",
  borderStrong: "#90949C",
  primary: "#4267B3",
  primaryDark: "#35538F",
  primarySoft: "#E4EAF4",
  success: "#2F7D5B",
  successSoft: "#E2F0E8",
  danger: "#B23A48",
  dangerSoft: "#F6E3E5",
  neutralSoft: "#EDEEEB",
};

const FONT_DISPLAY = "'Space Grotesk', sans-serif";
const FONT_BODY = "'Inter', sans-serif";
const FONT_MONO = "'IBM Plex Mono', monospace";

const HIGHLIGHTS_BY_STATUS = {
  "Active": { color: "#FFFFFF", label: "Active" },
  "Withdrawn": { color: "#FFA8EC", label: "Withdrawn" },
  "Dropped": { color: "#FFA8EC", label: "Dropped" },
  "LOA & Transfer": { color: "#FAE466", label: "LOA & Transfer" },
  "FRA": { color: "#7EE1F7", label: "FRA" },
  "ESR": { color: "#B8A6F0", label: "ESR" },
};
const HIGHLIGHTS = HIGHLIGHTS_BY_STATUS;

/* ----------------------------- data schema ----------------------------- */
const SECTIONS = [
  { id: "identification", label: "Identification" },
  { id: "contact", label: "Contact & Address" },
  { id: "programme", label: "Programme" },
  { id: "ol_al", label: "O/L & A/L" },
  { id: "status", label: "Status & Documents" },
  { id: "level_progress", label: "Level 04 & 05" },
  { id: "award_grad", label: "Award & Graduation" },
  { id: "comments", label: "Special Comments" },
];

const FIELDS = [
  { key: "dbr", label: "DBR", section: "identification" },
  { key: "applicationRefNo", label: "Application Ref No", section: "identification", mono: true },
  { key: "icbtRegNo", label: "ICBT Reg No", section: "identification", mono: true },
  { key: "certificationName", label: "Certification Name (International Students Only, see guidance tab)", section: "identification" },
  { key: "nicPassport", label: "NIC/Passport", section: "identification", mono: true },
  { key: "ljmuId", label: "LJMU ID", section: "identification", mono: true },

  { key: "homePhone", label: "Home Phone Number", section: "contact", subsection: "Contact" },
  { key: "mobile", label: "Mobile Number", section: "contact", subsection: "Contact", mono: true },
  { key: "email", label: "Email", section: "contact", subsection: "Contact" },
  { key: "country", label: "Country", section: "contact", subsection: "Address" },
  { key: "addressLine1", label: "Line 1", section: "contact", subsection: "Address" },
  { key: "addressLine2", label: "Line 2", section: "contact", subsection: "Address" },
  { key: "addressLine3", label: "Line 3", section: "contact", subsection: "Address" },
  { key: "addressLine4", label: "Line 4", section: "contact", subsection: "Address" },
  { key: "city", label: "City", section: "contact", subsection: "Address" },
  { key: "postCode", label: "Post Code", section: "contact", subsection: "Address" },

  { key: "career", label: "Career", section: "programme" },
  { key: "admitTerm", label: "Admit Term", section: "programme" },
  { key: "month", label: "Month", section: "programme" },
  { key: "programmeCode", label: "Programme Code", section: "programme", mono: true },
  { key: "planCode", label: "Plan Code", section: "programme", mono: true },
  { key: "academicLoad", label: "Academic Load", section: "programme", suggestions: ["Full-Time", "Part-Time"] },
  { key: "nqfLevel", label: "NQF Level", section: "programme" },
  { key: "international", label: "International?", section: "programme", suggestions: ["Yes", "No"] },

  { key: "olLocalInternational", label: "Local / International", section: "ol_al", subsection: "O/L Results", suggestions: ["Local", "International"] },
  { key: "olMaths", label: "Maths", section: "ol_al", subsection: "O/L Results" },
  { key: "olEnglish", label: "English", section: "ol_al", subsection: "O/L Results" },
  { key: "alLocalInternational", label: "Local / International", section: "ol_al", subsection: "A/L Results", suggestions: ["Local", "International"] },
  { key: "alStream", label: "Stream", section: "ol_al", subsection: "A/L Results" },
  { key: "alResult", label: "Result", section: "ol_al", subsection: "A/L Results" },

  { key: "offerStatus", label: "Offer Status", section: "status", suggestions: ["Offer Issued", "Conditional Offer", "Unconditional Offer", "Pending", "Rejected"] },
  { key: "classListConfirmation", label: "Class List Confirmation", section: "status", suggestions: ["Confirmed", "Pending"] },
  { key: "dropReason", label: "Drop (Reason)", section: "status", suggestions: ["Financial", "Academic", "Personal", "Transferred", "Other"] },
  { key: "withdrawalFormStatus", label: "Withdrawal Form Status", section: "status", suggestions: ["Not Submitted", "Submitted", "Approved"] },
  { key: "canvasActivationStatus", label: "Canvas Activation Status", section: "status", suggestions: ["Activated", "Pending", "Not Activated"] },
  { key: "idPhotoUpload", label: "ID Photo Upload", section: "status", suggestions: ["Uploaded", "Pending"] },
  { key: "ljmuIdCard", label: "LJMU ID Card", section: "status", suggestions: ["Issued", "Pending"] },
  { key: "activeStatus", label: "Active Status", section: "status", type: "select", options: ["Active", "Withdrawn", "Dropped", "LOA & Transfer", "FRA", "ESR"] },

  { key: "l4FirstAttempt", label: "L4 - First Attempt", section: "level_progress", subsection: "Level 04" },
  { key: "l4SecondAttempt", label: "L4 - Second Attempt", section: "level_progress", subsection: "Level 04" },
  { key: "l4ThirdAttempt", label: "L4 - Third Attempt", section: "level_progress", subsection: "Level 04" },
  { key: "l4Esr", label: "L4 - ESR", section: "level_progress", subsection: "Level 04" },
  { key: "level04Status", label: "Level 04 Status", section: "level_progress", subsection: "Level 04", suggestions: ["Pass", "Progressing", "Repeat", "Fail"] },
  { key: "l5FirstAttempt", label: "L5 - First Attempt", section: "level_progress", subsection: "Level 05" },
  { key: "l5SecondAttempt", label: "L5 - Second Attempt", section: "level_progress", subsection: "Level 05" },
  { key: "l5ThirdAttempt", label: "L5 - Third Attempt", section: "level_progress", subsection: "Level 05" },
  { key: "l5Esr", label: "L5 - ESR", section: "level_progress", subsection: "Level 05" },
  { key: "level05Status", label: "Level 05 Status", section: "level_progress", subsection: "Level 05", suggestions: ["Pass", "Progressing", "Repeat", "Fail"] },

  { key: "award", label: "Award", section: "award_grad", subsection: "Award" },
  { key: "awardMark", label: "Award Mark", section: "award_grad", subsection: "Award" },
  { key: "graduationEligibility", label: "Graduation Eligibility", section: "award_grad", subsection: "Graduation", suggestions: ["Eligible", "Not Eligible", "Pending"] },
  { key: "graduatedYear", label: "Graduated Year", section: "award_grad", subsection: "Graduation" },
  { key: "certificateReceivedDate", label: "Certificate Received Date", section: "award_grad", subsection: "Graduation", type: "date" },
  { key: "certificateHandedOverDate", label: "Certificate Handed over date to Registry", section: "award_grad", subsection: "Graduation", type: "date" },

  { key: "specialComments", label: "Special Comments", section: "comments", textarea: true },
];

const FIELD_BY_KEY = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

const SEARCH_KEYS = ["dbr", "applicationRefNo", "icbtRegNo", "nicPassport", "ljmuId", "email", "mobile", "certificationName", "programmeCode"];

const TABLE_COLUMNS = [
  "dbr", "applicationRefNo", "icbtRegNo", "ljmuId", "certificationName", "admitTerm",
  "programmeCode", "nicPassport", "offerStatus", "canvasActivationStatus", "level04Status",
  "level05Status", "graduationEligibility", "graduatedYear", "specialComments", "activeStatus",
];
const MONO_COLUMNS = new Set(["dbr", "applicationRefNo", "icbtRegNo", "ljmuId", "nicPassport", "programmeCode"]);
const PILL_COLUMNS = new Set(["offerStatus", "canvasActivationStatus", "level04Status", "level05Status", "graduationEligibility", "activeStatus"]);
const TRUNCATE_COLUMNS = new Set(["certificationName", "specialComments"]);

const DRAG_HANDLE_WIDTH = 32;
const FROZEN_WIDTH = { dbr: 100, applicationRefNo: 150, icbtRegNo: 150, ljmuId: 130 };
const FROZEN_KEYS = ["dbr", "applicationRefNo", "icbtRegNo", "ljmuId"];
const FROZEN_LEFT = {};
(() => {
  let acc = DRAG_HANDLE_WIDTH;
  FROZEN_KEYS.forEach((k) => { FROZEN_LEFT[k] = acc; acc += FROZEN_WIDTH[k]; });
})();
function frozenStyle(key, isHeader, rowBg) {
  if (!FROZEN_WIDTH[key]) return {};
  const isLast = key === FROZEN_KEYS[FROZEN_KEYS.length - 1];
  return {
    position: "sticky",
    left: FROZEN_LEFT[key],
    zIndex: isHeader ? 20 : 10,
    background: isHeader ? THEME.bg : (rowBg || THEME.surface),
    width: FROZEN_WIDTH[key],
    minWidth: FROZEN_WIDTH[key],
    boxShadow: isLast ? "2px 0 4px rgba(0,0,0,0.05)" : undefined,
  };
}
function dragHandleStyle(isHeader, rowBg) {
  return {
    position: "sticky",
    left: 0,
    zIndex: isHeader ? 20 : 10,
    background: isHeader ? THEME.bg : (rowBg || THEME.surface),
    width: DRAG_HANDLE_WIDTH,
    minWidth: DRAG_HANDLE_WIDTH,
  };
}

const FILTER_FIELD_KEYS = [
  "admitTerm", "programmeCode", "offerStatus", "canvasActivationStatus", "level04Status",
  "level05Status", "graduationEligibility", "graduatedYear", "activeStatus",
];
const FILTER_KEYS = FILTER_FIELD_KEYS.map((key) => ({ key, label: FIELD_BY_KEY[key].label }));

/* ----------------------------- helpers ----------------------------- */
function normalize(s) {
  return (s || "").toString().replace(/\*/g, "").replace(/[\r\n\t]+/g, " ").trim().toLowerCase().replace(/\s+/g, " ");
}

function makeId() {
  try { return crypto.randomUUID(); } catch (e) { return "id-" + Date.now() + "-" + Math.random().toString(16).slice(2); }
}

function makeEmptyRecord() {
  const rec = { id: makeId() };
  FIELDS.forEach((f) => { rec[f.key] = ""; });
  rec.activeStatus = "Active";
  return rec;
}

function getJourney(rec) {
  const dropped = !!(rec.dropReason && rec.dropReason.trim());
  const withdrawn = !!(rec.withdrawalFormStatus && rec.withdrawalFormStatus.trim());
  let stage = 0;
  if (rec.offerStatus && rec.offerStatus.trim()) stage = 1;
  if (rec.classListConfirmation && rec.classListConfirmation.trim()) stage = 2;
  if (rec.level04Status && rec.level04Status.trim()) stage = 3;
  if (rec.level05Status && rec.level05Status.trim()) stage = 4;
  if ((rec.graduationEligibility && /eligible/i.test(rec.graduationEligibility)) || (rec.graduatedYear && rec.graduatedYear.trim())) stage = 5;
  return { stage, dropped, withdrawn };
}

function getRowHighlight(rec) {
  const status = (rec.activeStatus || "").trim();
  return HIGHLIGHTS_BY_STATUS[status] || null;
}

function mapHeadersToFieldKeys(headerRow) {
  const used = new Set();
  return headerRow.map((h) => {
    const norm = normalize(h);
    if (!norm) return null;
    const match = FIELDS.find((f) => !used.has(f.key) && normalize(f.label) === norm);
    if (match) { used.add(match.key); return match.key; }
    return null;
  });
}

function recordLabel(rec) {
  if (!rec) return "";
  return rec.applicationRefNo || rec.icbtRegNo || rec.dbr || rec.nicPassport || "(unnamed record)";
}

function timeAgo(date) {
  if (!date) return "";
  const diffMs = Date.now() - new Date(date).getTime();
  const s = Math.round(diffMs / 1000);
  if (s < 10) return "just now";
  if (s < 60) return `${s}s ago`;
  const m = Math.round(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.round(m / 60);
  if (h < 24) return `${h}h ago`;
  const d = Math.round(h / 24);
  return `${d}d ago`;
}

/* ----------------------------- small UI pieces ----------------------------- */
function StatCard({ icon: Icon, label, value, tint }) {
  return (
    <div style={{ background: THEME.surface, border: `1px solid ${THEME.border}`, borderRadius: 10 }} className="px-3 py-3 flex items-center gap-3">
      <div style={{ background: tint, width: 34, height: 34, borderRadius: 8 }} className="flex items-center justify-center shrink-0">
        <Icon size={16} color={THEME.primaryDark} />
      </div>
      <div className="min-w-0">
        <div style={{ fontFamily: FONT_DISPLAY, color: THEME.ink }} className="text-lg font-semibold leading-none">{value}</div>
        <div style={{ color: THEME.inkMuted }} className="text-xs mt-1 truncate">{label}</div>
      </div>
    </div>
  );
}

function StatusPill({ value }) {
  if (!value || !value.trim()) return <span style={{ color: THEME.inkFaint }} className="text-xs">—</span>;
  const v = value.toLowerCase();
  let fg = THEME.inkMuted;
  if (/(unconditional|confirmed|^pass$|eligible$|^activated$|^issued$|^approved$|offer issued|^active$)/i.test(v) && !/not/.test(v)) { fg = THEME.success; }
  else if (/(reject|fail|not eligible|not activated|^inactive$)/i.test(v)) { fg = THEME.danger; }
  return <span style={{ color: fg }} className="text-xs font-medium whitespace-nowrap">{value}</span>;
}

function FieldInput({ field, value, onChange }) {
  const listId = `dl-${field.key}`;
  const commonStyle = {
    fontFamily: field.mono ? FONT_MONO : FONT_BODY,
    border: `1px solid ${THEME.border}`,
    background: THEME.surface,
    color: THEME.ink,
  };
  if (field.textarea) {
    return (
      <textarea
        value={value}
        onChange={(e) => onChange(field.key, e.target.value)}
        rows={3}
        style={commonStyle}
        className="w-full rounded-lg px-3 py-2 text-sm outline-none"
        onFocus={(e) => (e.target.style.borderColor = THEME.primary)}
        onBlur={(e) => (e.target.style.borderColor = THEME.border)}
      />
    );
  }
  if (field.type === "select") {
    return (
      <select
        value={value}
        onChange={(e) => onChange(field.key, e.target.value)}
        style={{ ...commonStyle, color: value ? THEME.ink : THEME.inkMuted }}
        className="w-full rounded-lg px-3 py-2 text-sm outline-none"
        onFocus={(e) => (e.target.style.borderColor = THEME.primary)}
        onBlur={(e) => (e.target.style.borderColor = THEME.border)}
      >
        <option value="">— Select —</option>
        {field.options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    );
  }
  return (
    <>
      <input
        type={field.type === "date" ? "date" : "text"}
        list={field.suggestions ? listId : undefined}
        value={value}
        onChange={(e) => onChange(field.key, e.target.value)}
        style={commonStyle}
        className="w-full rounded-lg px-3 py-2 text-sm outline-none"
        onFocus={(e) => (e.target.style.borderColor = THEME.primary)}
        onBlur={(e) => (e.target.style.borderColor = THEME.border)}
      />
      {field.suggestions && (
        <datalist id={listId}>
          {field.suggestions.map((s) => <option key={s} value={s} />)}
        </datalist>
      )}
    </>
  );
}

// Small pill in the header showing shared database sync state at a glance.
function SyncBadge({ status, lastSyncedAt, onRetry }) {
  const map = {
    idle: { icon: Cloud, text: "Not synced yet", color: THEME.inkMuted },
    pending: { icon: Cloud, text: "Changes pending…", color: THEME.inkMuted },
    saving: { icon: RefreshCw, text: "Saving to shared database…", color: THEME.primary, spin: true },
    saved: { icon: Cloud, text: lastSyncedAt ? `Saved to shared database · ${timeAgo(lastSyncedAt)}` : "Saved to shared database", color: THEME.success },
    error: { icon: AlertCircle, text: "Couldn't save — click to retry", color: THEME.danger },
  };
  const { icon: Icon, text, color, spin } = map[status] || map.idle;
  return (
    <button
      type="button"
      onClick={status === "error" ? onRetry : undefined}
      title={status === "error" ? "Retry saving to shared database" : text}
      style={{ color, cursor: status === "error" ? "pointer" : "default" }}
      className="flex items-center gap-1.5 text-xs font-medium"
    >
      <Icon size={13} className={spin ? "animate-spin" : ""} />
      <span className="hidden sm:inline">{text}</span>
    </button>
  );
}

/* ----------------------------- main app ----------------------------- */
export default function StudentRecords() {
  const { instance, accounts } = useMsal();
  const account = accounts && accounts[0];
  const displayName = account?.name || account?.username || "Signed in";

  const [students, setStudents] = useState([]);
  const [auditLog, setAuditLog] = useState([]);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState("");
  const [filters, setFilters] = useState({});
  const [showFilters, setShowFilters] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState(makeEmptyRecord());
  const [activeTab, setActiveTab] = useState(SECTIONS[0].id);
  const [toast, setToast] = useState(null);
  const [importing, setImporting] = useState(false);
  const [deleteTargetId, setDeleteTargetId] = useState(null);
  const [dragId, setDragId] = useState(null);
  const [dragOverId, setDragOverId] = useState(null);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [accountMenuOpen, setAccountMenuOpen] = useState(false);
  const [syncStatus, setSyncStatus] = useState("idle"); // idle | pending | saving | saved | error
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  const fileInputRef = useRef(null);
  const toastTimer = useRef(null);

  const showToast = useCallback((msg) => {
    setToast(msg);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 4500);
  }, []);

  /* ------------------------- load from SharePoint on sign-in ------------------------- */
  useEffect(() => {
    if (!account) return;
    let cancelled = false;
    (async () => {
      setLoading(true);
      try {
        const data = await loadFromSharePoint(instance, account);
        if (cancelled) return;
        if (data) {
          setStudents(Array.isArray(data.students) ? data.students : []);
          setAuditLog(Array.isArray(data.auditLog) ? data.auditLog : []);
          setLastSyncedAt(data.lastModified ? new Date(data.lastModified) : null);
          setSyncStatus("saved");
        } else {
          // First time this user has used the app — nothing in SharePoint yet.
          setStudents([]);
          setAuditLog([]);
          setSyncStatus("idle");
        }
      } catch (e) {
        if (!cancelled) {
          showToast("Couldn't load the shared database. Check your Microsoft 365 access and refresh.");
          setSyncStatus("error");
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [account?.homeAccountId]);

  const studentsRef = useRef(students);
  useEffect(() => { studentsRef.current = students; }, [students]);
  const auditLogRef = useRef(auditLog);
  useEffect(() => { auditLogRef.current = auditLog; }, [auditLog]);

  /* ------------------------- debounced SharePoint save ------------------------- */
  const pendingRef = useRef(null);
  const saveTimerRef = useRef(null);

  const flushSync = useCallback(async () => {
    if (!pendingRef.current || !account) return;
    const payload = pendingRef.current;
    setSyncStatus("saving");
    try {
      await saveToSharePoint(instance, account, {
        students: payload.students,
        auditLog: payload.auditLog,
        lastModified: new Date().toISOString(),
        lastModifiedBy: displayName,
      });
      pendingRef.current = null;
      setSyncStatus("saved");
      setLastSyncedAt(new Date());
    } catch (e) {
      setSyncStatus("error");
      showToast("Couldn't save to SharePoint — your changes are kept in this tab, click the sync status to retry.");
    }
  }, [instance, account, displayName, showToast]);

  const scheduleSync = useCallback((nextStudents, nextAuditLog) => {
    pendingRef.current = { students: nextStudents, auditLog: nextAuditLog };
    setSyncStatus("pending");
    clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(flushSync, 1200);
  }, [flushSync]);

  // Apply a local change immediately (snappy UI) and queue a SharePoint save.
  const persist = useCallback((nextStudents, nextAuditLog) => {
    setStudents(nextStudents);
    setAuditLog(nextAuditLog);
    scheduleSync(nextStudents, nextAuditLog);
  }, [scheduleSync]);

  function logChange(action, rec, details) {
    return {
      id: makeId(),
      timestamp: new Date().toISOString(),
      user: displayName,
      action,
      recordLabel: recordLabel(rec),
      details: details || "",
    };
  }

  const filterOptions = useMemo(() => {
    const opts = {};
    FILTER_KEYS.forEach(({ key }) => {
      opts[key] = Array.from(new Set(students.map((s) => (s[key] || "").trim()).filter(Boolean))).sort();
    });
    return opts;
  }, [students]);

  const filtered = useMemo(() => {
    const q = normalize(query);
    return students.filter((s) => {
      if (q) {
        const hit = SEARCH_KEYS.some((k) => normalize(s[k]).includes(q));
        if (!hit) return false;
      }
      for (const { key } of FILTER_KEYS) {
        if (filters[key] && (s[key] || "").trim() !== filters[key]) return false;
      }
      return true;
    });
  }, [students, query, filters]);

  const stats = useMemo(() => {
    let graduated = 0, dropped = 0, active = 0;
    students.forEach((s) => {
      const j = getJourney(s);
      if (j.dropped || j.withdrawn) dropped++;
      else if (j.stage === 5) graduated++;
      else active++;
    });
    return { total: students.length, graduated, dropped, active };
  }, [students]);

  const activeFilterCount = Object.values(filters).filter(Boolean).length;

  function openNew() {
    setFormData(makeEmptyRecord());
    setEditingId(null);
    setActiveTab(SECTIONS[0].id);
    setDrawerOpen(true);
  }

  function openEdit(rec) {
    setFormData({ ...rec });
    setEditingId(rec.id);
    setActiveTab(SECTIONS[0].id);
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  function handleFieldChange(key, val) {
    setFormData((prev) => ({ ...prev, [key]: val }));
  }

  function handleSave() {
    if (editingId) {
      const next = students.map((s) => (s.id === editingId ? { ...formData, id: editingId } : s));
      const nextLog = [logChange("Updated", formData), ...auditLog];
      persist(next, nextLog);
      showToast("Student record updated.");
    } else {
      const rec = { ...formData };
      const next = [...students, rec];
      const nextLog = [logChange("Added", rec), ...auditLog];
      persist(next, nextLog);
      showToast("Student registered.");
    }
    setDrawerOpen(false);
  }

  function requestDelete(id) {
    setDeleteTargetId(id);
  }

  function handleRowDragStart(id) {
    setDragId(id);
  }

  function handleRowDragOver(e, id) {
    e.preventDefault();
    if (id !== dragId) setDragOverId(id);
  }

  function handleRowDrop(targetId) {
    if (!dragId || dragId === targetId) { setDragId(null); setDragOverId(null); return; }
    const next = [...students];
    const fromIndex = next.findIndex((s) => s.id === dragId);
    const toIndex = next.findIndex((s) => s.id === targetId);
    if (fromIndex === -1 || toIndex === -1) { setDragId(null); setDragOverId(null); return; }
    const [moved] = next.splice(fromIndex, 1);
    next.splice(toIndex, 0, moved);
    persist(next, auditLogRef.current);
    setDragId(null);
    setDragOverId(null);
  }

  function handleRowDragEnd() {
    setDragId(null);
    setDragOverId(null);
  }

  function confirmDelete() {
    const id = deleteTargetId;
    if (!id) return;
    setDeleteTargetId(null);
    if (editingId === id) setDrawerOpen(false);
    const target = studentsRef.current.find((s) => s.id === id);
    const next = studentsRef.current.filter((s) => s.id !== id);
    const nextLog = [logChange("Deleted", target), ...auditLogRef.current];
    showToast("Student record deleted.");
    persist(next, nextLog);
  }

  function handleImportClick() {
    fileInputRef.current?.click();
  }

  async function handleImportFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImporting(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
      if (!rows.length) { showToast("That file looks empty."); setImporting(false); return; }
      const headerRow = rows[0];
      const keyMap = mapHeadersToFieldKeys(headerRow);
      const unmatched = headerRow.filter((h, i) => h && keyMap[i] === null);

      const existingByIdentity = new Map();
      students.forEach((s) => {
        const idKey = (s.applicationRefNo || s.icbtRegNo || s.nicPassport || s.dbr || "").trim().toLowerCase();
        if (idKey) existingByIdentity.set(idKey, s.id);
      });

      let next = [...students];
      let added = 0, updated = 0;
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.every((c) => !String(c).trim())) continue;
        const rec = {};
        keyMap.forEach((key, i) => {
          if (key) rec[key] = String(row[i] ?? "").trim();
        });
        if (!Object.values(rec).some((v) => v)) continue;
        const idKey = (rec.applicationRefNo || rec.icbtRegNo || rec.nicPassport || rec.dbr || "").trim().toLowerCase();
        if (idKey && existingByIdentity.has(idKey)) {
          const existingId = existingByIdentity.get(idKey);
          next = next.map((s) => (s.id === existingId ? { ...s, ...rec, id: existingId } : s));
          updated++;
        } else {
          const full = makeEmptyRecord();
          next = [...next, { ...full, ...rec }];
          added++;
          if (idKey) existingByIdentity.set(idKey, full.id);
        }
      }
      const nextLog = [
        {
          id: makeId(),
          timestamp: new Date().toISOString(),
          user: displayName,
          action: "Imported",
          recordLabel: `${added} new, ${updated} updated`,
          details: file.name,
        },
        ...auditLog,
      ];
      persist(next, nextLog);
      const unmatchedMsg = unmatched.length ? ` ${unmatched.length} column(s) not recognized: ${unmatched.slice(0, 4).join(", ")}${unmatched.length > 4 ? "…" : ""}` : "";
      showToast(`Imported ${added} new, updated ${updated}.${unmatchedMsg}`);
    } catch (err) {
      showToast("Couldn't read that file. Please check it's a valid .xlsx.");
    } finally {
      setImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  function handleExport() {
    const header = FIELDS.map((f) => f.label);
    const rows = students.map((s) => FIELDS.map((f) => s[f.key] || ""));
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Students");
    XLSX.writeFile(wb, "student-records.xlsx");
  }

  function handleSignOut() {
    instance.logoutPopup().catch(() => instance.logoutRedirect());
  }

  function renderCell(key, value) {
    if (PILL_COLUMNS.has(key)) return <StatusPill value={value} />;
    if (TRUNCATE_COLUMNS.has(key)) {
      return <span className="line-clamp-2 block max-w-[200px]" title={value || ""}>{value ? value : <span style={{ color: THEME.inkFaint }}>—</span>}</span>;
    }
    return (
      <span style={{ fontFamily: MONO_COLUMNS.has(key) ? FONT_MONO : undefined }}>
        {value ? value : <span style={{ color: THEME.inkFaint }}>—</span>}
      </span>
    );
  }

  return (
    <div
      style={{ background: THEME.bg, color: THEME.ink, fontFamily: FONT_BODY, height: "100vh" }}
      className="flex text-sm overflow-hidden"
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap');
        * { box-sizing: border-box; }
        input:focus, textarea:focus, select:focus { box-shadow: 0 0 0 3px ${THEME.primarySoft}; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${THEME.borderStrong}; border-radius: 8px; }
        .sr-table-row:hover { filter: brightness(0.97); }
        .line-clamp-2 { display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
        @media (prefers-reduced-motion: no-preference) {
          .sr-drawer { transition: transform 0.25s ease; }
        }
        .animate-spin { animation: sr-spin 1s linear infinite; }
        @keyframes sr-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
      `}</style>

      {/* sidebar */}
      <aside style={{ borderRight: `1px solid ${THEME.border}`, background: THEME.surface }} className="hidden md:flex w-64 shrink-0 flex-col p-5 gap-6 overflow-y-auto">
        <div>
          <div style={{ fontFamily: FONT_DISPLAY, color: THEME.primaryDark }} className="text-lg font-semibold tracking-tight">Students Master Database</div>
          <div style={{ color: THEME.inkFaint }} className="text-xs mt-0.5">ICBT · LJMU programme tracker</div>
        </div>
        <div className="flex flex-col gap-2.5">
          <StatCard icon={Users} label="Total students" value={stats.total} tint={THEME.primarySoft} />
          <StatCard icon={SlidersHorizontal} label="Active in pipeline" value={stats.active} tint={THEME.neutralSoft} />
          <StatCard icon={GraduationCap} label="Graduated" value={stats.graduated} tint={THEME.successSoft} />
          <StatCard icon={UserX} label="Dropped / withdrawn" value={stats.dropped} tint={THEME.dangerSoft} />
        </div>

        <div className="flex flex-col gap-1.5">
          <div style={{ color: THEME.inkFaint }} className="text-[11px] font-semibold uppercase tracking-wide mb-1">Highlight legend</div>
          {Object.values(HIGHLIGHTS).map((h) => (
            <div key={h.label} className="flex items-center gap-2">
              <span style={{ background: h.color, width: 12, height: 12, borderRadius: 4 }} className="shrink-0" />
              <span style={{ color: THEME.inkMuted }} className="text-xs">{h.label}</span>
            </div>
          ))}
        </div>

        <div className="mt-auto flex flex-col gap-2">
          <button type="button" onClick={handleImportClick}
            style={{ border: `1px solid ${THEME.border}`, color: THEME.ink }}
            className="flex items-center gap-2 text-xs font-medium px-3 py-2 rounded-lg hover:bg-gray-50"
          >
            <Upload size={14} /> Import from Excel
          </button>
          <button type="button" onClick={handleExport}
            style={{ border: `1px solid ${THEME.border}`, color: THEME.ink }}
            className="flex items-center gap-2 text-xs font-medium px-3 py-2 rounded-lg hover:bg-gray-50"
          >
            <Download size={14} /> Export to Excel
          </button>
          <button type="button" onClick={() => setHistoryOpen(true)}
            style={{ border: `1px solid ${THEME.border}`, color: THEME.ink }}
            className="flex items-center gap-2 text-xs font-medium px-3 py-2 rounded-lg hover:bg-gray-50"
          >
            <History size={14} /> Edit history
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleImportFile} className="hidden" />
          <div style={{ color: THEME.inkFaint }} className="text-[11px] text-center pt-2">Version 1.0</div>
        </div>
      </aside>

      {/* main */}
      <main className="flex-1 flex flex-col min-w-0 min-h-0">
        {/* header */}
        <header style={{ borderBottom: `1px solid ${THEME.border}`, background: THEME.surface }} className="px-4 md:px-6 py-4 flex flex-col gap-3">
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <div style={{ fontFamily: FONT_DISPLAY }} className="text-xl font-semibold md:hidden">Students Master Database</div>
            <div className="relative flex-1 min-w-[220px] max-w-md">
              <Search size={16} style={{ color: THEME.inkFaint }} className="absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Search by ref no, reg no, NIC, email, mobile…"
                style={{ border: `1px solid ${THEME.border}`, background: THEME.bg }}
                className="w-full rounded-lg pl-9 pr-3 py-2 text-sm outline-none"
              />
            </div>
            <div className="flex items-center gap-3">
              <SyncBadge status={syncStatus} lastSyncedAt={lastSyncedAt} onRetry={flushSync} />
              <button type="button" onClick={() => setShowFilters((v) => !v)}
                style={{ border: `1px solid ${THEME.border}`, color: THEME.ink, background: showFilters ? THEME.primarySoft : THEME.surface }}
                className="flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg"
              >
                <SlidersHorizontal size={14} /> Filters {activeFilterCount > 0 && <span style={{ background: THEME.primary, color: "#fff" }} className="rounded-full text-[10px] px-1.5 py-0.5 ml-0.5">{activeFilterCount}</span>}
              </button>
              <button type="button" onClick={openNew}
                style={{ background: THEME.primary, color: "#fff" }}
                className="flex items-center gap-1.5 text-xs font-semibold px-3.5 py-2 rounded-lg hover:opacity-90"
              >
                <UserPlus size={15} /> Register Student
              </button>
              <div className="relative">
                <button type="button" onClick={() => setAccountMenuOpen((v) => !v)}
                  style={{ background: THEME.primarySoft, color: THEME.primaryDark }}
                  className="w-8 h-8 rounded-full flex items-center justify-center text-xs font-semibold shrink-0"
                  title={displayName}
                >
                  {displayName.trim().charAt(0).toUpperCase() || <UserIcon size={14} />}
                </button>
                {accountMenuOpen && (
                  <>
                    <div className="fixed inset-0 z-40" onClick={() => setAccountMenuOpen(false)} />
                    <div style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }} className="absolute right-0 mt-2 w-52 rounded-lg shadow-lg z-50 py-1.5">
                      <div style={{ borderBottom: `1px solid ${THEME.border}` }} className="px-3 py-2">
                        <div className="text-xs font-semibold truncate">{displayName}</div>
                        <div style={{ color: THEME.inkFaint }} className="text-[11px] truncate">{account?.username}</div>
                      </div>
                      <button type="button" onClick={handleSignOut}
                        style={{ color: THEME.danger }}
                        className="w-full flex items-center gap-2 text-xs font-medium px-3 py-2 hover:bg-gray-50"
                      >
                        <LogOut size={13} /> Sign out
                      </button>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>

          {showFilters && (
            <div className="flex flex-col gap-2 pt-1">
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-2">
                {FILTER_KEYS.map(({ key, label }) => (
                  <div key={key} className="relative">
                    <select
                      value={filters[key] || ""}
                      onChange={(e) => setFilters((f) => ({ ...f, [key]: e.target.value }))}
                      style={{ border: `1px solid ${THEME.border}`, background: THEME.bg, color: filters[key] ? THEME.ink : THEME.inkMuted }}
                      className="w-full appearance-none text-xs rounded-lg pl-3 pr-7 py-2 outline-none truncate"
                    >
                      <option value="">{label}: All</option>
                      {(filterOptions[key] || []).map((v) => <option key={v} value={v}>{v}</option>)}
                    </select>
                    <ChevronDown size={12} style={{ color: THEME.inkFaint }} className="absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  </div>
                ))}
              </div>
              <div className="flex items-center gap-3">
                {activeFilterCount > 0 && (
                  <button type="button" onClick={() => setFilters({})} style={{ color: THEME.primary }} className="text-xs font-medium hover:underline">
                    Clear filters
                  </button>
                )}
                <span style={{ color: THEME.inkFaint }} className="text-xs ml-auto">{filtered.length} of {students.length} students</span>
              </div>
            </div>
          )}
          <div className="md:hidden flex gap-2">
            <button type="button" onClick={handleImportClick} style={{ border: `1px solid ${THEME.border}` }} className="flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg">
              <Upload size={13} /> Import
            </button>
            <button type="button" onClick={handleExport} style={{ border: `1px solid ${THEME.border}` }} className="flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg">
              <Download size={13} /> Export
            </button>
            <button type="button" onClick={() => setHistoryOpen(true)} style={{ border: `1px solid ${THEME.border}` }} className="flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg">
              <History size={13} /> History
            </button>
          </div>
        </header>

        {/* content */}
        <div className="flex-1 min-h-0 p-4 md:p-6 flex flex-col">
          {loading || importing ? (
            <div style={{ color: THEME.inkMuted }} className="flex items-center justify-center h-64 text-sm">
              {importing ? "Importing your file…" : "Loading the shared student database…"}
            </div>
          ) : students.length === 0 ? (
            <div style={{ background: THEME.surface, border: `1px dashed ${THEME.borderStrong}` }} className="rounded-xl flex flex-col items-center justify-center text-center py-20 px-6">
              <div style={{ background: THEME.primarySoft }} className="w-12 h-12 rounded-full flex items-center justify-center mb-4">
                <FileSpreadsheet size={20} color={THEME.primaryDark} />
              </div>
              <div style={{ fontFamily: FONT_DISPLAY }} className="text-base font-semibold">No students yet</div>
              <div style={{ color: THEME.inkMuted }} className="text-sm mt-1 max-w-sm">Import your existing Excel master sheet, or register a student manually to get started.</div>
              <div className="flex gap-2 mt-5">
                <button type="button" onClick={handleImportClick} style={{ background: THEME.primary, color: "#fff" }} className="flex items-center gap-1.5 text-xs font-semibold px-4 py-2.5 rounded-lg">
                  <Upload size={14} /> Import from Excel
                </button>
                <button type="button" onClick={openNew} style={{ border: `1px solid ${THEME.border}` }} className="flex items-center gap-1.5 text-xs font-semibold px-4 py-2.5 rounded-lg">
                  <UserPlus size={14} /> Register manually
                </button>
              </div>
            </div>
          ) : filtered.length === 0 ? (
            <div style={{ color: THEME.inkMuted }} className="flex flex-col items-center justify-center h-64 text-sm gap-2">
              <div>No students match your search or filters.</div>
              <button type="button" onClick={() => { setQuery(""); setFilters({}); }} style={{ color: THEME.primary }} className="text-xs font-medium hover:underline">Clear search & filters</button>
            </div>
          ) : (
            <div style={{ border: `1px solid ${THEME.border}`, background: THEME.surface }} className="rounded-xl overflow-hidden flex-1 min-h-0 flex flex-col">
              <div className="flex-1 min-h-0 overflow-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr style={{ borderBottom: `1px solid ${THEME.border}` }}>
                      <th style={{ top: 0, position: "sticky", zIndex: 15, ...dragHandleStyle(true) }} className="px-2 py-3" />
                      {TABLE_COLUMNS.map((key) => (
                        <th key={key} style={{ color: THEME.inkMuted, background: THEME.bg, top: 0, position: "sticky", zIndex: 15, ...frozenStyle(key, true) }} className="text-xs font-semibold uppercase tracking-wide px-4 py-3 whitespace-nowrap">{FIELD_BY_KEY[key].label.split(" (")[0]}</th>
                      ))}
                      <th style={{ background: THEME.bg, top: 0, position: "sticky", zIndex: 15 }} className="px-4 py-3" />
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((s) => {
                      const highlight = getRowHighlight(s);
                      const isDragOver = dragOverId === s.id && dragId !== s.id;
                      return (
                        <tr
                          key={s.id}
                          className="sr-table-row"
                          onDragOver={(e) => handleRowDragOver(e, s.id)}
                          onDrop={() => handleRowDrop(s.id)}
                          onDragEnd={handleRowDragEnd}
                          style={{
                            borderBottom: `1px solid ${THEME.border}`,
                            background: highlight ? highlight.color : undefined,
                            boxShadow: isDragOver ? `inset 0 2px 0 ${THEME.primary}` : undefined,
                            opacity: dragId === s.id ? 0.5 : 1,
                          }}
                        >
                          <td className="px-2 py-3" style={dragHandleStyle(false, highlight ? highlight.color : undefined)}>
                            <span
                              draggable
                              onDragStart={() => handleRowDragStart(s.id)}
                              style={{ color: THEME.inkFaint, cursor: "grab" }}
                              className="flex items-center justify-center"
                              title="Drag to reorder"
                            >
                              <GripVertical size={14} />
                            </span>
                          </td>
                          {TABLE_COLUMNS.map((key) => (
                            <td key={key} style={frozenStyle(key, false, highlight ? highlight.color : undefined)} className="px-4 py-3 text-xs">{renderCell(key, s[key])}</td>
                          ))}
                          <td className="px-4 py-3" style={{ background: highlight ? highlight.color : undefined }}>
                            <div className="flex items-center gap-1 justify-end">
                              <button type="button" onClick={() => openEdit(s)} style={{ color: THEME.inkMuted }} className="p-1.5 rounded hover:bg-black/10" title="Edit">
                                <Pencil size={14} />
                              </button>
                              <button type="button" onClick={() => requestDelete(s.id)} style={{ color: THEME.danger }} className="p-1.5 rounded hover:bg-black/10" title="Delete">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* drawer */}
      {drawerOpen && (
        <>
          <div onClick={closeDrawer} className="fixed inset-0 z-40" style={{ background: "rgba(24,34,48,0.35)" }} />
          <div className="sr-drawer fixed right-0 top-0 h-full z-50 w-full sm:w-[600px] flex flex-col" style={{ background: THEME.surface, boxShadow: "-8px 0 24px rgba(0,0,0,0.08)" }}>
            <div style={{ borderBottom: `1px solid ${THEME.border}` }} className="px-5 py-4 flex items-center justify-between">
              <div>
                <div style={{ fontFamily: FONT_DISPLAY }} className="text-base font-semibold">{editingId ? "Edit Student Record" : "New Student Registration"}</div>
                {editingId && <div style={{ fontFamily: FONT_MONO, color: THEME.inkFaint }} className="text-xs mt-0.5">{formData.applicationRefNo || formData.icbtRegNo || editingId}</div>}
              </div>
              <button type="button" onClick={closeDrawer} style={{ color: THEME.inkMuted }} className="p-1.5 rounded hover:bg-gray-100"><X size={18} /></button>
            </div>

            <div style={{ borderBottom: `1px solid ${THEME.border}` }} className="px-5 flex gap-1 overflow-x-auto shrink-0">
              {SECTIONS.map((sec) => (
                <button type="button"
                  key={sec.id}
                  onClick={() => setActiveTab(sec.id)}
                  style={{
                    color: activeTab === sec.id ? THEME.primary : THEME.inkMuted,
                    borderBottom: activeTab === sec.id ? `2px solid ${THEME.primary}` : "2px solid transparent",
                  }}
                  className="text-xs font-medium px-2.5 py-3 whitespace-nowrap"
                >
                  {sec.label}
                </button>
              ))}
            </div>

            <div className="flex-1 overflow-y-auto px-5 py-5">
              {(() => {
                const tabFields = FIELDS.filter((f) => f.section === activeTab);
                const subOrder = [];
                tabFields.forEach((f) => {
                  const sub = f.subsection || "__none";
                  if (!subOrder.includes(sub)) subOrder.push(sub);
                });
                return subOrder.map((sub, idx) => (
                  <div key={sub} className={idx > 0 ? "mt-6 pt-5" : ""} style={idx > 0 ? { borderTop: `1px solid ${THEME.border}` } : undefined}>
                    {sub !== "__none" && (
                      <div style={{ fontFamily: FONT_DISPLAY, color: THEME.primaryDark }} className="text-xs font-semibold uppercase tracking-wide mb-3">{sub}</div>
                    )}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {tabFields.filter((f) => (f.subsection || "__none") === sub).map((f) => (
                        <div key={f.key} className={f.textarea ? "sm:col-span-2" : ""}>
                          <label style={{ color: THEME.inkMuted }} className="text-xs font-medium block mb-1.5">{f.label}</label>
                          <FieldInput field={f} value={formData[f.key] || ""} onChange={handleFieldChange} />
                        </div>
                      ))}
                    </div>
                  </div>
                ));
              })()}
            </div>

            <div style={{ borderTop: `1px solid ${THEME.border}` }} className="px-5 py-4 flex items-center justify-between gap-3">
              {editingId ? (
                <button type="button" onClick={() => requestDelete(editingId)} style={{ color: THEME.danger }} className="text-xs font-medium flex items-center gap-1.5 hover:underline">
                  <Trash2 size={14} /> Delete record
                </button>
              ) : <span />}
              <div className="flex items-center gap-2">
                <button type="button" onClick={closeDrawer} style={{ border: `1px solid ${THEME.border}`, color: THEME.ink }} className="text-xs font-semibold px-4 py-2.5 rounded-lg">Cancel</button>
                <button type="button" onClick={handleSave} style={{ background: THEME.primary, color: "#fff" }} className="flex items-center gap-1.5 text-xs font-semibold px-4 py-2.5 rounded-lg hover:opacity-90">
                  Save <ChevronRight size={14} />
                </button>
              </div>
            </div>
          </div>
        </>
      )}

      {/* edit history drawer */}
      {historyOpen && (
        <>
          <div onClick={() => setHistoryOpen(false)} className="fixed inset-0 z-40" style={{ background: "rgba(24,34,48,0.35)" }} />
          <div className="sr-drawer fixed right-0 top-0 h-full z-50 w-full sm:w-[420px] flex flex-col" style={{ background: THEME.surface, boxShadow: "-8px 0 24px rgba(0,0,0,0.08)" }}>
            <div style={{ borderBottom: `1px solid ${THEME.border}` }} className="px-5 py-4 flex items-center justify-between">
              <div>
                <div style={{ fontFamily: FONT_DISPLAY }} className="text-base font-semibold flex items-center gap-2"><History size={16} /> Edit History</div>
                <div style={{ color: THEME.inkFaint }} className="text-xs mt-0.5">Every add, edit and delete, most recent first.</div>
              </div>
              <button type="button" onClick={() => setHistoryOpen(false)} style={{ color: THEME.inkMuted }} className="p-1.5 rounded hover:bg-gray-100"><X size={18} /></button>
            </div>
            <div className="flex-1 overflow-y-auto px-5 py-4">
              {auditLog.length === 0 ? (
                <div style={{ color: THEME.inkFaint }} className="text-xs text-center py-10">No changes logged yet.</div>
              ) : (
                <div className="flex flex-col gap-3">
                  {auditLog.map((entry) => {
                    const actionColor =
                      entry.action === "Deleted" ? THEME.danger :
                      entry.action === "Added" ? THEME.success :
                      THEME.primary;
                    return (
                      <div key={entry.id} style={{ borderBottom: `1px solid ${THEME.border}` }} className="pb-3">
                        <div className="flex items-center justify-between gap-2">
                          <span style={{ color: actionColor }} className="text-xs font-semibold">{entry.action}</span>
                          <span style={{ color: THEME.inkFaint }} className="text-[11px]">{timeAgo(entry.timestamp)}</span>
                        </div>
                        <div className="text-xs mt-0.5">{entry.recordLabel}</div>
                        <div style={{ color: THEME.inkMuted }} className="text-[11px] mt-0.5">
                          {entry.user}{entry.details ? ` · ${entry.details}` : ""}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* delete confirmation */}
      {deleteTargetId && (() => {
        const target = students.find((s) => s.id === deleteTargetId);
        return (
          <div
            onClick={() => setDeleteTargetId(null)}
            className="fixed inset-0 z-[70] flex items-center justify-center px-4"
            style={{ background: "rgba(24,34,48,0.45)" }}
          >
            <div onClick={(e) => e.stopPropagation()} style={{ background: THEME.surface }} className="rounded-xl w-full max-w-sm p-5 shadow-xl">
              <div style={{ background: THEME.dangerSoft }} className="w-10 h-10 rounded-full flex items-center justify-center mb-3">
                <Trash2 size={18} color={THEME.danger} />
              </div>
              <div style={{ fontFamily: FONT_DISPLAY }} className="text-base font-semibold">Delete this student record?</div>
              <div style={{ color: THEME.inkMuted }} className="text-xs mt-1.5">
                {target ? (target.applicationRefNo || target.icbtRegNo || target.dbr || "This record") : "This record"} will be permanently removed. This action can't be undone.
              </div>
              <div className="flex items-center justify-end gap-2 mt-5">
                <button type="button" onClick={() => setDeleteTargetId(null)} style={{ border: `1px solid ${THEME.border}`, color: THEME.ink }} className="text-xs font-semibold px-4 py-2 rounded-lg">Cancel</button>
                <button type="button" onClick={confirmDelete} style={{ background: THEME.danger, color: "#fff" }} className="text-xs font-semibold px-4 py-2 rounded-lg hover:opacity-90">Delete</button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* toast */}
      {toast && (
        <div
          style={{ background: THEME.ink, color: "#fff" }}
          className="fixed bottom-5 left-1/2 -translate-x-1/2 z-[60] text-xs font-medium px-4 py-2.5 rounded-lg shadow-lg max-w-[90vw] text-center"
        >
          {toast}
        </div>
      )}
    </div>
  );
}
