import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { auth, db, googleProvider } from "./firebase";
import { signInWithEmailAndPassword, createUserWithEmailAndPassword, signInWithPopup, signOut, onAuthStateChanged } from "firebase/auth";
import { collection, doc, onSnapshot, addDoc, updateDoc, deleteDoc, setDoc, getDoc, runTransaction, getDocs, writeBatch } from "firebase/firestore";

// ─── DESIGN TOKENS ───────────────────────────────────────────────────────────
const G = {
  green: "#1a7f4b", greenL: "#22a864", greenXL: "#d6f5e3", greenDim: "#0e4f2f",
  red: "#d93025", amber: "#f59e0b", blue: "#1a73e8", purple: "#8b5cf6",
};

// ─── THEMES ──────────────────────────────────────────────────────────────────
const lightTheme = {
  bg:"#f8fafb", surface:"#ffffff", surface2:"#f1f3f4", border:"#e0e0e0",
  text:"#202124", textSec:"#5f6368", textMut:"#9aa0a6",
  accent:G.green, accentL:G.greenL, accentBg:G.greenXL,
  shadow:"0 1px 3px rgba(0,0,0,.12), 0 1px 2px rgba(0,0,0,.08)",
  shadowMd:"0 4px 16px rgba(0,0,0,.1)", card:"#ffffff",
};
const darkTheme = {
  bg:"#0f1117", surface:"#1a1d27", surface2:"#252836", border:"#2e3244",
  text:"#f0f1f5", textSec:"#9ca3af", textMut:"#6b7280",
  accent:G.greenL, accentL:"#4ade80", accentBg:"#064e3b",
  shadow:"0 1px 3px rgba(0,0,0,.5), 0 1px 2px rgba(0,0,0,.4)",
  shadowMd:"0 4px 16px rgba(0,0,0,.5)", card:"#1a1d27",
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function genId() { return `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`; }
function pct(f, l) { return l > 0 ? Math.round((f / l) * 100) : 0; }
function validEmail(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); }
function fmtDate(d) { if (!d) return "—"; const dt = new Date(d); return dt.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric", hour:"2-digit", minute:"2-digit" }); }

function getPoolStatus(pool) {
  if (!pool.startTime || !pool.endTime) return "Upcoming";
  const now = Date.now(), s = new Date(pool.startTime).getTime(), e = new Date(pool.endTime).getTime();
  if (now < s) return "Upcoming";
  if (now > e) return "Completed";
  return "In Progress";
}
function statusColor(s, t) {
  if (s === "In Progress") return { bg: "#fef3c7", text: "#92400e", dot: G.amber };
  if (s === "Upcoming") return { bg: t.accentBg, text: t.accent, dot: t.accent };
  return { bg: t.surface2, text: t.textSec, dot: t.textMut };
}

// ─── UI PRIMITIVES ───────────────────────────────────────────────────────────
function Card({ children, t, style = {}, onClick }) {
  return (
    <div onClick={onClick} style={{
      background: t.card, border: `1px solid ${t.border}`, borderRadius: 16,
      boxShadow: t.shadow, padding: 24, transition: "box-shadow .2s, transform .15s",
      cursor: onClick ? "pointer" : "default", ...style,
    }}
    onMouseEnter={e => { if (onClick) { e.currentTarget.style.boxShadow = t.shadowMd; e.currentTarget.style.transform = "translateY(-2px)"; } }}
    onMouseLeave={e => { e.currentTarget.style.boxShadow = t.shadow; e.currentTarget.style.transform = ""; }}
    >{children}</div>
  );
}

function Btn({ children, variant = "primary", t, onClick, style = {}, size = "md", disabled = false }) {
  const base = {
    border: "none", borderRadius: 8, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer",
    transition: "all .15s", fontFamily: "inherit", display: "inline-flex", alignItems: "center", gap: 6,
    opacity: disabled ? 0.55 : 1, fontSize: size === "sm" ? 13 : 14,
    padding: size === "sm" ? "6px 14px" : "10px 20px", whiteSpace: "nowrap",
  };
  const styles = {
    primary: { ...base, background: t.accent, color: "#fff" },
    secondary: { ...base, background: t.surface2, color: t.text, border: `1px solid ${t.border}` },
    danger: { ...base, background: G.red, color: "#fff" },
    ghost: { ...base, background: "transparent", color: t.accent },
  };
  return (
    <button disabled={disabled} onClick={onClick} style={{ ...styles[variant] || styles.primary, ...style }}
      {...(!disabled && {
        onMouseEnter: e => { e.currentTarget.style.filter = "brightness(1.15)"; e.currentTarget.style.transform = "translateY(-1px)"; },
        onMouseLeave: e => { e.currentTarget.style.filter = ""; e.currentTarget.style.transform = ""; },
      })}>{children}</button>
  );
}

function Input({ label, type = "text", value, onChange, placeholder, t, style = {}, error }) {
  return (
    <div style={{ marginBottom: 16, ...style }}>
      {label && <label style={{ display: "block", fontSize: 13, fontWeight: 500, color: t.textSec, marginBottom: 6 }}>{label}</label>}
      <input type={type} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder}
        style={{
          width: "100%", boxSizing: "border-box", padding: "10px 14px", borderRadius: 8,
          border: `1px solid ${error ? G.red : t.border}`, background: t.surface2, color: t.text,
          fontSize: 14, outline: "none", fontFamily: "inherit", transition: "border-color .15s",
        }}
        onFocus={e => e.target.style.borderColor = t.accent}
        onBlur={e => e.target.style.borderColor = error ? G.red : t.border}
      />
      {error && <div style={{ fontSize: 12, color: G.red, marginTop: 4 }}>{error}</div>}
    </div>
  );
}

function Badge({ children, color, bg }) {
  return <span style={{ fontSize: 11, fontWeight: 600, padding: "3px 9px", borderRadius: 20, background: bg, color }}>{children}</span>;
}
function ProgressBar({ value, t }) {
  const c = value >= 100 ? G.red : value >= 80 ? G.amber : t.accent;
  return (<div style={{ background: t.surface2, borderRadius: 99, height: 6, overflow: "hidden" }}>
    <div style={{ width: `${Math.min(value, 100)}%`, background: c, height: "100%", borderRadius: 99, transition: "width .4s" }} />
  </div>);
}

function StatusBadge({ pool, t }) {
  const s = getPoolStatus(pool);
  const sc = statusColor(s, t);
  return (
    <Badge color={sc.text} bg={sc.bg}>
      <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
        <span style={{ width: 6, height: 6, borderRadius: "50%", background: sc.dot, display: "inline-block" }} />
        {s}
      </span>
    </Badge>
  );
}

// ─── ICON ────────────────────────────────────────────────────────────────────
function Icon({ name, size = 18, color }) {
  const c = color || "currentColor";
  const p = { width: size, height: size, viewBox: "0 0 24 24", fill: "none", stroke: c, strokeWidth: "2", strokeLinecap: "round", strokeLinejoin: "round" };
  const icons = {
    sun: <svg {...p}><circle cx="12" cy="12" r="5" /><line x1="12" y1="1" x2="12" y2="3" /><line x1="12" y1="21" x2="12" y2="23" /><line x1="4.22" y1="4.22" x2="5.64" y2="5.64" /><line x1="18.36" y1="18.36" x2="19.78" y2="19.78" /><line x1="1" y1="12" x2="3" y2="12" /><line x1="21" y1="12" x2="23" y2="12" /><line x1="4.22" y1="19.78" x2="5.64" y2="18.36" /><line x1="18.36" y1="5.64" x2="19.78" y2="4.22" /></svg>,
    moon: <svg {...p}><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" /></svg>,
    plus: <svg {...p}><line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" /></svg>,
    search: <svg {...p}><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>,
    users: <svg {...p}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 0 0-3-3.87" /><path d="M16 3.13a4 4 0 0 1 0 7.75" /></svg>,
    user: <svg {...p}><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" /><circle cx="12" cy="7" r="4" /></svg>,
    chart: <svg {...p}><line x1="18" y1="20" x2="18" y2="10" /><line x1="12" y1="20" x2="12" y2="4" /><line x1="6" y1="20" x2="6" y2="14" /></svg>,
    settings: <svg {...p}><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" /></svg>,
    home: <svg {...p}><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" /><polyline points="9 22 9 12 15 12 15 22" /></svg>,
    back: <svg {...p}><polyline points="15 18 9 12 15 6" /></svg>,
    check: <svg {...p} strokeWidth="2.5"><polyline points="20 6 9 17 4 12" /></svg>,
    edit: <svg {...p}><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" /></svg>,
    trash: <svg {...p}><polyline points="3 6 5 6 21 6" /><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6" /><path d="M10 11v6" /><path d="M14 11v6" /><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" /></svg>,
    export: <svg {...p}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>,
    upload: <svg {...p}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>,
    bell: <svg {...p}><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9" /><path d="M13.73 21a2 2 0 0 1-3.46 0" /></svg>,
    pool: <svg {...p}><rect x="2" y="3" width="20" height="14" rx="2" /><path d="M8 21h8" /><path d="M12 17v4" /></svg>,
    lock: <svg {...p}><rect x="3" y="11" width="18" height="11" rx="2" /><path d="M7 11V7a5 5 0 0 1 10 0v4" /></svg>,
    close: <svg {...p}><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>,
    logout: <svg {...p}><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><polyline points="16 17 21 12 16 7" /><line x1="21" y1="12" x2="9" y2="12" /></svg>,
    download: <svg {...p}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>,
    key: <svg {...p}><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4" /></svg>,
    clock: <svg {...p}><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></svg>,
    google: <svg width={size} height={size} viewBox="0 0 24 24"><path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" /><path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" /><path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" /><path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" /></svg>,
    file: <svg {...p}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg>,
  };
  return icons[name] || null;
}

// ─── TOAST SYSTEM ────────────────────────────────────────────────────────────
function ToastContainer({ toasts }) {
  return (
    <div style={{ position: "fixed", top: 72, right: 20, zIndex: 9999, display: "flex", flexDirection: "column", gap: 8, maxWidth: 360 }}>
      {toasts.map(toast => (
        <div key={toast.id} style={{
          padding: "12px 20px", borderRadius: 12,
          background: toast.type === "error" ? G.red : toast.type === "warning" ? G.amber : G.green,
          color: "#fff", fontSize: 14, fontWeight: 500,
          boxShadow: "0 8px 24px rgba(0,0,0,.3)",
          display: "flex", alignItems: "center", gap: 8,
          animation: "slideIn .3s ease",
        }}>
          <Icon name={toast.type === "error" ? "close" : "check"} size={16} color="#fff" />
          {toast.message}
        </div>
      ))}
    </div>
  );
}

// ─── CONFIRM MODAL ───────────────────────────────────────────────────────────
function ConfirmModal({ t, title, message, onConfirm, onCancel, danger = true }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500, padding: 20 }} onClick={onCancel}>
      <Card t={t} style={{ width: "100%", maxWidth: 400 }} onClick={e => e.stopPropagation()}>
        <h3 style={{ margin: "0 0 8px", color: t.text, fontSize: 18 }}>{title}</h3>
        <p style={{ color: t.textSec, fontSize: 14, marginBottom: 24, lineHeight: 1.5 }}>{message}</p>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <Btn t={t} variant="secondary" onClick={onCancel}>Cancel</Btn>
          <Btn t={t} variant={danger ? "danger" : "primary"} onClick={onConfirm}>Confirm</Btn>
        </div>
      </Card>
    </div>
  );
}

// ─── EXPORT HELPERS ──────────────────────────────────────────────────────────
function exportExcel(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, filename);
}
function exportCSV(data, filename) {
  if (!data.length) return;
  const headers = Object.keys(data[0]).join(",");
  const rows = data.map(r => Object.values(r).map(v => `"${v}"`).join(",")).join("\n");
  const blob = new Blob([headers + "\n" + rows], { type: "text/csv" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename; a.click();
}
function exportPDF(data, title) {
  const w = window.open("", "_blank");
  w.document.write(`<html><head><title>${title}</title><style>body{font-family:Arial,sans-serif;padding:24px}table{width:100%;border-collapse:collapse;margin-top:16px}th,td{border:1px solid #ddd;padding:10px;text-align:left;font-size:13px}th{background:#1a7f4b;color:#fff}tr:nth-child(even){background:#f9f9f9}h2{color:#1a7f4b;margin:0}p{color:#888;font-size:12px}</style></head><body><h2>${title}</h2><table><tr>${Object.keys(data[0] || {}).map(k => `<th>${k}</th>`).join("")}</tr>${data.map(r => `<tr>${Object.values(r).map(v => `<td>${v}</td>`).join("")}</tr>`).join("")}</table><p>Generated on ${new Date().toLocaleString()}</p></body></html>`);
  w.document.close(); setTimeout(() => w.print(), 300);
}
function downloadTemplate() {
  const ws = XLSX.utils.aoa_to_sheet([["Name", "Email", "Register Number", "Mobile Number"], ["John Doe", "john@example.com", "CS001", "9876543210"]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Members");
  ws["!cols"] = [{ wch: 20 }, { wch: 25 }, { wch: 18 }, { wch: 15 }];
  XLSX.writeFile(wb, "members_template.xlsx");
}

// ─── NAV BAR ─────────────────────────────────────────────────────────────────
function NavBar({ t, dark, setDark, page, setPage, user, notifications, onMarkRead, onLogout }) {
  const [showProfile, setShowProfile] = useState(false);
  const [showNotif, setShowNotif] = useState(false);
  const unread = notifications.filter(n => !n.read).length;
  const profRef = useRef(null);
  const notifRef = useRef(null);

  useEffect(() => {
    function handleClick(e) {
      if (profRef.current && !profRef.current.contains(e.target)) setShowProfile(false);
      if (notifRef.current && !notifRef.current.contains(e.target)) setShowNotif(false);
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  return (
    <header style={{
      position: "sticky", top: 0, zIndex: 100, background: t.surface,
      borderBottom: `1px solid ${t.border}`, display: "flex", alignItems: "center",
      justifyContent: "space-between", padding: "0 24px", height: 60, boxShadow: t.shadow,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }} onClick={() => setPage("home")}>
        <div style={{ width: 32, height: 32, background: `linear-gradient(135deg,${G.green},${G.greenL})`, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <span style={{ color: "#fff", fontSize: 14, fontWeight: 800 }}>A</span>
        </div>
        <span style={{ fontWeight: 700, fontSize: 17, color: t.text }} className="hide-sm">Allotment</span>
      </div>

      <nav style={{ display: "flex", gap: 4 }}>
        {[{ id: "home", icon: "home", label: "Groups" }, { id: "stats", icon: "chart", label: "Stats" }, { id: "settings", icon: "settings", label: "Settings" }].map(n => (
          <button key={n.id} onClick={() => setPage(n.id)} style={{
            display: "flex", alignItems: "center", gap: 6, padding: "6px 14px",
            borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
            background: page === n.id ? t.accentBg : "transparent",
            color: page === n.id ? t.accent : t.textSec, transition: "all .15s",
          }}>
            <Icon name={n.icon} size={16} color={page === n.id ? t.accent : t.textSec} />
            <span className="hide-sm">{n.label}</span>
          </button>
        ))}
      </nav>

      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <button onClick={() => setDark(!dark)} style={{
          background: t.surface2, border: `1px solid ${t.border}`, borderRadius: 8,
          padding: "6px 10px", cursor: "pointer", display: "flex", alignItems: "center",
        }}>
          <Icon name={dark ? "sun" : "moon"} size={16} color={t.textSec} />
        </button>

        {/* Notifications */}
        <div ref={notifRef} style={{ position: "relative" }}>
          <button onClick={() => { setShowNotif(!showNotif); if (!showNotif) onMarkRead(); }} style={{
            background: t.surface2, border: `1px solid ${t.border}`, borderRadius: 8,
            padding: "6px 10px", cursor: "pointer", display: "flex", position: "relative",
          }}>
            <Icon name="bell" size={16} color={t.textSec} />
            {unread > 0 && <span style={{ position: "absolute", top: -4, right: -4, width: 16, height: 16, borderRadius: "50%", background: G.red, color: "#fff", fontSize: 9, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center" }}>{unread}</span>}
          </button>
          {showNotif && (
            <div style={{
              position: "absolute", top: 44, right: 0, width: 300, maxHeight: 320, overflowY: "auto",
              background: t.surface, border: `1px solid ${t.border}`, borderRadius: 12,
              boxShadow: t.shadowMd, zIndex: 200,
            }}>
              <div style={{ padding: "12px 16px", borderBottom: `1px solid ${t.border}`, fontWeight: 700, fontSize: 14, color: t.text }}>Notifications</div>
              {notifications.length === 0 ? (
                <div style={{ padding: "24px 16px", textAlign: "center", color: t.textMut, fontSize: 13 }}>No notifications</div>
              ) : notifications.slice(0, 10).map(n => (
                <div key={n.id} style={{ padding: "10px 16px", borderBottom: `1px solid ${t.border}`, fontSize: 13, color: t.text }}>
                  <div>{n.message}</div>
                  <div style={{ fontSize: 11, color: t.textMut, marginTop: 2 }}>{new Date(n.time).toLocaleTimeString()}</div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Profile */}
        <div ref={profRef} style={{ position: "relative" }}>
          <div onClick={() => setShowProfile(!showProfile)} style={{
            width: 34, height: 34, borderRadius: "50%",
            background: `linear-gradient(135deg,${G.green},${G.greenL})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer",
          }}>
            {user?.name?.[0]?.toUpperCase() || "U"}
          </div>
          {showProfile && (
            <div style={{
              position: "absolute", top: 44, right: 0, width: 220, background: t.surface,
              border: `1px solid ${t.border}`, borderRadius: 12, boxShadow: t.shadowMd, zIndex: 200, overflow: "hidden",
            }}>
              <div style={{ padding: "14px 16px", borderBottom: `1px solid ${t.border}` }}>
                <div style={{ fontWeight: 600, fontSize: 14, color: t.text }}>{user?.name}</div>
                <div style={{ fontSize: 12, color: t.textMut }}>{user?.email}</div>
                <Badge color={user?.role === "admin" ? "#92400e" : t.accent} bg={user?.role === "admin" ? "#fef3c7" : t.accentBg}>{user?.role}</Badge>
              </div>
              {[{ icon: "user", label: "Profile", action: () => { setPage("profile"); setShowProfile(false); } },
                { icon: "logout", label: "Logout", action: () => { onLogout(); setShowProfile(false); } },
              ].map(item => (
                <button key={item.label} onClick={item.action} style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "10px 16px", width: "100%",
                  border: "none", background: "transparent", cursor: "pointer", color: item.label === "Logout" ? G.red : t.text,
                  fontSize: 14, transition: "background .1s", textAlign: "left",
                }}
                  onMouseEnter={e => e.currentTarget.style.background = t.surface2}
                  onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                >
                  <Icon name={item.icon} size={16} color={item.label === "Logout" ? G.red : t.textSec} />{item.label}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
    </header>
  );
}

// ─── AUTH PAGE ────────────────────────────────────────────────────────────────
function AuthPage({ t }) {
  const [tab, setTab] = useState("login");
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [name, setName] = useState("");
  const [errors, setErrors] = useState({});
  const [loading, setLoading] = useState(false);

  async function handleSubmit() {
    const errs = {};
    if (tab === "register" && !name.trim()) errs.name = "Name is required";
    if (!email.trim()) errs.email = "Email is required";
    else if (!validEmail(email)) errs.email = "Invalid email format";
    if (!pass.trim()) errs.pass = "Password is required";
    else if (pass.length < 6) errs.pass = "Min 6 characters";
    setErrors(errs);
    if (Object.keys(errs).length > 0) return;

    setLoading(true);
    try {
      if (tab === "register") {
        const cred = await createUserWithEmailAndPassword(auth, email, pass);
        await setDoc(doc(db, "users", cred.user.uid), { name, email, role: "user", phone: "", regNo: "" });
      } else {
        await signInWithEmailAndPassword(auth, email, pass);
      }
    } catch (error) {
      setErrors({ form: error.message });
    }
    setLoading(false);
  }

  async function handleGoogle() {
    try {
      const cred = await signInWithPopup(auth, googleProvider);
      const userDoc = await getDoc(doc(db, "users", cred.user.uid));
      if (!userDoc.exists()) {
        await setDoc(doc(db, "users", cred.user.uid), { name: cred.user.displayName, email: cred.user.email, role: "user", phone: "", regNo: "" });
      }
    } catch (error) {
      setErrors({ form: error.message });
    }
  }

  const at = { ...t, surface2: "rgba(255,255,255,.08)", border: "rgba(255,255,255,.15)", text: "#fff", textSec: "rgba(255,255,255,.6)", accent: G.greenL };

  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(135deg,${G.greenDim} 0%,#1a1a2e 60%,#0f0f1a 100%)`, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ position: "fixed", top: -80, right: -80, width: 300, height: 300, borderRadius: "50%", background: "rgba(34,168,100,.1)", pointerEvents: "none" }} />
      <div style={{ position: "fixed", bottom: -60, left: -60, width: 200, height: 200, borderRadius: "50%", background: "rgba(34,168,100,.07)", pointerEvents: "none" }} />
      <div style={{
        width: "100%", maxWidth: 400, background: "rgba(255,255,255,.06)", backdropFilter: "blur(20px)",
        border: "1px solid rgba(255,255,255,.12)", borderRadius: 24, padding: 40,
        boxShadow: "0 24px 64px rgba(0,0,0,.4)",
      }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{
            width: 56, height: 56, borderRadius: 16, margin: "0 auto 14px",
            background: `linear-gradient(135deg,${G.green},${G.greenL})`,
            display: "flex", alignItems: "center", justifyContent: "center",
            boxShadow: `0 8px 24px ${G.green}55`,
          }}>
            <span style={{ color: "#fff", fontSize: 24, fontWeight: 800 }}>A</span>
          </div>
          <h1 style={{ color: "#fff", fontSize: 22, fontWeight: 700, margin: 0 }}>Allotment</h1>
          <p style={{ color: "rgba(255,255,255,.5)", fontSize: 13, marginTop: 4 }}>Seat allocation, reimagined</p>
        </div>
        <div style={{ display: "flex", background: "rgba(255,255,255,.07)", borderRadius: 10, padding: 3, marginBottom: 24 }}>
          {["login", "register"].map(tb => (
            <button key={tb} onClick={() => { setTab(tb); setErrors({}); }} style={{
              flex: 1, padding: "8px 0", border: "none", borderRadius: 8, cursor: "pointer",
              background: tab === tb ? G.green : "transparent", color: tab === tb ? "#fff" : "rgba(255,255,255,.5)",
              fontWeight: 600, fontSize: 13, transition: "all .2s",
            }}>{tb === "register" ? "Register" : "Sign In"}</button>
          ))}
        </div>
        {errors.form && <div style={{ color: G.red, fontSize: 13, marginBottom: 16, textAlign: "center" }}>{errors.form}</div>}
        {tab === "register" && <Input label="Full Name" value={name} onChange={setName} placeholder="Your name" t={at} error={errors.name} />}
        <Input label="Email" type="email" value={email} onChange={setEmail} placeholder="you@example.com" t={at} error={errors.email} />
        <Input label="Password" type="password" value={pass} onChange={setPass} placeholder="••••••••" t={at} error={errors.pass} />
        
        <button disabled={loading} onClick={handleSubmit} style={{
          width: "100%", padding: "12px 0", border: "none", borderRadius: 10, cursor: loading ? "not-allowed" : "pointer",
          background: `linear-gradient(135deg,${G.green},${G.greenL})`, color: "#fff", fontWeight: 700, fontSize: 15,
          boxShadow: `0 4px 16px ${G.green}55`, transition: "all .2s", opacity: loading ? 0.7 : 1, marginTop: 16
        }}
          onMouseEnter={e => !loading && (e.currentTarget.style.transform = "translateY(-1px)")}
          onMouseLeave={e => !loading && (e.currentTarget.style.transform = "")}
        >{loading ? "Processing..." : (tab === "register" ? "Create Account" : "Sign In")}</button>
        <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "18px 0" }}>
          <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,.1)" }} />
          <span style={{ color: "rgba(255,255,255,.3)", fontSize: 12 }}>or</span>
          <div style={{ flex: 1, height: 1, background: "rgba(255,255,255,.1)" }} />
        </div>
        <button onClick={handleGoogle} style={{
          width: "100%", padding: "11px 0", border: "1px solid rgba(255,255,255,.2)",
          borderRadius: 10, cursor: "pointer", background: "rgba(255,255,255,.06)",
          color: "#fff", fontWeight: 600, fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center", gap: 10,
        }}
          onMouseEnter={e => e.currentTarget.style.background = "rgba(255,255,255,.1)"}
          onMouseLeave={e => e.currentTarget.style.background = "rgba(255,255,255,.06)"}
        ><Icon name="google" size={18} /> Continue with Google</button>
      </div>
    </div>
  );
}

// ─── HOME PAGE ───────────────────────────────────────────────────────────────
function HomePage({ t, groups, setPage, setSelectedGroup, addToast, addNotif }) {
  const [search, setSearch] = useState("");
  const [showModal, setShowModal] = useState(false);
  const [editGroup, setEditGroup] = useState(null);
  const [form, setForm] = useState({ title: "", description: "" });
  const [members, setMembers] = useState([]);
  const [memberForm, setMemberForm] = useState({ name: "", email: "", regNo: "", mobile: "" });
  const [importResult, setImportResult] = useState(null);
  const [formErrors, setFormErrors] = useState({});
  const [memberErrors, setMemberErrors] = useState({});
  const [confirmDel, setConfirmDel] = useState(null);
  const fileRef = useRef(null);

  const filtered = groups.filter(g => g.title.toLowerCase().includes(search.toLowerCase()) || g.description.toLowerCase().includes(search.toLowerCase()));

  function openCreate() { setEditGroup(null); setForm({ title: "", description: "" }); setMembers([]); setImportResult(null); setFormErrors({}); setShowModal(true); }
  function openEdit(g, e) { e.stopPropagation(); setEditGroup(g); setForm({ title: g.title, description: g.description }); setMembers([...g.memberList]); setImportResult(null); setFormErrors({}); setShowModal(true); }

  async function saveGroup() {
    const errs = {};
    if (!form.title.trim()) errs.title = "Title is required";
    if (!form.description.trim()) errs.description = "Description is required";
    if (groups.some(g => g.title.toLowerCase() === form.title.toLowerCase().trim() && g.id !== editGroup?.id)) errs.title = "Group with this name already exists";
    setFormErrors(errs);
    if (Object.keys(errs).length > 0) return;

    try {
      if (editGroup) {
        await updateDoc(doc(db, "groups", editGroup.id), { title: form.title.trim(), description: form.description.trim(), memberList: members });
        addToast("Group updated successfully");
        addNotif(`Group "${form.title}" updated`);
      } else {
        await addDoc(collection(db, "groups"), { title: form.title.trim(), description: form.description.trim(), memberList: members, created: new Date().toISOString().split("T")[0] });
        addToast(`Group "${form.title}" created with ${members.length} members`);
        addNotif(`New group "${form.title}" created`);
      }
      setShowModal(false);
    } catch (e) {
      addToast(e.message, "error");
    }
  }

  function addMember() {
    const errs = {};
    if (!memberForm.name.trim()) errs.name = "Required";
    if (!memberForm.email.trim()) errs.email = "Required";
    else if (!validEmail(memberForm.email)) errs.email = "Invalid email";
    if (members.some(m => m.email.toLowerCase() === memberForm.email.toLowerCase())) errs.email = "Duplicate email";
    if (memberForm.regNo && members.some(m => m.regNo === memberForm.regNo)) errs.regNo = "Duplicate Reg No";
    setMemberErrors(errs);
    if (Object.keys(errs).length > 0) return;
    setMembers([...members, { id: `m${genId()}`, ...memberForm }]);
    setMemberForm({ name: "", email: "", regNo: "", mobile: "" });
    setMemberErrors({});
  }

  function handleImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let success = 0, failed = [];
        const existEmails = new Set(members.map(m => m.email.toLowerCase()));
        const existRegs = new Set(members.filter(m => m.regNo).map(m => m.regNo));
        const newMembers = [];
        rows.forEach((row, i) => {
          const name = String(row["Name"] || row["name"] || "").trim();
          const email = String(row["Email"] || row["email"] || "").trim();
          const regNo = String(row["Register Number"] || row["register number"] || row["RegNo"] || row["regNo"] || "").trim();
          const mobile = String(row["Mobile Number"] || row["mobile number"] || row["Mobile"] || row["mobile"] || "").trim();
          if (!name || !email) { failed.push({ row: i + 2, reason: "Missing name or email" }); return; }
          if (!validEmail(email)) { failed.push({ row: i + 2, reason: `Invalid email: ${email}` }); return; }
          if (existEmails.has(email.toLowerCase())) { failed.push({ row: i + 2, reason: `Duplicate email: ${email}` }); return; }
          if (regNo && existRegs.has(regNo)) { failed.push({ row: i + 2, reason: `Duplicate Reg No: ${regNo}` }); return; }
          existEmails.add(email.toLowerCase());
          if (regNo) existRegs.add(regNo);
          newMembers.push({ id: `m${genId()}_${i}`, name, email, regNo, mobile });
          success++;
        });
        setMembers(prev => [...prev, ...newMembers]);
        setImportResult({ success, failed });
        if (success > 0) addToast(`${success} members imported`);
        if (failed.length > 0) addToast(`${failed.length} entries failed`, "warning");
      } catch { addToast("Failed to parse Excel file", "error"); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  function deleteGroup(g, e) { e.stopPropagation(); setConfirmDel(g); }

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "32px 24px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28, flexWrap: "wrap", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 24, fontWeight: 700, color: t.text }}>My Groups</h2>
          <p style={{ margin: "4px 0 0", color: t.textSec, fontSize: 14 }}>{groups.length} group{groups.length !== 1 ? "s" : ""} total</p>
        </div>
        <Btn t={t} onClick={openCreate}><Icon name="plus" size={16} color="#fff" /> New Group</Btn>
      </div>

      <div style={{ position: "relative", marginBottom: 24 }}>
        <div style={{ position: "absolute", left: 14, top: "50%", transform: "translateY(-50%)", pointerEvents: "none" }}><Icon name="search" size={16} color={t.textMut} /></div>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search groups…"
          style={{ width: "100%", boxSizing: "border-box", padding: "10px 14px 10px 40px", borderRadius: 10, border: `1px solid ${t.border}`, background: t.surface, color: t.text, fontSize: 14, outline: "none", fontFamily: "inherit" }} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: 18 }}>
        {filtered.map(g => (
          <Card key={g.id} t={t} onClick={() => { setSelectedGroup(g); setPage("group"); }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
              <div style={{ width: 42, height: 42, borderRadius: 12, background: `linear-gradient(135deg,${G.green}22,${G.greenL}33)`, display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${G.green}33` }}>
                <Icon name="users" size={20} color={t.accent} />
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button onClick={e => openEdit(g, e)} style={{ background: "none", border: "none", cursor: "pointer", padding: 4 }}><Icon name="edit" size={15} color={t.textSec} /></button>
                <button onClick={e => deleteGroup(g, e)} style={{ background: "none", border: "none", cursor: "pointer", padding: 4 }}><Icon name="trash" size={15} color={G.red} /></button>
              </div>
            </div>
            <h3 style={{ margin: "0 0 6px", fontSize: 16, fontWeight: 700, color: t.text }}>{g.title}</h3>
            <p style={{ margin: "0 0 16px", fontSize: 13, color: t.textSec, lineHeight: 1.5 }}>{g.description}</p>
            <div style={{ display: "flex", gap: 16 }}>
              <span style={{ fontSize: 12, color: t.textMut }}><strong style={{ color: t.text }}>{g.memberList?.length || 0}</strong> members</span>
              <span style={{ fontSize: 12, color: t.textMut }}>{g.created}</span>
            </div>
          </Card>
        ))}
        {filtered.length === 0 && (
          <div style={{ gridColumn: "1/-1", textAlign: "center", padding: "60px 0", color: t.textSec }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>📋</div>
            <p style={{ fontSize: 16, fontWeight: 600, color: t.text, margin: 0 }}>No groups found</p>
            <p style={{ fontSize: 13, marginTop: 6 }}>Create a group to get started</p>
          </div>
        )}
      </div>

      {/* Create/Edit Group Modal */}
      {showModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", display: "flex", alignItems: "flex-start", justifyContent: "center", zIndex: 200, padding: "40px 20px", overflowY: "auto" }} onClick={() => setShowModal(false)}>
          <Card t={t} style={{ width: "100%", maxWidth: 600 }} onClick={e => e.stopPropagation()}>
            <h3 style={{ margin: "0 0 20px", color: t.text, fontSize: 18 }}>{editGroup ? "Edit Group" : "Create New Group"}</h3>
            <Input label="Group Name" value={form.title} onChange={v => setForm({ ...form, title: v })} placeholder="e.g. Engineering 2025" t={t} error={formErrors.title} />
            <Input label="Description" value={form.description} onChange={v => setForm({ ...form, description: v })} placeholder="Brief description" t={t} error={formErrors.description} />

            {/* Members Section */}
            <div style={{ borderTop: `1px solid ${t.border}`, marginTop: 8, paddingTop: 16 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 8 }}>
                <h4 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: t.text }}>Members ({members.length})</h4>
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn t={t} size="sm" variant="secondary" onClick={downloadTemplate}><Icon name="download" size={13} color={t.textSec} /> Template</Btn>
                  <Btn t={t} size="sm" variant="secondary" onClick={() => fileRef.current?.click()}><Icon name="upload" size={13} color={t.textSec} /> Import Excel</Btn>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleImport} style={{ display: "none" }} />
                </div>
              </div>

              {importResult && (
                <div style={{ padding: 12, borderRadius: 8, background: t.surface2, marginBottom: 12, fontSize: 13 }}>
                  <div style={{ color: G.green, fontWeight: 600 }}>✓ {importResult.success} imported successfully</div>
                  {importResult.failed.length > 0 && (
                    <div style={{ marginTop: 6 }}>
                      <div style={{ color: G.red, fontWeight: 600 }}>✗ {importResult.failed.length} failed:</div>
                      {importResult.failed.slice(0, 5).map((f, i) => <div key={i} style={{ color: t.textSec, marginLeft: 12 }}>Row {f.row}: {f.reason}</div>)}
                      {importResult.failed.length > 5 && <div style={{ color: t.textMut, marginLeft: 12 }}>...and {importResult.failed.length - 5} more</div>}
                    </div>
                  )}
                </div>
              )}

              {/* Manual Add */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                <Input label="Name" value={memberForm.name} onChange={v => setMemberForm({ ...memberForm, name: v })} placeholder="Full name" t={t} style={{ marginBottom: 8 }} error={memberErrors.name} />
                <Input label="Email" value={memberForm.email} onChange={v => setMemberForm({ ...memberForm, email: v })} placeholder="email@example.com" t={t} style={{ marginBottom: 8 }} error={memberErrors.email} />
                <Input label="Register Number" value={memberForm.regNo} onChange={v => setMemberForm({ ...memberForm, regNo: v })} placeholder="CS001" t={t} style={{ marginBottom: 8 }} error={memberErrors.regNo} />
                <Input label="Mobile" value={memberForm.mobile} onChange={v => setMemberForm({ ...memberForm, mobile: v })} placeholder="9876543210" t={t} style={{ marginBottom: 8 }} />
              </div>
              <Btn t={t} size="sm" onClick={addMember} style={{ marginBottom: 12 }}><Icon name="plus" size={13} color="#fff" /> Add Member</Btn>

              {/* Member List */}
              {members.length > 0 && (
                <div style={{ maxHeight: 200, overflowY: "auto", border: `1px solid ${t.border}`, borderRadius: 8 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead><tr style={{ background: t.surface2 }}>
                      {["Name", "Email", "Reg No", "Mobile", ""].map(h => <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: t.textSec, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>)}
                    </tr></thead>
                    <tbody>
                      {members.map(m => (
                        <tr key={m.id} style={{ borderTop: `1px solid ${t.border}` }}>
                          <td style={{ padding: "6px 10px", color: t.text }}>{m.name}</td>
                          <td style={{ padding: "6px 10px", color: t.textSec }}>{m.email}</td>
                          <td style={{ padding: "6px 10px", color: t.textSec }}>{m.regNo}</td>
                          <td style={{ padding: "6px 10px", color: t.textSec }}>{m.mobile}</td>
                          <td style={{ padding: "6px 10px" }}><button onClick={() => setMembers(members.filter(x => x.id !== m.id))} style={{ background: "none", border: "none", cursor: "pointer" }}><Icon name="close" size={13} color={G.red} /></button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 20 }}>
              <Btn t={t} variant="secondary" onClick={() => setShowModal(false)}>Cancel</Btn>
              <Btn t={t} onClick={saveGroup}>{editGroup ? "Save Changes" : "Create Group"}</Btn>
            </div>
          </Card>
        </div>
      )}

      {confirmDel && <ConfirmModal t={t} title="Delete Group" message={`Are you sure you want to delete "${confirmDel.title}"? This action cannot be undone.`}
        onCancel={() => setConfirmDel(null)} onConfirm={async () => {
          try {
            await deleteDoc(doc(db, "groups", confirmDel.id));
            addToast("Group deleted");
            addNotif(`Group "${confirmDel.title}" deleted`);
            setConfirmDel(null);
          } catch(e) { addToast(e.message, "error"); }
        }} />}
    </div>
  );
}

// ─── GROUP PAGE ──────────────────────────────────────────────────────────────
function GroupPage({ t, group, pools, setPage, setSelectedPool, addToast, addNotif }) {
  const groupPools = pools[group.id] || [];
  const [confirmDel, setConfirmDel] = useState(null);

  function deletePool(pool) { setConfirmDel(pool); }
  async function confirmDeletePool() {
    try {
      await deleteDoc(doc(db, "pools", confirmDel.id));
      addToast("Pool deleted");
      addNotif(`Pool "${confirmDel.title}" deleted`);
      setConfirmDel(null);
    } catch(e) { addToast(e.message, "error"); }
  }

  return (
    <div style={{ maxWidth: 1000, margin: "0 auto", padding: "32px 24px" }}>
      <button onClick={() => setPage("home")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", color: t.textSec, fontSize: 14, marginBottom: 20, padding: 0 }}>
        <Icon name="back" size={16} color={t.textSec} /> Back to Groups
      </button>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 28, flexWrap: "wrap", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 24, fontWeight: 700, color: t.text }}>{group.title}</h2>
          <p style={{ margin: "4px 0 0", color: t.textSec, fontSize: 14 }}>{group.description} · {group.memberList?.length || 0} members</p>
        </div>
        <Btn t={t} onClick={() => { setSelectedPool(null); setPage("createPool"); }}>
          <Icon name="plus" size={16} color="#fff" /> New Pool
        </Btn>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))", gap: 18 }}>
        {groupPools.map(pool => {
          const status = getPoolStatus(pool);
          return (
            <Card key={pool.id} t={t}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                <StatusBadge pool={pool} t={t} />
                <div style={{ display: "flex", gap: 6 }}>
                  <button onClick={() => { setSelectedPool(pool); setPage("createPool"); }} style={{ background: "none", border: "none", cursor: "pointer" }}><Icon name="edit" size={15} color={t.textSec} /></button>
                  <button onClick={() => deletePool(pool)} style={{ background: "none", border: "none", cursor: "pointer" }}><Icon name="trash" size={15} color={G.red} /></button>
                </div>
              </div>
              <h3 style={{ margin: "0 0 4px", fontSize: 16, fontWeight: 700, color: t.text }}>{pool.title}</h3>
              <p style={{ margin: "0 0 8px", fontSize: 13, color: t.textSec }}>{pool.description}</p>
              <div style={{ display: "flex", gap: 12, fontSize: 11, color: t.textMut, marginBottom: 14, flexWrap: "wrap" }}>
                <span><Icon name="clock" size={11} color={t.textMut} /> {fmtDate(pool.startTime)}</span>
                <span>→ {fmtDate(pool.endTime)}</span>
              </div>

              {pool.options.length > 0 && (
                <div style={{ marginBottom: 14 }}>
                  {pool.options.slice(0, 2).map(o => (
                    <div key={o.id} style={{ marginBottom: 8 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, marginBottom: 4 }}>
                        <span style={{ color: t.text, fontWeight: 500 }}>{o.name}</span>
                        <span style={{ color: t.textSec }}>{o.filledSeats || 0}/{o.seatLimit}</span>
                      </div>
                      <ProgressBar value={pct(o.filledSeats || 0, o.seatLimit)} t={t} />
                    </div>
                  ))}
                  {pool.options.length > 2 && <span style={{ fontSize: 11, color: t.textMut }}>+{pool.options.length - 2} more</span>}
                </div>
              )}

              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {status === "In Progress" && (
                  <Btn t={t} size="sm" onClick={() => { setSelectedPool(pool); setPage("livePool"); }}>
                    <Icon name="pool" size={13} color="#fff" /> Join Pool
                  </Btn>
                )}
                {status === "Upcoming" && <Badge color={t.accent} bg={t.accentBg}>Opens {fmtDate(pool.startTime)}</Badge>}
                <Btn t={t} size="sm" variant="secondary" onClick={() => { setSelectedPool(pool); setPage("responses"); }}>Responses</Btn>
              </div>
            </Card>
          );
        })}
        {groupPools.length === 0 && (
          <div style={{ gridColumn: "1/-1", textAlign: "center", padding: "60px 0", color: t.textSec }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>🗂️</div>
            <p style={{ fontSize: 16, fontWeight: 600, color: t.text, margin: 0 }}>No pools yet</p>
            <p style={{ fontSize: 13, marginTop: 6 }}>Create your first allocation pool</p>
          </div>
        )}
      </div>

      {confirmDel && <ConfirmModal t={t} title="Delete Pool" message={`Delete "${confirmDel.title}"? All options and responses will be lost.`}
        onCancel={() => setConfirmDel(null)} onConfirm={confirmDeletePool} />}
    </div>
  );
}

// ─── CREATE/EDIT POOL PAGE ───────────────────────────────────────────────────
function CreatePoolPage({ t, pool, setPage, selectedGroup, addToast, addNotif }) {
  const [form, setForm] = useState({
    title: pool?.title || "", description: pool?.description || "",
    instructions: pool?.instructions || "",
    startTime: pool?.startTime || "", endTime: pool?.endTime || "",
  });
  const [options, setOptions] = useState(pool?.options?.map(o => ({ ...o })) || []);
  const [errors, setErrors] = useState({});

  function addOption() {
    setOptions([...options, { id: `o${genId()}`, name: "", description: "", seatLimit: 30, filledSeats: 0 }]);
  }
  function removeOption(id) { setOptions(options.filter(o => o.id !== id)); }
  function updateOption(id, key, val) { setOptions(options.map(o => o.id === id ? { ...o, [key]: val } : o)); }

  async function savePool(publish = false) {
    const errs = {};
    if (!form.title.trim()) errs.title = "Title is required";
    if (!form.startTime) errs.startTime = "Start time is required";
    if (!form.endTime) errs.endTime = "End time is required";
    if (form.startTime && form.endTime && new Date(form.startTime) >= new Date(form.endTime)) errs.endTime = "End must be after start";
    if (options.length === 0 && publish) errs.options = "Add at least one option";
    options.forEach((o, i) => { if (!o.name.trim()) errs[`opt_${i}`] = "Option name required"; if (o.seatLimit <= 0) errs[`seat_${i}`] = "Seat limit must be > 0"; });
    setErrors(errs);
    if (Object.keys(errs).length > 0) { addToast("Please fix the errors", "error"); return; }

    const poolData = { ...form, options: options.map(o => ({ ...o, filledSeats: pool ? o.filledSeats || 0 : 0 })) };

    try {
      if (pool && selectedGroup) {
        await updateDoc(doc(db, "pools", pool.id), poolData);
        addToast("Pool updated successfully");
        addNotif(`Pool "${form.title}" updated`);
      } else if (selectedGroup) {
        await addDoc(collection(db, "pools"), { ...poolData, groupId: selectedGroup.id });
        addToast(`Pool "${form.title}" created`);
        addNotif(`New pool "${form.title}" created`);
      }
      setPage("group");
    } catch(e) { addToast(e.message, "error"); }
  }

  return (
    <div style={{ maxWidth: 760, margin: "0 auto", padding: "32px 24px" }}>
      <button onClick={() => setPage("group")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", color: t.textSec, fontSize: 14, marginBottom: 20, padding: 0 }}>
        <Icon name="back" size={16} color={t.textSec} /> Back
      </button>
      <h2 style={{ margin: "0 0 24px", fontSize: 22, fontWeight: 700, color: t.text }}>{pool ? "Edit Pool" : "Create Pool"}</h2>

      <Card t={t} style={{ marginBottom: 20 }}>
        <h3 style={{ margin: "0 0 16px", fontSize: 16, fontWeight: 600, color: t.text }}>Pool Details</h3>
        <Input label="Pool Title *" value={form.title} onChange={v => setForm({ ...form, title: v })} placeholder="e.g. Branch Allocation 2025" t={t} error={errors.title} />
        <Input label="Description" value={form.description} onChange={v => setForm({ ...form, description: v })} placeholder="Brief description" t={t} />
        <Input label="Instructions" value={form.instructions} onChange={v => setForm({ ...form, instructions: v })} placeholder="Instructions for participants" t={t} />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Input label="Start Date & Time *" type="datetime-local" value={form.startTime} onChange={v => setForm({ ...form, startTime: v })} t={t} error={errors.startTime} />
          <Input label="End Date & Time *" type="datetime-local" value={form.endTime} onChange={v => setForm({ ...form, endTime: v })} t={t} error={errors.endTime} />
        </div>
      </Card>

      <Card t={t} style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: t.text }}>Options {errors.options && <span style={{ fontSize: 12, color: G.red, fontWeight: 400 }}>— {errors.options}</span>}</h3>
          <Btn t={t} size="sm" onClick={addOption}><Icon name="plus" size={13} color="#fff" /> Add Option</Btn>
        </div>
        {options.map((opt, idx) => (
          <div key={opt.id} style={{ border: `1px solid ${t.border}`, borderRadius: 10, padding: 16, marginBottom: 12, background: t.surface2 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <span style={{ fontWeight: 600, fontSize: 13, color: t.textSec }}>Option {idx + 1}</span>
              <button onClick={() => removeOption(opt.id)} style={{ background: "none", border: "none", cursor: "pointer" }}><Icon name="trash" size={15} color={G.red} /></button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 10 }}>
              <div>
                <label style={{ fontSize: 12, color: t.textSec, display: "block", marginBottom: 4 }}>Name *</label>
                <input value={opt.name} onChange={e => updateOption(opt.id, "name", e.target.value)} placeholder="Option name"
                  style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 7, border: `1px solid ${errors[`opt_${idx}`] ? G.red : t.border}`, background: t.surface, color: t.text, fontSize: 13, outline: "none" }} />
                {errors[`opt_${idx}`] && <div style={{ fontSize: 11, color: G.red, marginTop: 2 }}>{errors[`opt_${idx}`]}</div>}
              </div>
              <div>
                <label style={{ fontSize: 12, color: t.textSec, display: "block", marginBottom: 4 }}>Seat Limit *</label>
                <input type="number" value={opt.seatLimit} onChange={e => updateOption(opt.id, "seatLimit", Math.max(0, +e.target.value))}
                  style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 7, border: `1px solid ${errors[`seat_${idx}`] ? G.red : t.border}`, background: t.surface, color: t.text, fontSize: 13, outline: "none" }} />
              </div>
            </div>
            <div style={{ marginTop: 8 }}>
              <label style={{ fontSize: 12, color: t.textSec, display: "block", marginBottom: 4 }}>Description (optional)</label>
              <input value={opt.description || ""} onChange={e => updateOption(opt.id, "description", e.target.value)} placeholder="Brief description"
                style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px", borderRadius: 7, border: `1px solid ${t.border}`, background: t.surface, color: t.text, fontSize: 13, outline: "none" }} />
            </div>
          </div>
        ))}
        {options.length === 0 && <div style={{ textAlign: "center", padding: "32px 0", color: t.textMut, fontSize: 14 }}>No options yet. Add one to get started.</div>}
      </Card>

      <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", flexWrap: "wrap" }}>
        <Btn t={t} variant="secondary" onClick={() => setPage("group")}>Cancel</Btn>
        <Btn t={t} variant="secondary" onClick={() => savePool(false)}>Save Draft</Btn>
        <Btn t={t} onClick={() => savePool(true)}><Icon name="check" size={15} color="#fff" /> Publish Pool</Btn>
      </div>
    </div>
  );
}

// ─── LIVE POOL PAGE ──────────────────────────────────────────────────────────
function LivePoolPage({ t, pool, setPage, pools, selectedGroup, responses, user, addToast }) {
  const [selected, setSelected] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const poolResponses = responses[pool?.id] || [];
  const alreadySubmitted = poolResponses.some(r => r.userId === user?.uid);

  const currentPool = (pools[selectedGroup?.id] || []).find(p => p.id === pool?.id) || pool;
  const status = getPoolStatus(currentPool);
  const liveOptions = currentPool?.options || [];

  async function submit() {
    if (!selected || status !== "In Progress" || alreadySubmitted) return;
    const opt = liveOptions.find(o => o.id === selected);
    if (!opt || (opt.filledSeats || 0) >= opt.seatLimit) { addToast("This option is full", "error"); return; }
    
    setSubmitting(true);
    try {
      const poolRef = doc(db, "pools", pool.id);
      await runTransaction(db, async (transaction) => {
        const pDoc = await transaction.get(poolRef);
        if (!pDoc.exists()) throw new Error("Pool not found");
        
        const pData = pDoc.data();
        const optIndex = pData.options.findIndex(o => o.id === selected);
        if (optIndex === -1) throw new Error("Option not found");
        
        const currentSeats = pData.options[optIndex].filledSeats || 0;
        if (currentSeats >= pData.options[optIndex].seatLimit) {
          throw new Error("Option is now full. Please select another.");
        }
        
        // Update pool
        pData.options[optIndex].filledSeats = currentSeats + 1;
        transaction.update(poolRef, { options: pData.options });
        
        // Add response
        const newRespRef = doc(collection(db, "responses"));
        transaction.set(newRespRef, {
          poolId: pool.id,
          userId: user.uid,
          name: user.name,
          email: user.email,
          selected: pData.options[optIndex].name,
          allotted: pData.options[optIndex].name,
          time: new Date().toISOString(),
          verified: true
        });
      });
      addToast("Preference submitted successfully!");
    } catch(e) {
      addToast(e.message, "error");
    }
    setSubmitting(false);
  }

  if (alreadySubmitted) {
    const resp = poolResponses.find(r => r.userId === user?.uid);
    return (
      <div style={{ maxWidth: 520, margin: "80px auto", padding: "0 24px", textAlign: "center" }}>
        <div style={{ width: 72, height: 72, borderRadius: "50%", background: `linear-gradient(135deg,${G.green},${G.greenL})`, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 24px", boxShadow: `0 8px 24px ${G.green}44` }}>
          <Icon name="check" size={32} color="#fff" />
        </div>
        <h2 style={{ color: t.text, fontSize: 22, fontWeight: 700, marginBottom: 8 }}>Submitted!</h2>
        <p style={{ color: t.textSec, marginBottom: 28 }}>Your preference has been recorded and you're in the queue.</p>
        <Card t={t} style={{ textAlign: "left", marginBottom: 20 }}>
          <div style={{ display: "grid", gap: 12 }}>
            {[["Selected", resp?.selected], ["Time", fmtDate(resp?.time)], ["Status", resp?.allotted === "Pending" ? "Pending Allotment" : "Allotted"]].map(([k, v]) => (
              <div key={k} style={{ display: "flex", justifyContent: "space-between" }}>
                <span style={{ color: t.textSec, fontSize: 14 }}>{k}</span>
                <span style={{ color: t.text, fontWeight: 600, fontSize: 14 }}>{v}</span>
              </div>
            ))}
          </div>
        </Card>
        <Btn t={t} variant="secondary" onClick={() => setPage("group")}>Back to Group</Btn>
      </div>
    );
  }

  const sc = statusColor(status, t);

  return (
    <div style={{ maxWidth: 700, margin: "0 auto", padding: "32px 24px" }}>
      <button onClick={() => setPage("group")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", color: t.textSec, fontSize: 14, marginBottom: 20, padding: 0 }}>
        <Icon name="back" size={16} color={t.textSec} /> Back
      </button>

      <div style={{ marginBottom: 24 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}><StatusBadge pool={currentPool} t={t} /></div>
        <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: t.text }}>{pool?.title}</h2>
        <p style={{ margin: "6px 0 0", color: t.textSec }}>{pool?.description}</p>
        {currentPool?.instructions && <p style={{ margin: "8px 0 0", color: t.textMut, fontSize: 13, fontStyle: "italic" }}>{currentPool.instructions}</p>}
        <div style={{ display: "flex", gap: 16, marginTop: 10, fontSize: 12, color: t.textMut }}>
          <span><Icon name="clock" size={12} color={t.textMut} /> Ends: {fmtDate(currentPool?.endTime)}</span>
        </div>
      </div>

      {status !== "In Progress" && (
        <Card t={t} style={{ marginBottom: 16, background: `${G.amber}15`, border: `1px solid ${G.amber}33` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, color: G.amber }}>
            <Icon name="clock" size={16} color={G.amber} />
            <span style={{ fontSize: 13, fontWeight: 500 }}>
              {status === "Upcoming" ? `Pool opens at ${fmtDate(currentPool?.startTime)}` : "This pool has ended. Submissions are closed."}
            </span>
          </div>
        </Card>
      )}

      {status === "In Progress" && (
        <Card t={t} style={{ marginBottom: 16, background: `linear-gradient(135deg,${t.accentBg},${t.surface})`, border: `1px solid ${t.accent}33` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, color: t.accent }}>
            <Icon name="lock" size={16} color={t.accent} />
            <span style={{ fontSize: 13, fontWeight: 500 }}>First Come First Serve — earliest submissions get first priority</span>
          </div>
        </Card>
      )}

      <h3 style={{ fontSize: 16, fontWeight: 600, color: t.text, marginBottom: 16 }}>Choose Your Option</h3>
      <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 24 }}>
        {liveOptions.map(opt => {
          const filled = opt.filledSeats || 0;
          const full = filled >= opt.seatLimit;
          const p = pct(filled, opt.seatLimit);
          const isSelected = selected === opt.id;
          const canSelect = status === "In Progress" && !full && !submitting;
          return (
            <div key={opt.id} onClick={() => canSelect && setSelected(opt.id)}
              style={{
                border: `2px solid ${isSelected ? t.accent : t.border}`, borderRadius: 12, padding: 18,
                cursor: canSelect ? "pointer" : "not-allowed",
                background: isSelected ? t.accentBg : full ? t.surface2 : t.surface,
                opacity: full ? 0.6 : 1, transition: "all .2s",
              }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ width: 20, height: 20, borderRadius: "50%", border: `2px solid ${isSelected ? t.accent : t.border}`, background: isSelected ? t.accent : "transparent", display: "flex", alignItems: "center", justifyContent: "center", transition: "all .2s" }}>
                    {isSelected && <Icon name="check" size={11} color="#fff" />}
                  </div>
                  <div>
                    <span style={{ fontWeight: 600, fontSize: 15, color: full ? t.textMut : t.text }}>{opt.name}</span>
                    {opt.description && <div style={{ fontSize: 12, color: t.textMut }}>{opt.description}</div>}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {full && <Badge color={G.red} bg="#fee2e2">FULL</Badge>}
                  <span style={{ fontSize: 13, color: t.textSec, fontWeight: 500 }}>{filled}/{opt.seatLimit}</span>
                </div>
              </div>
              <ProgressBar value={p} t={t} />
              <div style={{ fontSize: 12, color: t.textMut, marginTop: 6 }}>{full ? "No seats remaining" : `${opt.seatLimit - filled} seats available`}</div>
            </div>
          );
        })}
      </div>
      <Btn t={t} onClick={submit} disabled={!selected || status !== "In Progress" || submitting} style={{ width: "100%" }}>
        {submitting ? "Submitting..." : "Submit Preference"}
      </Btn>
    </div>
  );
}

// ─── RESPONSES PAGE ──────────────────────────────────────────────────────────
function ResponsesPage({ t, pool, setPage, responses }) {
  const [search, setSearch] = useState("");
  const poolResp = responses[pool?.id] || [];
  const filtered = poolResp.filter(r => r.name?.toLowerCase().includes(search.toLowerCase()) || r.email?.toLowerCase().includes(search.toLowerCase()));

  function doExport(format) {
    const data = poolResp.map(r => ({ Name: r.name, Email: r.email, Selected: r.selected, Allotted: r.allotted, Time: fmtDate(r.time), Status: r.verified ? "Verified" : "Pending" }));
    if (data.length === 0) return;
    const fn = `${pool?.title || "responses"}_${Date.now()}`;
    if (format === "excel") exportExcel(data, `${fn}.xlsx`);
    else if (format === "csv") exportCSV(data, `${fn}.csv`);
    else exportPDF(data, pool?.title || "Responses");
  }

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "32px 24px" }}>
      <button onClick={() => setPage("group")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", color: t.textSec, fontSize: 14, marginBottom: 20, padding: 0 }}>
        <Icon name="back" size={16} color={t.textSec} /> Back
      </button>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24, flexWrap: "wrap", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: t.text }}>Responses</h2>
          <p style={{ margin: "4px 0 0", color: t.textSec, fontSize: 14 }}>{pool?.title} · {poolResp.length} responses</p>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Btn t={t} variant="secondary" size="sm" onClick={() => doExport("excel")}><Icon name="export" size={13} color={t.textSec} /> Excel</Btn>
          <Btn t={t} variant="secondary" size="sm" onClick={() => doExport("csv")}><Icon name="export" size={13} color={t.textSec} /> CSV</Btn>
          <Btn t={t} variant="secondary" size="sm" onClick={() => doExport("pdf")}><Icon name="file" size={13} color={t.textSec} /> PDF</Btn>
        </div>
      </div>

      <Card t={t} style={{ marginBottom: 16 }}>
        <div style={{ position: "relative" }}>
          <div style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)" }}><Icon name="search" size={15} color={t.textMut} /></div>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search by name or email…"
            style={{ width: "100%", boxSizing: "border-box", padding: "9px 14px 9px 36px", borderRadius: 8, border: `1px solid ${t.border}`, background: t.surface2, color: t.text, fontSize: 13, outline: "none" }} />
        </div>
      </Card>

      <Card t={t} style={{ overflow: "auto", padding: 0 }}>
        {filtered.length === 0 ? (
          <div style={{ padding: "40px 20px", textAlign: "center", color: t.textMut }}>No responses yet</div>
        ) : (
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ background: t.surface2 }}>
              {["Name", "Email", "Selected", "Allotted", "Time", "Status"].map(h => (
                <th key={h} style={{ padding: "12px 16px", textAlign: "left", fontWeight: 600, color: t.textSec, whiteSpace: "nowrap", borderBottom: `1px solid ${t.border}` }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>{filtered.map((r, i) => (
              <tr key={r.id} style={{ borderBottom: `1px solid ${t.border}`, background: i % 2 === 0 ? t.surface : t.surface2 + "44" }}>
                <td style={{ padding: "12px 16px", color: t.text, fontWeight: 500, whiteSpace: "nowrap" }}>{r.name}</td>
                <td style={{ padding: "12px 16px", color: t.textSec }}>{r.email}</td>
                <td style={{ padding: "12px 16px", color: t.textSec }}>{r.selected}</td>
                <td style={{ padding: "12px 16px" }}><span style={{ color: r.allotted === "Pending" ? G.amber : t.accent, fontWeight: 500 }}>{r.allotted}</span></td>
                <td style={{ padding: "12px 16px", color: t.textSec, whiteSpace: "nowrap" }}>{fmtDate(r.time)}</td>
                <td style={{ padding: "12px 16px" }}><Badge color={r.verified ? "#065f46" : "#92400e"} bg={r.verified ? "#d1fae5" : "#fef3c7"}>{r.verified ? "Verified" : "Pending"}</Badge></td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

// ─── PROFILE PAGE ────────────────────────────────────────────────────────────
function ProfilePage({ t, user, setPage, addToast }) {
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({ name: user.name || "", phone: user.phone || "", regNo: user.regNo || "" });

  async function saveProfile() {
    if (!form.name.trim()) { addToast("Name is required", "error"); return; }
    try {
      await updateDoc(doc(db, "users", user.uid), form);
      setEditing(false);
      addToast("Profile updated successfully");
    } catch(e) { addToast(e.message, "error"); }
  }

  return (
    <div style={{ maxWidth: 600, margin: "0 auto", padding: "32px 24px" }}>
      <button onClick={() => setPage("home")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", color: t.textSec, fontSize: 14, marginBottom: 20, padding: 0 }}>
        <Icon name="back" size={16} color={t.textSec} /> Back
      </button>

      <div style={{ textAlign: "center", marginBottom: 28 }}>
        <div style={{ width: 72, height: 72, borderRadius: "50%", background: `linear-gradient(135deg,${G.green},${G.greenL})`, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px", fontSize: 28, fontWeight: 800, color: "#fff" }}>
          {user.name?.[0]?.toUpperCase() || "U"}
        </div>
        <h2 style={{ margin: 0, color: t.text, fontSize: 22 }}>{user.name}</h2>
        <p style={{ color: t.textSec, fontSize: 14, margin: "4px 0" }}>{user.email}</p>
        <Badge color={user.role === "admin" ? "#92400e" : t.accent} bg={user.role === "admin" ? "#fef3c7" : t.accentBg}>{user.role}</Badge>
      </div>

      <Card t={t} style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600, color: t.text }}>Profile Details</h3>
          {!editing && <Btn t={t} size="sm" variant="secondary" onClick={() => setEditing(true)}><Icon name="edit" size={13} color={t.textSec} /> Edit</Btn>}
        </div>
        {editing ? (
          <>
            <Input label="Full Name" value={form.name} onChange={v => setForm({ ...form, name: v })} t={t} />
            <Input label="Phone" value={form.phone} onChange={v => setForm({ ...form, phone: v })} placeholder="9876543210" t={t} />
            <Input label="Register Number" value={form.regNo} onChange={v => setForm({ ...form, regNo: v })} placeholder="CS001" t={t} />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <Btn t={t} variant="secondary" onClick={() => { setEditing(false); setForm({ name: user.name || "", phone: user.phone || "", regNo: user.regNo || "" }); }}>Cancel</Btn>
              <Btn t={t} onClick={saveProfile}>Save Changes</Btn>
            </div>
          </>
        ) : (
          <div style={{ display: "grid", gap: 12 }}>
            {[["Name", user.name], ["Email", user.email], ["Phone", user.phone || "—"], ["Register No", user.regNo || "—"], ["Role", user.role]].map(([k, v]) => (
              <div key={k} style={{ display: "flex", justifyContent: "space-between", padding: "8px 0", borderBottom: `1px solid ${t.border}` }}>
                <span style={{ color: t.textSec, fontSize: 14 }}>{k}</span>
                <span style={{ color: t.text, fontWeight: 500, fontSize: 14 }}>{v}</span>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

// ─── SETTINGS PAGE ───────────────────────────────────────────────────────────
function SettingsPage({ t, dark, setDark, groups, pools, responses, addToast }) {
  const [confirm, setConfirm] = useState(null);

  const Section = ({ title, children }) => (<Card t={t} style={{ marginBottom: 16 }}><h3 style={{ margin: "0 0 16px", fontSize: 15, fontWeight: 700, color: t.text }}>{title}</h3>{children}</Card>);
  const Row = ({ label, desc, action }) => (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0", borderBottom: `1px solid ${t.border}`, gap: 12, flexWrap: "wrap" }}>
      <div style={{ flex: 1, minWidth: 150 }}><div style={{ fontSize: 14, fontWeight: 500, color: t.text }}>{label}</div>{desc && <div style={{ fontSize: 12, color: t.textMut, marginTop: 2 }}>{desc}</div>}</div>
      {action}
    </div>
  );
  const Toggle = ({ on, toggle }) => (
    <div onClick={toggle} style={{ width: 44, height: 24, borderRadius: 99, background: on ? t.accent : t.surface2, position: "relative", cursor: "pointer", transition: "background .2s", flexShrink: 0 }}>
      <div style={{ position: "absolute", top: 2, left: on ? 22 : 2, width: 20, height: 20, borderRadius: "50%", background: "#fff", transition: "left .2s", boxShadow: "0 1px 4px rgba(0,0,0,.2)" }} />
    </div>
  );

  function handleExport() {
    const allPools = Object.values(pools).flat();
    const allResp = Object.values(responses).flat();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(groups.map(g => ({ Title: g.title, Description: g.description, Members: g.memberList?.length || 0, Created: g.created }))), "Groups");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allPools.map(p => ({ Title: p.title, Description: p.description, Status: getPoolStatus(p), Start: p.startTime, End: p.endTime, Options: p.options.length }))), "Pools");
    if (allResp.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allResp.map(r => ({ Name: r.name, Email: r.email, Selected: r.selected, Allotted: r.allotted, Time: r.time }))), "Responses");
    XLSX.writeFile(wb, `allotment_export_${Date.now()}.xlsx`);
    addToast("Data exported as Excel");
  }

  async function clearCollection(collName) {
    try {
      const q = query(collection(db, collName));
      const snap = await getDocs(q);
      const batch = writeBatch(db);
      snap.docs.forEach(doc => batch.delete(doc.ref));
      await batch.commit();
      addToast(`All ${collName} cleared`);
    } catch(e) { addToast(e.message, "error"); }
  }

  return (
    <div style={{ maxWidth: 720, margin: "0 auto", padding: "32px 24px" }}>
      <h2 style={{ margin: "0 0 24px", fontSize: 22, fontWeight: 700, color: t.text }}>Settings</h2>
      <Section title="Appearance"><Row label="Dark Mode" desc="Switch between light and dark themes" action={<Toggle on={dark} toggle={() => setDark(!dark)} />} /></Section>
      <Section title="Data Management">
        <Row label="Export All Data" desc="Download groups, pools and responses as Excel" action={<Btn t={t} size="sm" variant="secondary" onClick={handleExport}>Export</Btn>} />
        <Row label="Clear Responses" desc="Remove all response records" action={<Btn t={t} size="sm" variant="danger" onClick={() => setConfirm({ title: "Clear All Responses", message: "This will permanently delete all response records. Continue?", action: () => clearCollection("responses") })}>Clear</Btn>} />
        <Row label="Clear All Pools" desc="Delete all pools across groups" action={<Btn t={t} size="sm" variant="danger" onClick={() => setConfirm({ title: "Clear All Pools", message: "This will permanently delete all pools and their data. Continue?", action: async () => { await clearCollection("pools"); await clearCollection("responses"); } })}>Clear</Btn>} />
      </Section>
      <Section title="About">
        <div style={{ fontSize: 13, color: t.textSec, lineHeight: 1.7 }}>
          <div style={{ fontWeight: 600 }}>Allotment App v3.0 (Firebase)</div>
          <div>First Come First Serve seat allocation platform</div>
          <div>Built with React + Vite + Firebase</div>
        </div>
      </Section>

      {confirm && <ConfirmModal t={t} title={confirm.title} message={confirm.message}
        onCancel={() => setConfirm(null)} onConfirm={() => { confirm.action(); setConfirm(null); }} />}
    </div>
  );
}

// ─── STATS PAGE ──────────────────────────────────────────────────────────────
function StatsPage({ t, groups, pools, responses }) {
  const allPools = Object.values(pools).flat();
  const allOptions = allPools.flatMap(p => p.options || []);
  const totalSeats = allOptions.reduce((s, o) => s + (o.seatLimit || 0), 0);
  const filledSeats = allOptions.reduce((s, o) => s + (o.filledSeats || 0), 0);
  const totalResponses = Object.values(responses).flat().length;
  const totalMembers = groups.reduce((s, g) => s + (g.memberList?.length || 0), 0);
  const inProgress = allPools.filter(p => getPoolStatus(p) === "In Progress").length;
  const completed = allPools.filter(p => getPoolStatus(p) === "Completed").length;

  const stats = [
    { label: "Groups", value: groups.length, icon: "users", color: G.blue },
    { label: "Total Pools", value: allPools.length, icon: "pool", color: G.green },
    { label: "In Progress", value: inProgress, icon: "clock", color: G.amber },
    { label: "Completed", value: completed, icon: "check", color: G.purple },
    { label: "Responses", value: totalResponses, icon: "chart", color: "#e11d48" },
    { label: "Members", value: totalMembers, icon: "user", color: G.greenL },
    { label: "Total Seats", value: totalSeats, icon: "lock", color: "#8b5cf6" },
    { label: "Filled", value: filledSeats, icon: "check", color: G.green },
  ];

  const BAR_MAX = Math.max(...allOptions.map(o => o.seatLimit || 0), 1);
  const CHART_H = 120;

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: "32px 24px" }}>
      <h2 style={{ margin: "0 0 24px", fontSize: 22, fontWeight: 700, color: t.text }}>Statistics</h2>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(140px,1fr))", gap: 14, marginBottom: 28 }}>
        {stats.map(s => (
          <Card key={s.label} t={t} style={{ padding: 18 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: `${s.color}22`, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 10 }}><Icon name={s.icon} size={18} color={s.color} /></div>
            <div style={{ fontSize: 26, fontWeight: 800, color: t.text }}>{s.value}</div>
            <div style={{ fontSize: 12, color: t.textSec, marginTop: 2 }}>{s.label}</div>
          </Card>
        ))}
      </div>

      <Card t={t} style={{ marginBottom: 20 }}>
        <h3 style={{ margin: "0 0 20px", fontSize: 15, fontWeight: 700, color: t.text }}>Seat Occupancy by Option</h3>
        {allOptions.length > 0 ? (
          <div style={{ display: "flex", gap: 8, alignItems: "flex-end", overflowX: "auto", paddingBottom: 8 }}>
            {allOptions.slice(0, 12).map(opt => {
              const p = pct(opt.filledSeats || 0, opt.seatLimit || 0);
              const barH = Math.max(((opt.filledSeats||0) / BAR_MAX) * CHART_H, 4);
              const emptyH = Math.max((((opt.seatLimit||0) - (opt.filledSeats||0)) / BAR_MAX) * CHART_H, 0);
              return (
                <div key={opt.id} style={{ flex: 1, minWidth: 50, display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                  <span style={{ fontSize: 11, color: t.textSec, fontWeight: 600 }}>{p}%</span>
                  <div style={{ width: "100%", height: CHART_H, display: "flex", flexDirection: "column", justifyContent: "flex-end" }}>
                    <div style={{ height: emptyH, background: t.surface2, borderRadius: "4px 4px 0 0", transition: "height .4s" }} />
                    <div style={{ height: barH, background: `linear-gradient(180deg,${G.greenL},${G.green})`, borderRadius: "4px 4px 0 0", transition: "height .4s" }} />
                  </div>
                  <span style={{ fontSize: 10, color: t.textMut, textAlign: "center", maxWidth: 55, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{opt.name}</span>
                </div>
              );
            })}
          </div>
        ) : <div style={{ textAlign: "center", color: t.textMut, padding: 40 }}>No data yet</div>}
      </Card>

      <Card t={t}>
        <h3 style={{ margin: "0 0 16px", fontSize: 15, fontWeight: 700, color: t.text }}>Overall Fill Rate</h3>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 14, marginBottom: 8 }}>
          <span style={{ color: t.textSec }}>Seats filled</span>
          <span style={{ color: t.text, fontWeight: 600 }}>{filledSeats} / {totalSeats}</span>
        </div>
        <ProgressBar value={totalSeats ? pct(filledSeats, totalSeats) : 0} t={t} />
        <div style={{ fontSize: 12, color: t.textMut, marginTop: 6 }}>{totalSeats ? pct(filledSeats, totalSeats) : 0}% capacity utilization</div>
      </Card>
    </div>
  );
}

// ─── ROOT APP ────────────────────────────────────────────────────────────────
export default function App() {
  const [dark, setDark] = useState(true);
  const t = dark ? darkTheme : lightTheme;

  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [page, setPage] = useState("home");
  const [groups, setGroups] = useState([]);
  const [pools, setPools] = useState({});
  const [responses, setResponses] = useState({});
  const [selectedGroup, setSelectedGroup] = useState(null);
  const [selectedPool, setSelectedPool] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [notifications, setNotifications] = useState([{ id: 1, message: "Welcome to Allotment!", time: Date.now(), read: false }]);
  const [, setTick] = useState(0);

  // Firebase Auth State
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (u) => {
      if (u) {
        const uDoc = await getDoc(doc(db, "users", u.uid));
        if (uDoc.exists()) setUser({ uid: u.uid, email: u.email, ...uDoc.data() });
        else setUser({ uid: u.uid, email: u.email, name: u.displayName || "User", role: "user" });
      } else {
        setUser(null);
      }
      setAuthLoading(false);
    });
    return () => unsub();
  }, []);

  // Firebase Firestore Listeners
  useEffect(() => {
    if (!user) return;
    
    const unsubGroups = onSnapshot(collection(db, "groups"), snap => {
      setGroups(snap.docs.map(d => ({ id: d.id, ...d.data() })));
    });
    
    const unsubPools = onSnapshot(collection(db, "pools"), snap => {
      const pObj = {};
      snap.docs.forEach(d => {
        const p = { id: d.id, ...d.data() };
        if (!pObj[p.groupId]) pObj[p.groupId] = [];
        pObj[p.groupId].push(p);
      });
      setPools(pObj);
    });
    
    const unsubResp = onSnapshot(collection(db, "responses"), snap => {
      const rObj = {};
      snap.docs.forEach(d => {
        const r = { id: d.id, ...d.data() };
        if (!rObj[r.poolId]) rObj[r.poolId] = [];
        rObj[r.poolId].push(r);
      });
      setResponses(rObj);
    });

    return () => { unsubGroups(); unsubPools(); unsubResp(); };
  }, [user]);

  // Real-time tick for status updates
  useEffect(() => {
    const iv = setInterval(() => setTick(tick => tick + 1), 10000);
    return () => clearInterval(iv);
  }, []);

  const addToast = useCallback((message, type = "success") => {
    const id = Date.now() + Math.random();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => setToasts(prev => prev.filter(toast => toast.id !== id)), 3500);
  }, []);

  const addNotif = useCallback((message) => {
    setNotifications(prev => [{ id: Date.now(), message, time: Date.now(), read: false }, ...prev.slice(0, 19)]);
  }, []);

  function markRead() { setNotifications(prev => prev.map(n => ({ ...n, read: true }))); }
  
  async function handleLogout() { 
    try {
      await signOut(auth);
      setPage("home");
      addToast("Logged out successfully"); 
    } catch(e) { addToast("Failed to log out", "error"); }
  }

  // Keep selectedGroup synced with groups state
  const currentGroup = selectedGroup ? groups.find(g => g.id === selectedGroup.id) || selectedGroup : null;

  if (authLoading) return <div style={{ minHeight: "100vh", background: t.bg, display: "flex", alignItems: "center", justifyContent: "center", color: t.textSec }}>Loading...</div>;
  if (!user) return <AuthPage t={t} />;

  return (
    <div style={{ minHeight: "100vh", background: t.bg, color: t.text, fontFamily: "'Inter','Google Sans',Roboto,Arial,sans-serif", transition: "background .3s,color .3s" }}>
      <ToastContainer toasts={toasts} />
      <NavBar t={t} dark={dark} setDark={setDark} page={page} setPage={setPage} user={user} notifications={notifications} onMarkRead={markRead} onLogout={handleLogout} />

      <main>
        {page === "home" && <HomePage t={t} groups={groups} setPage={setPage} setSelectedGroup={setSelectedGroup} addToast={addToast} addNotif={addNotif} />}
        {page === "group" && currentGroup && <GroupPage t={t} group={currentGroup} pools={pools} setPage={setPage} setSelectedPool={setSelectedPool} addToast={addToast} addNotif={addNotif} />}
        {page === "createPool" && <CreatePoolPage t={t} pool={selectedPool} setPage={setPage} selectedGroup={currentGroup} addToast={addToast} addNotif={addNotif} />}
        {page === "livePool" && selectedPool && <LivePoolPage t={t} pool={selectedPool} setPage={setPage} pools={pools} selectedGroup={currentGroup} responses={responses} user={user} addToast={addToast} />}
        {page === "responses" && <ResponsesPage t={t} pool={selectedPool} setPage={setPage} responses={responses} />}
        {page === "profile" && <ProfilePage t={t} user={user} setPage={setPage} addToast={addToast} />}
        {page === "stats" && <StatsPage t={t} groups={groups} pools={pools} responses={responses} />}
        {page === "settings" && <SettingsPage t={t} dark={dark} setDark={setDark} groups={groups} pools={pools} responses={responses} addToast={addToast} />}
      </main>
    </div>
  );
}
