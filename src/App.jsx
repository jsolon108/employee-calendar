import { useState, useMemo, useEffect } from "react";
import { supabase } from "./supabase";
import { useMsal } from "@azure/msal-react";
import { loginRequest, EDITOR_GROUP_ID } from "./authConfig";
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";

const BRANCH_GROUPS = {
  "── New York ──": ["Baldwin", "Bohemia", "Brooklyn", "Farmingdale", "Manhattan", "New Hyde Park"],
  "── Connecticut ──": ["Hartford", "Milford", "Stamford"],
  "New York": ["Baldwin", "Bohemia", "Brooklyn", "Farmingdale", "Manhattan", "New Hyde Park"],
  "Connecticut": ["Hartford", "Milford", "Stamford"],
};

const BRANCH_STATE = {
  "Baldwin": "NY", "Bohemia": "NY", "Brooklyn": "NY",
  "Farmingdale": "NY", "Manhattan": "NY", "New Hyde Park": "NY",
  "Hartford": "CT", "Milford": "CT", "Stamford": "CT",
  "New York": "NY", "Connecticut": "CT",
};

const BRANCH_COLORS = {
  "Baldwin":        { bg: "#FEE2E2", color: "#991B1B", dot: "#EF4444" },
  "Bohemia":        { bg: "#DCFCE7", color: "#166534", dot: "#22C55E" },
  "Brooklyn":       { bg: "#FDF4FF", color: "#7E22CE", dot: "#A855F7" },
  "Farmingdale":    { bg: "#DBEAFE", color: "#1D4ED8", dot: "#3B82F6" },
  "Hartford":       { bg: "#FCE7F3", color: "#9D174D", dot: "#EC4899" },
  "Manhattan":      { bg: "#E2E8F0", color: "#1E293B", dot: "#475569" },
  "Milford":        { bg: "#FEF9C3", color: "#713F12", dot: "#EAB308" },
  "New Hyde Park":  { bg: "#FFF7ED", color: "#9A3412", dot: "#F97316" },
  "Stamford":       { bg: "#CCFBF1", color: "#115E59", dot: "#14B8A6" },
  "New York":       { bg: "#EFF6FF", color: "#1D4ED8", dot: "#93C5FD" },
  "Connecticut":    { bg: "#F0FDF4", color: "#166534", dot: "#86EFAC" },
};

const EVENT_TYPES = ["Out of Office", "Half Day", "Coming in Late", "Leaving Early", "Training", "Counter Day", "Customer/Jobsite Visit", "Company Event", "Branch Event", "Holiday"];
const EVENT_TYPE_CONFIG = {
  "Out of Office":   { bg: "#FEE2E2", color: "#991B1B" },
  "Half Day":        { bg: "#FEF3C7", color: "#92400E" },
  "Coming in Late":  { bg: "#FEF3C7", color: "#92400E" },
  "Leaving Early":   { bg: "#FEF3C7", color: "#92400E" },
  "Training":        { bg: "#EDE9FE", color: "#5B21B6" },
  "Customer/Jobsite Visit": { bg: "#D1FAE5", color: "#065F46" },
  "Company Event":   { bg: "#0F172A", color: "#fff" },
  "Branch Event":    { bg: "#0369A1", color: "#fff" },
  "Holiday":         { bg: "#1E3A5F", color: "#fff" },
};

const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const DAYS_OF_WEEK = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
const DAYS_SHORT = ["Su","Mo","Tu","We","Th","Fr","Sa"];

function formatDate(d) {
  if (!d) return "—";
  return new Date(d + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

// ─── Login Screen ─────────────────────────────────────────────────────────────
function LoginScreen({ onMicrosoftLogin, onLogin }) {
  const [showLocal, setShowLocal] = useState(false);
  const [localUser, setLocalUser] = useState("");
  const [localPass, setLocalPass] = useState("");
  const [localError, setLocalError] = useState("");
  const LOCAL_ADMIN = { username: "admin", password: "Johnstone2024!" };
  const handleLocalLogin = () => {
    if (localUser === LOCAL_ADMIN.username && localPass === LOCAL_ADMIN.password) {
      onLogin({ id: 0, name: "Admin", email: "admin@local", role: "editor", branch: null, avatar: "AD" });
    } else { setLocalError("Incorrect username or password."); }
  };
  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0F172A 0%, #1E3A5F 60%, #1a3a6b 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'DM Sans', sans-serif", padding: 24 }}>
      <div style={{ width: "100%", maxWidth: 440 }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <img src="/logo.png" alt="Johnstone Supply" style={{ height: 80, width: "auto", margin: "0 auto 16px", display: "block" }} />
          <div style={{ color: "#fff", fontWeight: 800, fontSize: 26 }}>Johnstone Employee Calendar</div>
          <div style={{ color: "#94A3B8", fontSize: 12, letterSpacing: "0.1em", textTransform: "uppercase", marginTop: 4 }}>B&F Johnstone Supply</div>
        </div>
        <div style={{ background: "#fff", borderRadius: 20, overflow: "hidden", boxShadow: "0 32px 80px rgba(0,0,0,0.35)" }}>
          {!showLocal ? (
            <div style={{ padding: "40px 40px 36px" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: "#0F172A", marginBottom: 6 }}>Welcome</div>
              <div style={{ fontSize: 14, color: "#64748B", marginBottom: 32 }}>Sign in to view the employee calendar.</div>
              <button onClick={onMicrosoftLogin} style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 12, padding: "13px 20px", borderRadius: 10, border: "1.5px solid #E2E8F0", background: "#fff", color: "#0F172A", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                <svg width="20" height="20" viewBox="0 0 21 21" fill="none"><rect x="1" y="1" width="9" height="9" fill="#F25022"/><rect x="11" y="1" width="9" height="9" fill="#7FBA00"/><rect x="1" y="11" width="9" height="9" fill="#00A4EF"/><rect x="11" y="11" width="9" height="9" fill="#FFB900"/></svg>
                Sign in with Microsoft
              </button>
              <div style={{ textAlign: "center", marginTop: 24, color: "#94A3B8", fontSize: 12 }}>Microsoft Entra ID (Azure AD) · Single Sign-On</div>
              <div style={{ textAlign: "center", marginTop: 12 }}>
                <button onClick={() => setShowLocal(true)} style={{ background: "none", border: "none", color: "#CBD5E1", fontSize: 11, cursor: "pointer", fontFamily: "inherit" }}>Emergency access</button>
              </div>
            </div>
          ) : (
            <div style={{ padding: "36px 40px" }}>
              <div style={{ fontSize: 18, fontWeight: 700, color: "#0F172A", marginBottom: 4 }}>Emergency Access</div>
              <div style={{ fontSize: 13, color: "#94A3B8", marginBottom: 24 }}>For authorized administrators only.</div>
              <div style={{ marginBottom: 14 }}>
                <label style={{ display: "block", fontSize: 11, fontWeight: 600, color: "#64748B", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 }}>Username</label>
                <input value={localUser} onChange={e => { setLocalUser(e.target.value); setLocalError(""); }} placeholder="Username" style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1.5px solid #E2E8F0", fontSize: 14, outline: "none", boxSizing: "border-box", fontFamily: "inherit" }} />
              </div>
              <div style={{ marginBottom: 6 }}>
                <label style={{ display: "block", fontSize: 11, fontWeight: 600, color: "#64748B", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 }}>Password</label>
                <input type="password" value={localPass} onChange={e => { setLocalPass(e.target.value); setLocalError(""); }} onKeyDown={e => e.key === "Enter" && handleLocalLogin()} placeholder="Password" style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: `1.5px solid ${localError ? "#DC2626" : "#E2E8F0"}`, fontSize: 14, outline: "none", boxSizing: "border-box", fontFamily: "inherit" }} />
                {localError && <div style={{ fontSize: 11, color: "#DC2626", marginTop: 4, fontWeight: 600 }}>⚠ {localError}</div>}
              </div>
              <button onClick={handleLocalLogin} style={{ width: "100%", padding: "12px", borderRadius: 10, border: "none", background: "linear-gradient(135deg, #0F172A, #1E3A5F)", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", marginTop: 16 }}>Sign In</button>
              <div style={{ textAlign: "center", marginTop: 16 }}>
                <button onClick={() => { setShowLocal(false); setLocalError(""); setLocalUser(""); setLocalPass(""); }} style={{ background: "none", border: "none", color: "#94A3B8", fontSize: 12, cursor: "pointer", fontFamily: "inherit" }}>← Back</button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Manage Employees Modal ───────────────────────────────────────────────────
function ManageEmployeesModal({ employees, onClose, onAdd, onDelete }) {
  const [newName, setNewName] = useState("");
  const [search, setSearch] = useState("");
  const [error, setError] = useState("");

  const filtered = employees.filter(e => e.name.toLowerCase().includes(search.toLowerCase()));

  const handleAdd = () => {
    const trimmed = newName.trim();
    if (!trimmed) { setError("Please enter a name."); return; }
    if (employees.some(e => e.name.toLowerCase() === trimmed.toLowerCase())) { setError("This employee already exists."); return; }
    onAdd(trimmed);
    setNewName("");
    setError("");
  };

  const inputStyle = { width: "100%", padding: "10px 12px", borderRadius: 8, border: "1.5px solid #E2E8F0", fontSize: 14, color: "#0F172A", outline: "none", boxSizing: "border-box", fontFamily: "inherit", background: "#fff" };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", backdropFilter: "blur(4px)", zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 16, width: "100%", maxWidth: 500, maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 25px 60px rgba(0,0,0,0.18)", fontFamily: "'DM Sans', sans-serif", overflow: "hidden" }} onClick={e => e.stopPropagation()}>
        <div style={{ background: "linear-gradient(135deg, #0F172A 0%, #1E3A5F 100%)", padding: "22px 28px", flexShrink: 0 }}>
          <div style={{ fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase", color: "#94A3B8", marginBottom: 4 }}>Settings</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#fff" }}>Manage Employees</div>
        </div>
        <div style={{ padding: "20px 28px", borderBottom: "1px solid #F1F5F9", flexShrink: 0 }}>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              value={newName}
              onChange={e => { setNewName(e.target.value); setError(""); }}
              onKeyDown={e => e.key === "Enter" && handleAdd()}
              placeholder="Add new employee name..."
              style={{ ...inputStyle, flex: 1 }}
            />
            <button onClick={handleAdd} style={{ padding: "10px 20px", borderRadius: 8, border: "none", background: "linear-gradient(135deg, #0F172A, #1E3A5F)", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>Add</button>
          </div>
          {error && <div style={{ fontSize: 11, color: "#DC2626", marginTop: 6, fontWeight: 600 }}>⚠ {error}</div>}
        </div>
        <div style={{ padding: "12px 28px", borderBottom: "1px solid #F1F5F9", flexShrink: 0 }}>
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder={`Search ${employees.length} employees...`} style={inputStyle} />
        </div>
        <div style={{ overflowY: "auto", flex: 1 }}>
          {filtered.map(emp => (
            <div key={emp.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 28px", borderBottom: "1px solid #F8FAFC" }}>
              <span style={{ fontSize: 14, color: "#0F172A" }}>{emp.name}</span>
              <button onClick={() => onDelete(emp.id)} style={{ padding: "4px 10px", borderRadius: 6, border: "1.5px solid #FEE2E2", background: "#FEF2F2", color: "#DC2626", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Remove</button>
            </div>
          ))}
          {filtered.length === 0 && <div style={{ padding: "24px 28px", color: "#94A3B8", fontSize: 14 }}>No employees found.</div>}
        </div>
        <div style={{ padding: "16px 28px", borderTop: "1px solid #F1F5F9", flexShrink: 0 }}>
          <button onClick={onClose} style={{ padding: "10px 24px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#64748B", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Close</button>
        </div>
      </div>
    </div>
  );
}

// ─── Event Modal ──────────────────────────────────────────────────────────────
function EventModal({ event, onClose, onSave, onDelete, isNew, employees = [] }) {
  const blank = { employeeName: "", branch: "Farmingdale", eventType: "Out of Office", startDate: "", endDate: "", notes: "" };
  const [form, setForm] = useState(event || blank);
  const [errors, setErrors] = useState({});
  const set = (k, v) => { setForm(f => ({ ...f, [k]: v })); setErrors(e => ({ ...e, [k]: null })); };

  const inputStyle = { width: "100%", padding: "10px 12px", borderRadius: 8, border: "1.5px solid #E2E8F0", fontSize: 14, color: "#0F172A", outline: "none", boxSizing: "border-box", fontFamily: "inherit", background: "#fff" };
  const labelStyle = { display: "block", fontSize: 11, fontWeight: 600, color: "#64748B", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 };
  const errStyle = { fontSize: 11, color: "#DC2626", marginTop: 4, fontWeight: 600 };
  const req = <span style={{ color: "#DC2626" }}> *</span>;
  const borderErr = (k) => ({ ...inputStyle, borderColor: errors[k] ? "#DC2626" : "#E2E8F0" });

  const validate = () => {
    const e = {};
    if (!form.employeeName.trim()) e.employeeName = "Required";
    if (!form.startDate) e.startDate = "Required";
    if (!form.endDate) e.endDate = "Required";
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const handleSave = () => { if (validate()) onSave(form); };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", backdropFilter: "blur(4px)", zIndex: 1100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 16, width: "100%", maxWidth: 520, maxHeight: "90vh", display: "flex", flexDirection: "column", boxShadow: "0 25px 60px rgba(0,0,0,0.18)", fontFamily: "'DM Sans', sans-serif", overflow: "hidden" }} onClick={e => e.stopPropagation()}>
        <div style={{ background: "linear-gradient(135deg, #0F172A 0%, #1E3A5F 100%)", padding: "22px 28px", flexShrink: 0 }}>
          <div style={{ fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase", color: "#94A3B8", marginBottom: 4 }}>{isNew ? "New Event" : "Editing Event"}</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#fff" }}>{isNew ? "Add Event" : "Edit Event"}</div>
        </div>
        <div style={{ overflowY: "auto", flex: 1, padding: "24px 28px", display: "flex", flexDirection: "column", gap: 16 }}>
          <div>
            <label style={labelStyle}>Employee Name / Event Name{req}</label>
            <input list="employee-list" value={form.employeeName} onChange={e => set("employeeName", e.target.value)} placeholder="e.g. John Smith" style={borderErr("employeeName")} autoComplete="off" />
            <datalist id="employee-list">
              {employees.map(emp => <option key={emp.id} value={emp.name} />)}
            </datalist>
            {errors.employeeName && <div style={errStyle}>⚠ {errors.employeeName}</div>}
          </div>
          <div>
            <label style={labelStyle}>Branch{req}</label>
            <select value={form.branch} onChange={e => set("branch", e.target.value)} style={inputStyle}>
              <option value="All Branches">All Branches</option>
              <optgroup label="── Groups ──">
                <option value="New York">🗺 New York (All NY Branches)</option>
                <option value="Connecticut">🗺 Connecticut (All CT Branches)</option>
              </optgroup>
              <optgroup label="── New York ──">
                <option value="Baldwin">Baldwin</option>
                <option value="Bohemia">Bohemia</option>
                <option value="Brooklyn">Brooklyn</option>
                <option value="Farmingdale">Farmingdale</option>
                <option value="Manhattan">Manhattan</option>
                <option value="New Hyde Park">New Hyde Park</option>
              </optgroup>
              <optgroup label="── Connecticut ──">
                <option value="Hartford">Hartford</option>
                <option value="Milford">Milford</option>
                <option value="Stamford">Stamford</option>
              </optgroup>
            </select>
          </div>
          <div>
            <label style={labelStyle}>Event Type{req}</label>
            <select value={form.eventType} onChange={e => set("eventType", e.target.value)} style={inputStyle}>
              {EVENT_TYPES.map(t => <option key={t}>{t}</option>)}
            </select>
          </div>
          <div>
            <label style={labelStyle}>Date Range{req}</label>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <DatePicker
                selectsRange
                startDate={form.startDate ? new Date(form.startDate + "T12:00:00") : null}
                endDate={form.endDate ? new Date(form.endDate + "T12:00:00") : null}
                onChange={([start, end]) => {
                  set("startDate", start ? start.toISOString().split("T")[0] : "");
                  set("endDate", end ? end.toISOString().split("T")[0] : "");
                }}
                placeholderText="Select date range"
                wrapperClassName="date-range-wrapper"
                className="date-range-input"
              />
            </div>
            {errors.startDate && <div style={errStyle}>⚠ {errors.startDate}</div>}
            {errors.endDate && <div style={errStyle}>⚠ {errors.endDate}</div>}
          </div>
          <div>
            <label style={labelStyle}>Notes / Time of Arrival or Departure</label>
            <textarea value={form.notes} onChange={e => set("notes", e.target.value)} rows={3} placeholder="Any additional notes..." style={{ ...inputStyle, resize: "vertical" }} />
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <input type="checkbox" id="noTimeOff" checked={form.noTimeOff || false} onChange={e => set("noTimeOff", e.target.checked)} style={{ width: 16, height: 16, cursor: "pointer" }} />
            <label htmlFor="noTimeOff" style={{ ...labelStyle, margin: 0, cursor: "pointer" }}>No time off requests</label>
          </div>
        </div>
        <div style={{ padding: "16px 28px", borderTop: "1px solid #F1F5F9", display: "flex", gap: 10, justifyContent: "flex-end", flexShrink: 0, background: "#fff" }}>
          {!isNew && onDelete && (
            <button onClick={() => onDelete(form.id)} style={{ padding: "10px 20px", borderRadius: 8, border: "1.5px solid #FEE2E2", background: "#FEF2F2", color: "#DC2626", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", marginRight: "auto" }}>Delete Event</button>
          )}
          <button onClick={onClose} style={{ padding: "10px 20px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#64748B", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Cancel</button>
          <button onClick={handleSave} style={{ padding: "10px 24px", borderRadius: 8, border: "none", background: "linear-gradient(135deg, #0F172A, #1E3A5F)", color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{isNew ? "Add Event" : "Save Changes"}</button>
        </div>
      </div>
    </div>
  );
}

// ─── Detail Modal ─────────────────────────────────────────────────────────────
function DetailModal({ event: ev, onClose, onEdit, canEdit }) {
  const branchCfg = BRANCH_COLORS[ev.branch] || { bg: "#F1F5F9", color: "#475569", dot: "#94A3B8" };
  const typeCfg = EVENT_TYPE_CONFIG[ev.eventType] || { bg: "#F1F5F9", color: "#475569" };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", backdropFilter: "blur(4px)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 16, width: "100%", maxWidth: 460, boxShadow: "0 25px 60px rgba(0,0,0,0.2)", overflow: "hidden", fontFamily: "'DM Sans', sans-serif" }} onClick={e => e.stopPropagation()}>
        <div style={{ background: "linear-gradient(135deg, #0F172A 0%, #1E3A5F 100%)", padding: "22px 28px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 11, letterSpacing: "0.12em", textTransform: "uppercase", color: "#94A3B8", marginBottom: 4 }}>{ev.branch}</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: "#fff" }}>{ev.employeeName}</div>
          </div>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 8, color: "#fff", fontSize: 20, cursor: "pointer", width: 32, height: 32, display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
        </div>
        <div style={{ padding: "24px 28px" }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "3px 10px", borderRadius: 20, background: branchCfg.bg, color: branchCfg.color, fontSize: 12, fontWeight: 600 }}>
              <span style={{ width: 7, height: 7, borderRadius: "50%", background: branchCfg.dot }} />{ev.branch}
            </span>
            <span style={{ padding: "3px 10px", borderRadius: 20, background: typeCfg.bg, color: typeCfg.color, fontSize: 12, fontWeight: 600 }}>{ev.eventType}</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <div>
              <div style={{ fontSize: 11, color: "#94A3B8", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>Start Date</div>
              <div style={{ fontSize: 13, color: "#0F172A", fontWeight: 600 }}>{formatDate(ev.startDate)}</div>
            </div>
            <div>
              <div style={{ fontSize: 11, color: "#94A3B8", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 4 }}>End Date</div>
              <div style={{ fontSize: 13, color: "#0F172A", fontWeight: 600 }}>{formatDate(ev.endDate)}</div>
            </div>
          </div>
          {ev.notes && <div style={{ padding: "10px 14px", background: "#F8FAFC", borderRadius: 8, fontSize: 13, color: "#475569", marginBottom: 16 }}><span style={{ fontWeight: 600, color: "#64748B" }}>Notes: </span>{ev.notes}</div>}
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <button onClick={onClose} style={{ padding: "9px 20px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#64748B", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Close</button>
            {canEdit && <button onClick={() => onEdit(ev)} style={{ padding: "9px 20px", borderRadius: 8, border: "none", background: "linear-gradient(135deg, #0F172A, #1E3A5F)", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Edit</button>}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Calendar ─────────────────────────────────────────────────────────────────
function CalendarView({ events, onSelectEvent, branch }) {
  const today = new Date();
  const [calYear, setCalYear] = useState(today.getFullYear());
  const [calMonth, setCalMonth] = useState(today.getMonth());
  const [calMode, setCalMode] = useState("month");
  const [weekOffset, setWeekOffset] = useState(0);
  const [expandedDay, setExpandedDay] = useState(null);

  const firstDay = new Date(calYear, calMonth, 1).getDay();
  const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
  const totalCells = Math.ceil((firstDay + daysInMonth) / 7) * 7;

  const getWeekStart = (offset) => {
    const d = new Date(today);
    d.setDate(today.getDate() - today.getDay() + offset * 7);
    d.setHours(0, 0, 0, 0);
    return d;
  };
  const weekStart = getWeekStart(weekOffset);
  const weekDays = Array.from({ length: 7 }, (_, i) => { const d = new Date(weekStart); d.setDate(weekStart.getDate() + i); return d; });
  const weekLabel = () => {
    const s = weekDays[0], e = weekDays[6];
    if (s.getMonth() === e.getMonth()) return `${MONTHS[s.getMonth()]} ${s.getDate()}–${e.getDate()}, ${s.getFullYear()}`;
    return `${MONTHS_SHORT[s.getMonth()]} ${s.getDate()} – ${MONTHS_SHORT[e.getMonth()]} ${e.getDate()}, ${e.getFullYear()}`;
  };

  const toDateStr = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  const eventsForDate = (dateStr) => {
    const filtered = events.filter(ev => ev.startDate <= dateStr && ev.endDate >= dateStr);
    return filtered.sort((a, b) => {
      const typeOrder = { "Company Event": 0, "Holiday": 1, "Training": 2, "Counter Day": 3, "Branch Event": 4 };
      const aTypeOrder = typeOrder[a.eventType] ?? 99;
      const bTypeOrder = typeOrder[b.eventType] ?? 99;
      if (aTypeOrder !== bTypeOrder) return aTypeOrder - bTypeOrder;
      const branchOrder = ["Baldwin","Bohemia","Brooklyn","Farmingdale","Manhattan","New Hyde Park","Hartford","Milford","Stamford","New York","Connecticut","All Branches"];
      const aBranch = branchOrder.indexOf(a.branch) === -1 ? 99 : branchOrder.indexOf(a.branch);
      const bBranch = branchOrder.indexOf(b.branch) === -1 ? 99 : branchOrder.indexOf(b.branch);
      if (aBranch !== bBranch) return aBranch - bBranch;
      const firstName = name => name.trim().split(" ")[0].toLowerCase();
      return firstName(a.employeeName).localeCompare(firstName(b.employeeName));
    });
  };

  const prevMonth = () => { if (calMonth === 0) { setCalMonth(11); setCalYear(y => y - 1); } else setCalMonth(m => m - 1); };
  const nextMonth = () => { if (calMonth === 11) { setCalMonth(0); setCalYear(y => y + 1); } else setCalMonth(m => m + 1); };

  const EventChip = ({ ev }) => {
    const branchCfg = BRANCH_COLORS[ev.branch] || { bg: "#F1F5F9", color: "#475569" };
    const isCompanyWide = ev.eventType === "Company Event" || ev.eventType === "Holiday";
    const isBranchEvent = ev.eventType === "Branch Event";
    const isTraining = ev.eventType === "Training";
    const isCounterDay = ev.eventType === "Counter Day";
    const isGroupEvent = ev.branch === "New York" || ev.branch === "Connecticut";
    const cfg = isCompanyWide ? EVENT_TYPE_CONFIG[ev.eventType] : branchCfg;
    const stateLabel = BRANCH_STATE[ev.branch] ? ` · ${BRANCH_STATE[ev.branch]}` : "";
    return (
      <button onClick={() => onSelectEvent(ev)} style={{ display: "block", width: "100%", textAlign: "left", padding: "3px 6px", borderRadius: 4, background: cfg.bg, color: cfg.color, border: "none", cursor: "pointer", fontFamily: "'DM Sans', sans-serif", lineHeight: 1.4, overflow: "hidden", marginBottom: 2 }}>
        <div style={{ fontSize: 10, fontWeight: 700, textOverflow: "ellipsis", whiteSpace: "nowrap", overflow: "hidden" }}>
          {isCompanyWide ? `📅 ${ev.employeeName}` : isGroupEvent ? `🗺 ${ev.branch}: ${ev.employeeName}` : isBranchEvent ? `📍 ${ev.branch}: ${ev.employeeName}` : isTraining ? `🎓 ${ev.employeeName}` : isCounterDay ? `🏪 ${ev.employeeName}` : `👤 ${ev.employeeName}`}
        </div>
        <div style={{ fontSize: 9, opacity: 0.75, textOverflow: "ellipsis", whiteSpace: "nowrap", overflow: "hidden" }}>
          {ev.noTimeOff ? "🚫 No time off requests" : isTraining ? `${ev.branch} — Training` : isCounterDay ? `${ev.branch} — Counter Day` : isCompanyWide ? ev.eventType : isBranchEvent ? "Branch Event" : isGroupEvent ? `${ev.eventType} · All ${ev.branch} Branches` : `${ev.eventType}${stateLabel}`}
        </div>
      </button>
    );
  };

  const monthCells = Array.from({ length: totalCells }, (_, i) => {
    const dayNum = i - firstDay + 1;
    if (dayNum < 1 || dayNum > daysInMonth) return null;
    const mm = String(calMonth + 1).padStart(2, "0"), dd = String(dayNum).padStart(2, "0");
    const dateStr = `${calYear}-${mm}-${dd}`;
    const isToday = today.getFullYear() === calYear && today.getMonth() === calMonth && today.getDate() === dayNum;
    return { dayNum, dateStr, isToday, events: eventsForDate(dateStr) };
  });

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <button onClick={calMode === "month" ? prevMonth : () => setWeekOffset(o => o - 1)} style={{ padding: "7px 14px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#475569", fontSize: 16, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>‹</button>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {calMode === "month" ? (
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <select value={calMonth} onChange={e => setCalMonth(Number(e.target.value))} style={{ appearance: "none", WebkitAppearance: "none", padding: "6px 28px 6px 12px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#0F172A", fontSize: 16, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", outline: "none" }}>
                {MONTHS.map((m, i) => <option key={m} value={i}>{m}</option>)}
              </select>
              <select value={calYear} onChange={e => setCalYear(Number(e.target.value))} style={{ appearance: "none", WebkitAppearance: "none", padding: "6px 28px 6px 12px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#0F172A", fontSize: 16, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", outline: "none" }}>
                {Array.from({ length: 10 }, (_, i) => today.getFullYear() - 7 + i).map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            </div>
          ) : (
            <div style={{ fontSize: 18, fontWeight: 700, color: "#0F172A" }}>{weekLabel()}</div>
          )}
          <div style={{ display: "flex", background: "#F1F5F9", borderRadius: 8, padding: 3, gap: 2 }}>
            {[["month", "Month"], ["week", "Week"]].map(([m, l]) => (
              <button key={m} onClick={() => setCalMode(m)} style={{ padding: "5px 12px", borderRadius: 6, border: "none", background: calMode === m ? "#fff" : "transparent", color: calMode === m ? "#0F172A" : "#94A3B8", fontSize: 12, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{l}</button>
            ))}
          </div>
        </div>
        <button onClick={calMode === "month" ? nextMonth : () => setWeekOffset(o => o + 1)} style={{ padding: "7px 14px", borderRadius: 8, border: "1.5px solid #E2E8F0", background: "#fff", color: "#475569", fontSize: 16, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>›</button>
      </div>

      {calMode === "month" && (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
            {DAYS_OF_WEEK.map(d => <div key={d} style={{ textAlign: "center", padding: "6px 0", fontSize: 11, fontWeight: 700, color: "#94A3B8", letterSpacing: "0.08em", textTransform: "uppercase" }}>{d}</div>)}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 }}>
            {monthCells.map((cell, i) => (
              <div key={i} style={{ minHeight: 200, background: cell === null ? "transparent" : cell.isToday ? "#F0F7FF" : "#FAFBFC", borderRadius: 8, border: cell === null ? "none" : cell.isToday ? "2px solid #3B82F6" : "1.5px solid #E2E8F0", padding: cell === null ? 0 : "6px 5px", boxSizing: "border-box", overflow: "hidden", position: "relative" }}>
                {cell !== null && (
                  <>
                    <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 4 }}>
                      <span style={{ width: 22, height: 22, borderRadius: "50%", background: cell.isToday ? "#3B82F6" : "transparent", color: cell.isToday ? "#fff" : "#94A3B8", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: cell.isToday ? 800 : 500 }}>{cell.dayNum}</span>
                    </div>
                    {cell.events.slice(0, 20).map(ev => <EventChip key={ev.id + cell.dateStr} ev={ev} />)}
                    {cell.events.length > 20 && (
                      <div onClick={() => setExpandedDay(cell.dateStr)} style={{ fontSize: 9, color: "#6366F1", fontWeight: 600, textAlign: "center", marginTop: 2, cursor: "pointer" }}>
                        +{cell.events.length - 20} more
                      </div>
                    )}
                  </>
                )}
                {expandedDay && (
                  <div onClick={() => setExpandedDay(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.08)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <div onClick={e => e.stopPropagation()} style={{ background: "white", borderRadius: 12, padding: 20, minWidth: 280, maxWidth: 360, maxHeight: "80vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
                      <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12 }}>
                        {new Date(expandedDay + "T12:00:00").toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric" })}
                      </div>
                      {(() => {
                        const dayCell = monthCells.find(c => c && c.dateStr === expandedDay);
                        return dayCell?.events.map(ev => <EventChip key={ev.id + expandedDay} ev={ev} />);
                      })()}
                      <div onClick={() => setExpandedDay(null)} style={{ marginTop: 14, textAlign: "center", fontSize: 12, color: "#94A3B8", cursor: "pointer" }}>Close</div>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </>
      )}

      {calMode === "week" && (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
            {weekDays.map((d, i) => {
              const isToday = d.toDateString() === today.toDateString();
              return (
                <div key={i} style={{ textAlign: "center", padding: "8px 4px" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#94A3B8", textTransform: "uppercase", letterSpacing: "0.06em" }}>{DAYS_SHORT[i]}</div>
                  <div style={{ marginTop: 4, width: 30, height: 30, borderRadius: "50%", background: isToday ? "#3B82F6" : "transparent", color: isToday ? "#fff" : "#0F172A", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: isToday ? 800 : 600, margin: "4px auto 0" }}>{d.getDate()}</div>
                </div>
              );
            })}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 }}>
            {weekDays.map((d, i) => {
              const isToday = d.toDateString() === today.toDateString();
              const dayEvents = eventsForDate(toDateStr(d));
              return (
                <div key={i} style={{ minHeight: 160, background: isToday ? "#F0F7FF" : "#FAFBFC", borderRadius: 10, border: isToday ? "2px solid #3B82F6" : "1.5px solid #E2E8F0", padding: "10px 7px", boxSizing: "border-box" }}>
                  {dayEvents.length === 0 ? <div style={{ fontSize: 11, color: "#E2E8F0", textAlign: "center", marginTop: 20 }}>—</div> : dayEvents.map(ev => <EventChip key={ev.id + toDateStr(d)} ev={ev} />)}
                </div>
              );
            })}
          </div>
        </>
      )}

      {(branch === "All Branches" || BRANCH_GROUPS[branch]) && (
        <div style={{ marginTop: 20, padding: "16px 20px", background: "#F8FAFC", borderRadius: 12, border: "1.5px solid #E2E8F0" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#64748B", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10 }}>Branch Colors</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
            {Object.entries(BRANCH_COLORS).filter(([name]) => !["New York", "Connecticut"].includes(name)).map(([name, cfg]) => (
              <div key={name} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 10, height: 10, borderRadius: 3, background: cfg.bg, border: `1.5px solid ${cfg.dot}`, flexShrink: 0 }} />
                <span style={{ fontSize: 11, color: "#64748B", fontWeight: 600 }}>{name}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function getUSHolidays(year) {
  const fixed = [
    { name: "New Year's Day", date: `${year}-01-01` },
    { name: "Independence Day", date: `${year}-07-04` },
    { name: "Christmas Day", date: `${year}-12-25` },
  ];
  const may = new Date(year, 4, 31);
  while (may.getDay() !== 1) may.setDate(may.getDate() - 1);
  fixed.push({ name: "Memorial Day", date: may.toISOString().slice(0, 10) });
  const sep = new Date(year, 8, 1);
  while (sep.getDay() !== 1) sep.setDate(sep.getDate() + 1);
  fixed.push({ name: "Labor Day", date: sep.toISOString().slice(0, 10) });
  const nov = new Date(year, 10, 1);
  while (nov.getDay() !== 4) nov.setDate(nov.getDate() + 1);
  nov.setDate(nov.getDate() + 21);
  fixed.push({ name: "Thanksgiving Day", date: nov.toISOString().slice(0, 10) });
  return fixed.map((h, i) => ({
    id: `HOLIDAY-${year}-${i}`,
    employeeName: h.name,
    branch: "All Branches",
    eventType: "Holiday",
    startDate: h.date,
    endDate: h.date,
    notes: "",
  }));
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [events, setEvents] = useState([]);
  const [employees, setEmployees] = useState([]);
  const [loading, setLoading] = useState(true);
  const [branch, setBranch] = useState("All Branches");
  const [editing, setEditing] = useState(null);
  const [adding, setAdding] = useState(false);
  const [viewing, setViewing] = useState(null);
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const [managingEmployees, setManagingEmployees] = useState(false);
  const [eventTypeFilter, setEventTypeFilter] = useState("All");

  const { instance } = useMsal();
  const canEdit = currentUser?.role === "editor";

  useEffect(() => {
    instance.handleRedirectPromise().then(result => {
      if (result) {
        const groups = result.idTokenClaims?.groups || [];
        const isEditor = groups.includes(EDITOR_GROUP_ID);
        setCurrentUser({
          id: result.account.localAccountId,
          name: result.account.name,
          email: result.account.username,
          role: isEditor ? "editor" : "viewer",
          avatar: result.account.name.split(" ").map(n => n[0]).join("").slice(0, 2).toUpperCase(),
        });
      } else {
        const accounts = instance.getAllAccounts();
        if (accounts.length > 0) {
          const account = accounts[0];
          instance.acquireTokenSilent({ account, scopes: [] }).then(silentResult => {
            const groups = silentResult.idTokenClaims?.groups || [];
            const isEditor = groups.includes(EDITOR_GROUP_ID);
            setCurrentUser({
              id: account.localAccountId,
              name: account.name,
              email: account.username,
              role: isEditor ? "editor" : "viewer",
              avatar: account.name.split(" ").map(n => n[0]).join("").slice(0, 2).toUpperCase(),
            });
          }).catch(() => {});
        }
      }
    }).catch(e => console.error(e));
  }, [instance]);

  useEffect(() => {
    if (!currentUser) return;
    const fetchEvents = async () => {
      setLoading(true);
      const from = new Date();
      from.setFullYear(from.getFullYear() - 2);
      const to = new Date();
      to.setFullYear(to.getFullYear() + 5);
      const fromStr = from.toISOString().split("T")[0];
      const toStr = to.toISOString().split("T")[0];
      let allData = [];
      let from2 = 0;
      while (true) {
        const { data: chunk } = await supabase.from("events").select("*").gte("startDate", fromStr).lte("startDate", toStr).order("startDate").range(from2, from2 + 999);
        if (!chunk || chunk.length === 0) break;
        allData = [...allData, ...chunk];
        if (chunk.length < 1000) break;
        from2 += 1000;
      }
      setEvents(allData);
      setLoading(false);
    };
    fetchEvents();
  }, [currentUser]);

  useEffect(() => {
    if (!currentUser) return;
    const fetchEmployees = async () => {
      const { data } = await supabase.from("employees").select("*").order("name");
      if (data) setEmployees(data);
    };
    fetchEmployees();
  }, [currentUser]);

  const handleAddEmployee = async (name) => {
    const newEmp = { id: `EMP-${Date.now()}`, name };
    await supabase.from("employees").insert([newEmp]);
    setEmployees(e => [...e, newEmp].sort((a, b) => a.name.localeCompare(b.name)));
  };

  const handleDeleteEmployee = async (id) => {
    if (!window.confirm("Remove this employee from the list?")) return;
    await supabase.from("employees").delete().eq("id", id);
    setEmployees(e => e.filter(emp => emp.id !== id));
  };

  const handleLogin = (a) => setCurrentUser(a);
  const handleLogout = () => {
    setCurrentUser(null);
    setUserMenuOpen(false);
    instance.logoutRedirect().catch(e => console.error(e));
  };
  const handleMicrosoftLogin = async () => {
    try { await instance.loginRedirect(loginRequest); }
    catch (e) { console.error(e); }
  };

  const filteredEvents = useMemo(() => {
    const today = new Date();
    const years = [today.getFullYear() - 1, today.getFullYear(), today.getFullYear() + 1];
    const holidays = years.flatMap(y => getUSHolidays(y));
    const allEvents = [...events, ...holidays];
    let branchFiltered;
    if (branch === "All Branches") {
      branchFiltered = allEvents;
     } else if (BRANCH_GROUPS[branch]) {
      branchFiltered = allEvents.filter(e =>
        BRANCH_GROUPS[branch].includes(e.branch) ||
        e.branch === branch ||
        e.eventType === "Company Event" ||
        e.eventType === "Holiday" ||
        e.eventType === "Training"
      );
    } else {
      branchFiltered = allEvents.filter(e =>
        e.branch === branch ||
        e.eventType === "Company Event" ||
        e.eventType === "Holiday" ||
        e.eventType === "Training" ||
        (e.branch === "New York" && BRANCH_GROUPS["New York"].includes(branch)) ||
        (e.branch === "Connecticut" && BRANCH_GROUPS["Connecticut"].includes(branch))
      );
    }
    return eventTypeFilter === "All" ? branchFiltered : branchFiltered.filter(e => e.eventType === eventTypeFilter);
  }, [events, branch, eventTypeFilter]);

  const handleSave = async (form) => {
    if (adding) {
      const newId = `EVT-${Date.now()}`;
      const newEvent = { ...form, id: newId };
      await supabase.from("events").insert([newEvent]);
      setEvents(e => [...e, newEvent].sort((a, b) => a.startDate.localeCompare(b.startDate)));
      setAdding(false);
    } else {
      await supabase.from("events").update(form).eq("id", form.id);
      setEvents(e => e.map(ev => ev.id === form.id ? form : ev));
      setEditing(null);
    }
  };

  const handleDelete = async (id) => {
    if (!window.confirm("Delete this event? This cannot be undone.")) return;
    await supabase.from("events").delete().eq("id", id);
    setEvents(e => e.filter(ev => ev.id !== id));
    setEditing(null);
    setViewing(null);
  };

  const handleEditFromDetail = (ev) => { setViewing(null); setEditing(ev); };

  if (!currentUser) return (
    <>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
      <LoginScreen onMicrosoftLogin={handleMicrosoftLogin} onLogin={handleLogin} />
    </>
  );

  return (
    <div style={{ minHeight: "100vh", background: "#F1F5F9", fontFamily: "'DM Sans', sans-serif" }}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
      <div style={{ background: "linear-gradient(135deg, #0F172A 0%, #1E3A5F 100%)", padding: "0 32px" }}>
        <div style={{ maxWidth: 1300, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 64 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img src="/logo.png" alt="Johnstone Supply" style={{ height: 44, width: "auto" }} />
            <div>
              <div style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>Johnstone Employee Calendar</div>
              <div style={{ color: "#94A3B8", fontSize: 11, letterSpacing: "0.06em" }}>B&F JOHNSTONE SUPPLY</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {canEdit && (
              <>
                <button onClick={() => setManagingEmployees(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 18px", borderRadius: 8, border: "1.5px solid rgba(255,255,255,0.2)", background: "rgba(255,255,255,0.1)", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                  👥 Employees
                </button>
                <button onClick={() => setAdding(true)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 18px", borderRadius: 8, border: "1.5px solid rgba(255,255,255,0.2)", background: "rgba(255,255,255,0.1)", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                  <span style={{ fontSize: 16, lineHeight: 1 }}>+</span> Add Event
                </button>
              </>
            )}
            <div style={{ position: "relative" }}>
              <button onClick={() => setUserMenuOpen(o => !o)} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 10px 5px 5px", borderRadius: 8, border: "1.5px solid rgba(255,255,255,0.2)", background: "rgba(255,255,255,0.1)", cursor: "pointer", fontFamily: "inherit" }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 11, fontWeight: 700 }}>{currentUser.avatar}</div>
                <div>
                  <div style={{ color: "#fff", fontSize: 12, fontWeight: 600 }}>{currentUser.name.split(" ")[0]}</div>
                  <div style={{ color: "#94A3B8", fontSize: 10, textTransform: "capitalize" }}>{currentUser.role}</div>
                </div>
                <span style={{ color: "rgba(255,255,255,0.5)", fontSize: 10 }}>▾</span>
              </button>
              {userMenuOpen && (
                <div style={{ position: "absolute", right: 0, top: "calc(100% + 8px)", background: "#fff", borderRadius: 12, boxShadow: "0 8px 30px rgba(0,0,0,0.15)", width: 220, zIndex: 200, overflow: "hidden" }}>
                  <div style={{ padding: "14px 16px", borderBottom: "1px solid #F1F5F9" }}>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "#0F172A" }}>{currentUser.name}</div>
                    <div style={{ fontSize: 11, color: "#94A3B8" }}>{currentUser.email}</div>
                    <div style={{ marginTop: 8, display: "inline-flex", padding: "2px 8px", borderRadius: 10, background: canEdit ? "#EFF6FF" : "#F1F5F9", color: canEdit ? "#2563EB" : "#64748B", fontSize: 10, fontWeight: 700, textTransform: "uppercase" }}>{canEdit ? "✏️ Editor" : "👁 Viewer"}</div>
                  </div>
                  <button onClick={handleLogout} style={{ width: "100%", padding: "12px 16px", border: "none", background: "none", textAlign: "left", fontSize: 13, color: "#DC2626", fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Sign out</button>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1300, margin: "0 auto", padding: "28px 32px" }} onClick={() => setUserMenuOpen(false)}>
        {!canEdit && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 16px", borderRadius: 10, background: "#F8FAFC", border: "1.5px solid #E2E8F0", marginBottom: 20, fontSize: 13, color: "#64748B" }}>
            👁 <span>You have <strong>view-only</strong> access.</span>
          </div>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 24, flexWrap: "wrap" }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: "#475569", textTransform: "uppercase", letterSpacing: "0.04em" }}>Branch</span>
          <div style={{ position: "relative" }}>
            <select value={branch} onChange={e => setBranch(e.target.value)} style={{ appearance: "none", WebkitAppearance: "none", padding: "9px 36px 9px 14px", borderRadius: 10, border: "1.5px solid #CBD5E1", background: "#fff", color: "#0F172A", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", outline: "none", minWidth: 200 }}>
              <option value="All Branches">All Branches</option>
              <optgroup label="── Groups ──">
                <option value="── New York ──">🗺 New York</option>
                <option value="── Connecticut ──">🗺 Connecticut</option>
              </optgroup>
              <optgroup label="── New York ──">
                {["Baldwin","Bohemia","Brooklyn","Farmingdale","Manhattan","New Hyde Park"].map(b => <option key={b} value={b}>{b}</option>)}
              </optgroup>
              <optgroup label="── Connecticut ──">
                {["Hartford","Milford","Stamford"].map(b => <option key={b} value={b}>{b}</option>)}
              </optgroup>
            </select>
            <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: "#64748B", fontSize: 12 }}>▾</span>
          </div>
          {branch !== "All Branches" && !BRANCH_GROUPS[branch] && (
            <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: BRANCH_COLORS[branch]?.bg || "#F1F5F9", color: BRANCH_COLORS[branch]?.color || "#475569", fontSize: 12, fontWeight: 700, border: `1.5px solid ${BRANCH_COLORS[branch]?.dot || "#94A3B8"}` }}>
              📍 {branch}
            </div>
          )}
          {BRANCH_GROUPS[branch] && (
            <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px", borderRadius: 20, background: "#EFF6FF", color: "#1D4ED8", fontSize: 12, fontWeight: 700, border: "1.5px solid #93C5FD" }}>
              🗺 {branch.replace(/──\s*/g, "").trim()}
            </div>
          )}
          <span style={{ fontSize: 11, fontWeight: 700, color: "#475569", textTransform: "uppercase", letterSpacing: "0.04em", marginLeft: 12 }}>Event Type</span>
          <div style={{ position: "relative" }}>
            <select value={eventTypeFilter} onChange={e => setEventTypeFilter(e.target.value)} style={{ appearance: "none", WebkitAppearance: "none", padding: "9px 36px 9px 14px", borderRadius: 10, border: "1.5px solid #CBD5E1", background: "#fff", color: "#0F172A", fontSize: 14, fontWeight: 700, cursor: "pointer", fontFamily: "inherit", outline: "none", minWidth: 200 }}>
              <option value="All">All Event Types</option>
              {EVENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
            </select>
            <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: "#64748B", fontSize: 12 }}>▾</span>
          </div>
        </div>
        {loading ? (
          <div style={{ textAlign: "center", padding: 48, color: "#94A3B8", fontSize: 14 }}>Loading events...</div>
        ) : (
          <div style={{ background: "#fff", borderRadius: 14, boxShadow: "0 1px 3px rgba(0,0,0,0.07)", padding: "24px 28px" }}>
            <CalendarView events={filteredEvents} onSelectEvent={setViewing} branch={branch} />
          </div>
        )}
      </div>

      {viewing && <DetailModal event={viewing} onClose={() => setViewing(null)} onEdit={handleEditFromDetail} canEdit={canEdit} />}
      {(editing || adding) && <EventModal event={editing} isNew={adding} onClose={() => { setEditing(null); setAdding(false); }} onSave={handleSave} onDelete={handleDelete} employees={employees} />}
      {managingEmployees && <ManageEmployeesModal employees={employees} onClose={() => setManagingEmployees(false)} onAdd={handleAddEmployee} onDelete={handleDeleteEmployee} />}
    </div>
  );
}
