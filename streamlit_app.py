import hashlib
import re
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# ─────────────────────────────────────────────
ONEDRIVE_FILE_URL = "https://emerson-my.sharepoint.com/:x:/p/savitri_lazarus/IQB7_WEjDxxfQZDKz88rVLHpASyvoQKl8XH61HiTWzkGANQ?e=gAlAOv"
# ─────────────────────────────────────────────

C = {
    "deep_blue": "#004B8D", "green": "#00573D", "navy": "#1B2552",
    "bright_blue": "#1DB1DE", "soft_green": "#7CCF8B", "teal": "#00AD7C",
    "light_blue": "#75D3EB", "gray": "#9FA1A4", "black": "#000000", "white": "#FFFFFF",
}
PALETTE = [C["deep_blue"], C["bright_blue"], C["teal"], C["soft_green"],
           C["navy"], C["light_blue"], C["green"], C["gray"]]
STATUS_COLORS = {
    "Delayed": "#C0392B", "At Risk": "#E67E22", "On Track": C["teal"],
    "Active": C["deep_blue"], "In Progress": C["bright_blue"],
    "Complete": C["soft_green"], "Completed": C["soft_green"],
    "Not Started": C["gray"], "Planning": C["light_blue"],
}

FUTURE_FIELDS = [
    ("Business Value",      ["Business Value","BusinessValue","Strategic Value"],
     "Compares strategic upside across projects to prioritise high-return work."),
    ("Financial Value",     ["Financial Value","FinancialValue","Fin Value"],
     "Quantifies revenue or cost-savings potential per project."),
    ("Dollars at Risk",     ["Dollars at Risk","DollarsAtRisk","$ at Risk","Risk Dollars"],
     "Estimates loss exposure if a project is delayed or cancelled."),
    ("Estimated Capacity",  ["Estimated Capacity","Capacity","FTE Capacity","Estimated FTE"],
     "Identifies where teams are overcommitted and where buffer exists."),
    ("Hard Deadline",       ["Hard Deadline","HardDeadline","Deadline Date","Due Date"],
     "Distinguishes movable work from immovable commitments."),
    ("Deadline Type",       ["Deadline Type","DeadlineType","Deadline Category"],
     "Classifies deadlines as regulatory, contractual, internal, or aspirational."),
    ("Dependency Criticality",["Dependency Criticality","DependencyCriticality","Criticality"],
     "Identifies projects that unblock downstream work."),
    ("Executive Sponsor",   ["Executive Sponsor","ExecSponsor","Sponsor","Executive Owner"],
     "Clarifies decision ownership and escalation path."),
    ("Confidence Level",    ["Confidence Level","Confidence","ConfidenceLevel"],
     "Flags uncertain projects that may need review or re-scoping."),
    ("Blocker Reason",      ["Blocker Reason","BlockerReason","Blocker","Block Reason"],
     "Separates delay caused by capacity, data, or dependency issues."),
]

st.set_page_config(page_title="RevOps Program Dashboard", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;}}
.main{{background-color:#F7F8FA;}}
.block-container{{padding:1.6rem 2.2rem 2rem 2.2rem;max-width:1440px;}}
.kpi-wrap{{background:{C["white"]};border-radius:12px;padding:18px 20px 14px 20px;
  border:1px solid #E8ECF0;box-shadow:0 2px 8px rgba(0,0,0,0.05);
  min-height:108px;display:flex;flex-direction:column;justify-content:space-between;}}
.kpi-wrap.ph{{opacity:0.5;}}
.kpi-label{{font-size:10px;font-weight:700;letter-spacing:0.07em;text-transform:uppercase;
  color:{C["gray"]};margin-bottom:2px;}}
.kpi-value{{font-size:32px;font-weight:700;color:{C["navy"]};line-height:1;}}
.kpi-value.danger{{color:#C0392B;}} .kpi-value.success{{color:{C["teal"]};}}
.kpi-value.warn{{color:#D97706;}} .kpi-value.ph{{color:{C["gray"]};font-size:13px;font-weight:400;margin-top:6px;}}
.kpi-sub{{font-size:10px;color:{C["gray"]};margin-top:4px;}}
.kpi-accent-bar{{height:3px;border-radius:2px;margin-bottom:8px;}}
.section-title{{font-size:12px;font-weight:700;letter-spacing:0.09em;text-transform:uppercase;
  color:{C["navy"]};margin-bottom:14px;padding-bottom:7px;
  border-bottom:2px solid {C["deep_blue"]};display:inline-block;}}
.section-divider{{border:none;border-top:1px solid #E8ECF0;margin:28px 0 24px 0;}}
.exec-header{{background:linear-gradient(135deg,{C["navy"]} 0%,{C["deep_blue"]} 100%);
  border-radius:12px;padding:22px 28px;color:white;margin-bottom:24px;}}
.exec-header h1{{font-size:21px;font-weight:700;color:white!important;margin:0 0 3px 0;}}
.exec-header .subtitle{{font-size:13px;color:rgba(255,255,255,0.6);margin-bottom:12px;}}
.exec-header .dynamic{{font-size:13px;color:rgba(255,255,255,0.88);
  border-top:1px solid rgba(255,255,255,0.15);padding-top:11px;margin-top:2px;}}
.risk-item{{background:#FEF3F2;border-left:3px solid #C0392B;border-radius:0 6px 6px 0;
  padding:9px 13px;margin-bottom:7px;font-size:12px;color:#1a1a1a;}}
.risk-item.warn{{background:#FFF8F0;border-left-color:#E67E22;}}
.risk-item.info{{background:#F0F7FF;border-left-color:{C["deep_blue"]};}}
.risk-item.ok{{background:#F0FFF8;border-left-color:{C["teal"]};}}
.risk-item.muted{{background:#F9FAFB;border-left-color:{C["gray"]};color:{C["gray"]};}}
.detail-card{{background:{C["white"]};border-radius:10px;padding:16px 18px;
  border:1px solid #E8ECF0;margin-bottom:10px;}}
.detail-label{{font-size:10px;font-weight:700;letter-spacing:0.05em;text-transform:uppercase;
  color:{C["gray"]};margin-bottom:2px;}}
.detail-value{{font-size:13px;font-weight:500;color:{C["navy"]};}}
.detail-value.ph{{color:{C["gray"]};font-style:italic;font-weight:400;}}
.status-badge{{display:inline-block;padding:2px 9px;border-radius:20px;
  font-size:10px;font-weight:700;letter-spacing:0.04em;}}
.empty-state{{text-align:center;padding:40px 24px;color:{C["gray"]};font-size:13px;}}
.edit-banner{{background:#FEF3C7;border:1px solid #F59E0B;border-radius:8px;
  padding:10px 16px;font-size:12px;color:#92400E;margin-bottom:14px;}}
.insight-box{{background:{C["white"]};border-radius:10px;padding:16px 18px;
  border:1px solid #E8ECF0;margin-bottom:12px;}}
.ff-card{{background:{C["white"]};border-radius:10px;padding:13px 16px;
  border:1px solid #E8ECF0;margin-bottom:8px;display:flex;
  align-items:flex-start;gap:14px;}}
.ff-name{{font-size:12px;font-weight:700;color:{C["navy"]};min-width:170px;}}
.ff-why{{font-size:12px;color:#374151;flex:1;line-height:1.5;}}
.ff-badge-yes{{font-size:10px;font-weight:700;background:#D1FAE5;color:#065F46;
  padding:2px 8px;border-radius:10px;white-space:nowrap;}}
.ff-badge-no{{font-size:10px;font-weight:700;background:#F3F4F6;color:{C["gray"]};
  padding:2px 8px;border-radius:10px;white-space:nowrap;}}
.ph-note{{background:#F9FAFB;border:1px dashed #D1D5DB;border-radius:8px;
  padding:10px 14px;font-size:11px;color:{C["gray"]};margin:6px 0 10px 0;}}
[data-testid="stSidebar"]{{background:{C["white"]};border-right:1px solid #E8ECF0;}}
#MainMenu{{visibility:hidden;}}footer{{visibility:hidden;}}header{{visibility:hidden;}}
.stTabs [data-baseweb="tab-list"]{{gap:4px;background:#F0F2F5;border-radius:8px;padding:4px;}}
.stTabs [data-baseweb="tab"]{{border-radius:6px;padding:5px 14px;font-size:12px;font-weight:500;}}
.stTabs [aria-selected="true"]{{background:{C["white"]};color:{C["navy"]};}}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ──────────────────────────────────
def normalize_cols(df):
    df.columns = [c.strip() for c in df.columns]
    return df

def get_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    for c in candidates:
        for col in df.columns:
            if c.lower().replace(" ","").replace("?","") in col.lower().replace(" ","").replace("?",""):
                return col
    return None

def build_download_url(url):
    if "/:x:/p/" in url or "/:x:/s/" in url:
        return url + ("&" if "?" in url else "?") + "download=1"
    if "_layouts/15/Doc.aspx" in url:
        m = re.search(r'sourcedoc=%7B([^%]+)%7D', url, re.IGNORECASE)
        if m:
            base = url.split("/_layouts/")[0]
            return f"{base}/_layouts/15/download.aspx?UniqueId={m.group(1)}"
    if "1drv.ms" in url:
        try:
            r = requests.get(url, allow_redirects=True, timeout=15)
            return r.url + ("&" if "?" in r.url else "?") + "download=1"
        except Exception:
            pass
    return url + ("&" if "?" in url else "?") + "download=1"

def chart_layout(fig, height=300, legend=False):
    fig.update_layout(
        height=height, margin=dict(t=14,b=14,l=8,r=8),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Inter, sans-serif", size=11, color="#374151"),
        showlegend=legend,
        legend=dict(orientation="h",yanchor="bottom",y=1.02,
                    xanchor="right",x=1,font=dict(size=10)) if legend else {},
        xaxis=dict(gridcolor="#F0F2F5",linecolor="#E8ECF0",tickfont=dict(size=10)),
        yaxis=dict(gridcolor="#F0F2F5",linecolor="#E8ECF0",tickfont=dict(size=10)),
    )
    fig.update_traces(marker_line_width=0)
    return fig

def status_badge_html(s):
    cm = {
        "delayed":("#FEE2E2","#C0392B"),"at risk":("#FEF3C7","#D97706"),
        "on track":("#D1FAE5","#065F46"),"active":("#DBEAFE","#1E40AF"),
        "in progress":("#E0F2FE","#0369A1"),"complete":("#D1FAE5","#065F46"),
        "completed":("#D1FAE5","#065F46"),"not started":("#F3F4F6","#374151"),
        "planning":("#EDE9FE","#5B21B6"),
    }
    bg,fg = cm.get(str(s).lower(),("#F3F4F6","#374151"))
    return f"<span class='status-badge' style='background:{bg};color:{fg};'>{s}</span>"

def normalize_cdm(val):
    if pd.isna(val) or str(val).strip()=="": return "Unknown"
    v = str(val).strip().lower()
    if v in ("yes","y","true","1"): return "Yes"
    if v in ("no","n","false","0"): return "No"
    return "Unknown"

def det_jitter(series, scale=0.15):
    def _j(v):
        h = int(hashlib.md5(str(v).encode()).hexdigest(),16)
        return ((h%1000)/1000.0-0.5)*2*scale
    return series.apply(_j)

def kpi_card(col, label, value, sub, color_class="", accent=None, is_ph=False):
    a = accent or C["deep_blue"]
    ph = " ph" if is_ph else ""
    col.markdown(f"""
    <div class="kpi-wrap{ph}">
      <div>
        <div class="kpi-accent-bar" style="background:{a};width:28px;"></div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color_class}">{value}</div>
      </div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def resolve_future(df, candidates):
    for c in candidates:
        f = get_col(df, c)
        if f: return f
    return None

def fval(row, col, fallback="Not yet captured"):
    if col is None: return fallback
    v = row.get(col)
    if v is None: return fallback
    try:
        if pd.isna(v): return fallback
    except Exception:
        pass
    return str(v).strip() or fallback

# ─── DATA LOAD ─────────────────────────────────
@st.cache_data(ttl=60)
def load_data(url):
    try:
        dl = build_download_url(url)
        hdr = {"User-Agent":"Mozilla/5.0"}
        r = requests.get(dl, headers=hdr, timeout=30, allow_redirects=True)
        r.raise_for_status()
        if "html" in r.headers.get("Content-Type","").lower():
            fb = url+("&download=1" if "?" in url else "?download=1")
            r = requests.get(fb, headers=hdr, timeout=30, allow_redirects=True)
            r.raise_for_status()
        content = BytesIO(r.content)
        sheets = {}
        for s in ["Projects","Project_Resources","Dependencies","Project_Value_Map","Value_Category_Dictionary"]:
            try:
                df = pd.read_excel(content, sheet_name=s, engine="openpyxl")
                sheets[s] = normalize_cols(df)
            except Exception:
                sheets[s] = None
        return sheets, None
    except Exception as e:
        return None, str(e)

with st.spinner(""):
    sheets, err = load_data(ONEDRIVE_FILE_URL)

if err:
    st.error(f"**Data load failed:** {err}")
    st.info("Ensure the SharePoint link allows 'Anyone with the link' to view.")
    st.stop()

proj_df  = sheets.get("Projects")
res_df   = sheets.get("Project_Resources")
dep_df   = sheets.get("Dependencies")
val_map_df  = sheets.get("Project_Value_Map")   # normalized many-to-one
val_dict_df = sheets.get("Value_Category_Dictionary")  # reference

for m in [s for s,d in sheets.items() if d is None]:
    st.warning(f"Sheet '{m}' could not be loaded.")

if proj_df is None:
    st.error("Projects sheet is required.")
    st.stop()

# ─── COLUMN MAP ────────────────────────────────
owner_col          = get_col(proj_df,"Owner","owner","PM","Project Owner")
team_col_p         = get_col(proj_df,"Team","team","Department")
status_col         = get_col(proj_df,"Status","status","Project Status")
cycle_col          = get_col(proj_df,"Cycle","cycle","Sprint","Quarter")
priority_col       = get_col(proj_df,"Priority","priority","Priority Type")
effort_col         = get_col(proj_df,"Effort","effort","Effort Score")
impact_col         = get_col(proj_df,"Impact","impact","Impact Score")
proj_id_col        = get_col(proj_df,"Project ID","ProjectID","ID","project_id")
proj_name_col      = get_col(proj_df,"Project","Project Name","project","Name")
delayed_impact_col = get_col(proj_df,"If Delayed Impact","Delayed Impact","delay_impact","Impact If Delayed")
notes_col          = get_col(proj_df,"Notes","notes","Risk Notes","Delay Notes","Comments")
cdm_col_raw        = get_col(proj_df,"Dependent on CDM Project?","Dependent on CDM Project",
                              "CDM Dependency","CDM Project","CDM","DependentonCDMProject")

CDM_COL = "__cdm__"
proj_df[CDM_COL] = proj_df[cdm_col_raw].apply(normalize_cdm) if cdm_col_raw else "Unknown"

# Resolve future fields
fc = {}  # future cols map: display_name -> actual col or None
for fname, cands, _ in FUTURE_FIELDS:
    fc[fname] = resolve_future(proj_df, cands)

team_col_r = None
pid_col_r  = None
if res_df is not None:
    team_col_r = get_col(res_df,"Team","team","Department","Resource Team")
    pid_col_r  = get_col(res_df,"Project ID","ProjectID","ID")

# ─── NEW SHEET COLUMN MAP ─────────────────────────
# Project_Value_Map columns
vm_pid_col   = get_col(val_map_df, "Project ID","ProjectID","ID") if val_map_df is not None else None
vm_cat_col   = get_col(val_map_df, "Value Category","ValueCategory","Category","value_category") if val_map_df is not None else None
vm_grp_col   = get_col(val_map_df, "Value Group","ValueGroup","Group","value_group") if val_map_df is not None else None

# Value_Category_Dictionary columns
vd_cat_col   = get_col(val_dict_df, "Value Category","Category") if val_dict_df is not None else None
vd_grp_col   = get_col(val_dict_df, "Value Group","Group") if val_dict_df is not None else None
vd_desc_col  = get_col(val_dict_df, "Description","Desc","definition") if val_dict_df is not None else None

# Project-level new fields (may or may not exist)
proj_type_col    = get_col(proj_df, "Project Type","ProjectType","Type","project_type")
priority_rank_col= get_col(proj_df, "Priority Rank","PriorityRank","Rank","rank")
biz_val_col      = get_col(proj_df, "Business Value","BusinessValue","business_value")
dar_proj_col     = get_col(proj_df, "Dollars at Risk","DollarsAtRisk","$ at Risk")
func_col         = get_col(proj_df, "Function","function","Functional Area","func")

# Priority Band helper
def assign_priority_band(df, rank_col):
    """Assign Top / Middle / Lower based on Priority Rank if present."""
    if rank_col is None or rank_col not in df.columns:
        return pd.Series(["Unranked"]*len(df), index=df.index)
    ranks = pd.to_numeric(df[rank_col], errors="coerce")
    n = ranks.notna().sum()
    if n == 0:
        return pd.Series(["Unranked"]*len(df), index=df.index)
    top_cut    = ranks.quantile(0.33)
    mid_cut    = ranks.quantile(0.67)
    def _band(r):
        if pd.isna(r): return "Unranked"
        if r <= top_cut: return "Top"
        if r <= mid_cut: return "Middle"
        return "Lower"
    return ranks.apply(_band)

proj_df["__band__"] = assign_priority_band(proj_df, priority_rank_col)

# ─── SESSION STATE ─────────────────────────────
for k,v in [("view","Executive Summary"),("edit_mode",False),("edits",{}),
             ("proj_edits",None),("vm_edits",None),("res_edits",None),("dep_edits",None),
             ("active_tab","Project Configuration")]:
    if k not in st.session_state:
        st.session_state[k] = v

# Initialise editable copies in session state once
if st.session_state["proj_edits"] is None:
    st.session_state["proj_edits"] = proj_df.copy()
if val_map_df is not None and st.session_state["vm_edits"] is None:
    st.session_state["vm_edits"] = val_map_df.copy()
if res_df is not None and st.session_state["res_edits"] is None:
    st.session_state["res_edits"] = res_df.copy()
if dep_df is not None and st.session_state["dep_edits"] is None:
    st.session_state["dep_edits"] = dep_df.copy()

# ─── SIDEBAR ───────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style='margin-bottom:18px;'>
      <div style='font-size:14px;font-weight:700;color:{C["navy"]};'>RevOps Dashboard</div>
      <div style='font-size:10px;color:{C["gray"]};margin-top:2px;letter-spacing:0.05em;'>FILTER CONTROLS</div>
    </div>""", unsafe_allow_html=True)

    base = proj_df.copy()
    sel_owners = []

    if owner_col:
        owners = sorted(base[owner_col].dropna().unique().tolist())
        default_owners = ["RevOps"] if "RevOps" in owners else owners
        sel_owners = st.multiselect("Owner", owners, default=default_owners)
        if sel_owners:
            base = base[base[owner_col].isin(sel_owners)]

    if team_col_p:
        sel_teams_sb = st.multiselect("Team", sorted(proj_df[team_col_p].dropna().unique().tolist()), default=[])
        if sel_teams_sb:
            base = base[base[team_col_p].isin(sel_teams_sb)]

    if status_col:
        sel_status_sb = st.multiselect("Status", sorted(proj_df[status_col].dropna().unique().tolist()), default=[])
        if sel_status_sb:
            base = base[base[status_col].isin(sel_status_sb)]

    if cycle_col:
        sel_cycles_sb = st.multiselect("Cycle", sorted(proj_df[cycle_col].dropna().unique().tolist()), default=[])
        if sel_cycles_sb:
            base = base[base[cycle_col].isin(sel_cycles_sb)]

    if priority_col:
        sel_pris_sb = st.multiselect("Priority Type", sorted(proj_df[priority_col].dropna().unique().tolist()), default=[])
        if sel_pris_sb:
            base = base[base[priority_col].isin(sel_pris_sb)]

    sel_cdm_sb = st.multiselect("CDM Dependency", ["Yes","No","Unknown"], default=[])
    if sel_cdm_sb:
        base = base[base[CDM_COL].isin(sel_cdm_sb)]

    st.markdown("<div style='font-size:10px;font-weight:700;color:#9FA1A4;letter-spacing:0.07em;"
                "text-transform:uppercase;margin:14px 0 6px 0;'>Future Filters</div>",
                unsafe_allow_html=True)
    for fname in ["Business Value","Dollars at Risk","Hard Deadline","Confidence Level","Dependency Criticality"]:
        actual = fc.get(fname)
        if actual:
            opts = sorted(base[actual].dropna().unique().tolist())
            sel_ff = st.multiselect(fname, opts, default=[], key=f"ff_{fname}")
            if sel_ff:
                base = base[base[actual].isin(sel_ff)]
        else:
            st.caption(f"_{fname}: not yet in data_")

    st.markdown("<hr style='border:none;border-top:1px solid #E8ECF0;margin:14px 0;'>",
                unsafe_allow_html=True)
    st.caption(f"**{len(base)}** of {len(proj_df)} projects shown")
    st.caption("Auto-refreshes every 60 s")

filtered = base.copy()

# ─── SHARED METRICS ────────────────────────────
total = len(filtered)
delayed_mask = pd.Series([False]*total, index=filtered.index)
if status_col:
    delayed_mask = filtered[status_col].str.lower().str.contains("delay", na=False)
delayed_count = int(delayed_mask.sum())

active_count = 0
if status_col:
    active_count = int(filtered[status_col].str.lower().str.contains(
        "active|in progress|in-progress", na=False, regex=True).sum())

teams_count = 0
if res_df is not None and team_col_r and pid_col_r and proj_id_col:
    ap = filtered[proj_id_col].dropna().unique()
    rf = res_df[res_df[pid_col_r].isin(ap)]
    teams_count = rf[team_col_r].nunique()

cdm_yes_count     = int((filtered[CDM_COL]=="Yes").sum())
cdm_unknown_count = int((filtered[CDM_COL]=="Unknown").sum())

avg_impact = None
avg_effort = None
if impact_col:
    v = pd.to_numeric(filtered[impact_col], errors="coerce")
    if v.notna().any(): avg_impact = round(float(v.mean()),1)
if effort_col:
    v = pd.to_numeric(filtered[effort_col], errors="coerce")
    if v.notna().any(): avg_effort = round(float(v.mean()),1)

# Future field aggregates
dollars_at_risk_total = None
hard_deadline_count   = None
high_value_count      = None

dar_col  = fc.get("Dollars at Risk")
hd_col   = fc.get("Hard Deadline")
bv_col   = fc.get("Business Value")
fv_col   = fc.get("Financial Value")
cap_col  = fc.get("Estimated Capacity")
dt_col   = fc.get("Deadline Type")
dc_col   = fc.get("Dependency Criticality")
es_col   = fc.get("Executive Sponsor")
cl_col   = fc.get("Confidence Level")
br_col   = fc.get("Blocker Reason")

if dar_col:
    v = pd.to_numeric(filtered[dar_col], errors="coerce")
    if v.notna().any(): dollars_at_risk_total = v.sum()
if hd_col:
    hard_deadline_count = int(filtered[hd_col].notna().sum())
if bv_col:
    v = pd.to_numeric(filtered[bv_col], errors="coerce")
    if v.notna().any():
        threshold = v.quantile(0.75) if v.notna().sum() > 3 else v.max()
        high_value_count = int((v >= threshold).sum())

# Dynamic summary
if owner_col and sel_owners:
    o_label = f"{', '.join(sel_owners)} " if len(sel_owners)<=2 else "filtered "
else:
    o_label = "RevOps " if (owner_col and "RevOps" in proj_df[owner_col].values) else ""

parts = [f"<strong>{total}</strong> {o_label}projects tracked"]
if teams_count: parts.append(f"across <strong>{teams_count}</strong> teams")
if delayed_count:
    parts.append(f"<strong>{delayed_count}</strong> delayed program{'s' if delayed_count!=1 else ''} requiring attention")
else:
    parts.append("no delays currently flagged")
dynamic_summary = ", ".join(parts[:2]) + (f", {parts[2]}" if len(parts)>2 else "") + "."

# ─── SCATTER HELPER ────────────────────────────
def scatter_effort_impact(df, height=300, show_legend=True):
    if not effort_col or not impact_col:
        st.warning("Effort / Impact columns not found.")
        return
    keep = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                         effort_col,impact_col,CDM_COL] if c]
    sdf = df[keep].copy()
    sdf[effort_col] = pd.to_numeric(sdf[effort_col], errors="coerce")
    sdf[impact_col] = pd.to_numeric(sdf[impact_col], errors="coerce")
    sdf = sdf.dropna(subset=[effort_col,impact_col]).reset_index(drop=True)
    if sdf.empty:
        st.markdown("<div class='empty-state'>No numeric data for scatter plot.</div>", unsafe_allow_html=True)
        return
    sdf["__x__"] = sdf[effort_col] + det_jitter(sdf.apply(lambda r: f"{r.name}_{r[effort_col]}",axis=1))
    sdf["__y__"] = sdf[impact_col] + det_jitter(sdf.apply(lambda r: f"{r.name}_{r[impact_col]}",axis=1))
    custom_cols = [c for c in [proj_id_col,proj_name_col,owner_col,effort_col,impact_col,CDM_COL] if c]
    fig = px.scatter(sdf, x="__x__", y="__y__",
                     color=status_col if status_col else None,
                     color_discrete_map=STATUS_COLORS,
                     custom_data=custom_cols,
                     template="plotly_white", opacity=0.78)
    ht = "<br>".join(
        f"<b>{'CDM' if c==CDM_COL else c}:</b> %{{customdata[{i}]}}"
        for i,c in enumerate(custom_cols)
    ) + "<extra></extra>"
    fig.update_traces(marker=dict(size=11,line=dict(width=1.5,color="white")), hovertemplate=ht)
    fig.update_layout(xaxis_title=effort_col or "Effort", yaxis_title=impact_col or "Impact")
    fig = chart_layout(fig, height=height, legend=show_legend)
    st.plotly_chart(fig, use_container_width=True)


# ─── TOP-LEVEL TABS ────────────────────────────
_tab_config = st.tabs(["📋  Project Configuration", "📊  Executive Dashboard",
                        "🔧  Working Team View"])
_tab_config_obj, _tab_exec_obj, _tab_team_obj = _tab_config

# ─── compatibility: map legacy view variable ───
# Each section is now inside a `with` block below.

# ══════════════════════════════════════════════
#  EXECUTIVE DASHBOARD TAB
# ══════════════════════════════════════════════
with _tab_exec_obj:
    st.markdown(f"""
    <div class="exec-header">
      <h1>RevOps Program Dashboard</h1>
      <div class="subtitle">Resource load, project risk, and dependency visibility</div>
      <div class="dynamic">{dynamic_summary}</div>
    </div>""", unsafe_allow_html=True)

    # ── KPI Row 1 ──────────────────────────────
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    kpi_card(k1,"Total Projects",   total,         "in current filters",    accent=C["navy"])
    kpi_card(k2,"Delayed Projects", delayed_count, "require attention",
             color_class="danger" if delayed_count else "",
             accent="#C0392B" if delayed_count else C["gray"])
    kpi_card(k3,"Active Projects",  active_count,  "in progress",           accent=C["deep_blue"])
    kpi_card(k4,"Teams Involved",   teams_count,   "across resource pool",  accent=C["teal"])
    kpi_card(k5,"CDM Dependent",    cdm_yes_count, "depend on CDM",
             color_class="warn" if cdm_yes_count else "", accent="#D97706")
    kpi_card(k6,"Unknown CDM",      cdm_unknown_count, "CDM status not set",
             color_class="warn" if cdm_unknown_count else "", accent="#D97706")

    # ── KPI Row 2 — future fields ───────────────
    st.markdown("<div style='margin-top:12px;'></div>", unsafe_allow_html=True)
    fk1,fk2,fk3,fk4 = st.columns(4)

    if dollars_at_risk_total is not None:
        disp = f"${dollars_at_risk_total:,.0f}"
        kpi_card(fk1,"Dollars at Risk",disp,"estimated exposure",
                 color_class="danger", accent="#C0392B")
    else:
        kpi_card(fk1,"Dollars at Risk","—","not yet captured",
                 color_class="ph", accent=C["gray"], is_ph=True)

    if high_value_count is not None:
        kpi_card(fk2,"High-Value Projects",high_value_count,"top quartile business value",
                 color_class="success", accent=C["teal"])
    else:
        kpi_card(fk2,"High-Value Projects","—","not yet captured",
                 color_class="ph", accent=C["gray"], is_ph=True)

    if hard_deadline_count is not None:
        kpi_card(fk3,"Hard Deadlines",hard_deadline_count,"projects with fixed deadlines",
                 color_class="warn" if hard_deadline_count else "", accent="#D97706")
    else:
        kpi_card(fk3,"Hard Deadlines","—","not yet captured",
                 color_class="ph", accent=C["gray"], is_ph=True)

    cl_captured = cl_col is not None
    if cl_captured:
        low_conf = int(filtered[cl_col].astype(str).str.lower().str.contains(
            "low|uncertain|tbd", na=False).sum())
        kpi_card(fk4,"Low Confidence",low_conf,"projects flagged uncertain",
                 color_class="warn" if low_conf else "", accent="#D97706")
    else:
        kpi_card(fk4,"Confidence Level","—","not yet captured",
                 color_class="ph", accent=C["gray"], is_ph=True)

    if not dar_col and not hd_col and not bv_col:
        st.markdown("<div class='ph-note'>ℹ️ Business Value, Dollars at Risk, and Hard Deadline fields are not yet "
                    "present in the source data. Add these columns to the Excel file to enable "
                    "financial and deadline risk analysis.</div>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    if total == 0:
        st.markdown("<div class='empty-state'>No projects match the current filters.</div>",
                    unsafe_allow_html=True)
        st.stop()

    # ── Overview charts ────────────────────────
    st.markdown("<div class='section-title'>Executive Overview</div>", unsafe_allow_html=True)
    st.markdown("")

    ch1,ch2 = st.columns(2)
    with ch1:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Top Teams by Project Load</div>",
                    unsafe_allow_html=True)
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            cr = res_df[res_df[pid_col_r].isin(ap)]
            if not cr.empty:
                tc = (cr.groupby(team_col_r)[pid_col_r].nunique().reset_index()
                        .rename(columns={team_col_r:"Team",pid_col_r:"Projects"})
                        .sort_values("Projects",ascending=True).tail(10))
                fig = px.bar(tc,x="Projects",y="Team",orientation="h",color="Projects",
                             color_continuous_scale=[[0,C["light_blue"]],[1,C["deep_blue"]]],
                             template="plotly_white")
                fig = chart_layout(fig,height=270)
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig,use_container_width=True)
            else:
                st.markdown("<div class='empty-state'>No resource data.</div>",unsafe_allow_html=True)
        else:
            st.warning("Resource data unavailable.")

    with ch2:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Portfolio Status</div>",
                    unsafe_allow_html=True)
        if status_col:
            sc = filtered[status_col].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            fig2 = px.bar(sc.sort_values("Count",ascending=False),
                          x="Status",y="Count",color="Status",
                          color_discrete_map=STATUS_COLORS,template="plotly_white")
            fig2 = chart_layout(fig2,height=270)
            fig2.update_layout(showlegend=False)
            st.plotly_chart(fig2,use_container_width=True)
        else:
            st.warning("Status column not found.")

    vis_col, risk_col = st.columns([3,2])
    with vis_col:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Impact vs. Effort</div>",
                    unsafe_allow_html=True)
        scatter_effort_impact(filtered, height=290, show_legend=True)

    with risk_col:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Risk Summary</div>",
                    unsafe_allow_html=True)
        shown = 0
        if delayed_count:
            st.markdown(f"<div class='risk-item'>⚠️ <strong>{delayed_count}</strong> "
                        f"delayed project{'s' if delayed_count!=1 else ''}</div>",
                        unsafe_allow_html=True); shown+=1
        if status_col and impact_col:
            hid = filtered[delayed_mask & (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hid.empty:
                st.markdown(f"<div class='risk-item'>🔴 <strong>{len(hid)}</strong> "
                            f"delayed + high-impact</div>",unsafe_allow_html=True); shown+=1
        if effort_col and impact_col:
            hi_both = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi_both.empty:
                st.markdown(f"<div class='risk-item warn'>🔶 <strong>{len(hi_both)}</strong> "
                            f"high-effort + high-impact</div>",unsafe_allow_html=True); shown+=1
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            tdf = res_df[res_df[pid_col_r].isin(ap)]
            if not tdf.empty:
                grp = tdf.groupby(team_col_r)[pid_col_r].nunique()
                tt=grp.idxmax(); tv=int(grp.max())
                st.markdown(f"<div class='risk-item warn'>📌 <strong>{tt}</strong> — "
                            f"highest load ({tv} projects)</div>",unsafe_allow_html=True); shown+=1
        if cdm_yes_count:
            st.markdown(f"<div class='risk-item info'>🔗 <strong>{cdm_yes_count}</strong> "
                        f"CDM-dependent</div>",unsafe_allow_html=True); shown+=1
        if cdm_unknown_count:
            st.markdown(f"<div class='risk-item info'>❓ <strong>{cdm_unknown_count}</strong> "
                        f"unknown CDM dependency</div>",unsafe_allow_html=True); shown+=1
        if dollars_at_risk_total is not None:
            st.markdown(f"<div class='risk-item warn'>💰 <strong>${dollars_at_risk_total:,.0f}</strong> "
                        f"estimated at risk</div>",unsafe_allow_html=True); shown+=1
        if hard_deadline_count:
            st.markdown(f"<div class='risk-item warn'>📅 <strong>{hard_deadline_count}</strong> "
                        f"hard deadline{'s' if hard_deadline_count!=1 else ''}</div>",
                        unsafe_allow_html=True); shown+=1
        if shown==0:
            st.markdown("<div class='risk-item ok'>✅ No critical risks identified.</div>",
                        unsafe_allow_html=True)

        # Narrative
        n_parts = []
        if delayed_count: n_parts.append(f"{delayed_count} delayed")
        if cdm_yes_count: n_parts.append(f"{cdm_yes_count} CDM-dependent")
        if dollars_at_risk_total: n_parts.append(f"${dollars_at_risk_total:,.0f} estimated at risk")
        if n_parts:
            note = " (based on current heuristics only)" if not dar_col else ""
            st.markdown(f"<div style='margin-top:10px;font-size:11px;color:#374151;line-height:1.5;'>"
                        f"Key risks: {'; '.join(n_parts)}.{note}</div>",unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Project Spotlight ──────────────────────
    st.markdown("<div class='section-title'>Project Spotlight</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:12px;color:#6B7280;margin-bottom:10px;'>"
                "Ranked by delay status, impact, and effort. Heuristic score — enrich with Business Value "
                "and Dollars at Risk for stronger prioritisation.</div>", unsafe_allow_html=True)
    spot = filtered.copy()
    spot["__score__"] = 0
    if status_col:
        spot["__score__"] += spot[status_col].str.lower().str.contains("delay",na=False).astype(int)*10
    if impact_col:
        spot["__score__"] += pd.to_numeric(spot[impact_col],errors="coerce").fillna(0)
    if effort_col:
        spot["__score__"] += pd.to_numeric(spot[effort_col],errors="coerce").fillna(0)*0.5
    if dar_col:
        spot["__score__"] += pd.to_numeric(spot[dar_col],errors="coerce").fillna(0)/1000
    spot = spot.sort_values("__score__",ascending=False).head(15)
    sp_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,cycle_col,
                            impact_col,effort_col,CDM_COL,delayed_impact_col,
                            dar_col,hd_col,bv_col] if c]
    disp = spot[sp_cols].rename(columns={CDM_COL:"CDM Dependency"}).reset_index(drop=True)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:36px 0 14px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      RevOps Program Dashboard · Executive View · Refreshes every 60 s · Source: SharePoint
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  WORKING TEAM TAB
# ══════════════════════════════════════════════
with _tab_team_obj:
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{C['green']} 0%,{C['teal']} 100%);
      border-radius:12px;padding:18px 24px;color:white;margin-bottom:20px;">
      <div style="font-size:18px;font-weight:700;">RevOps Working Team View</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.65);margin-top:3px;">{dynamic_summary}</div>
    </div>""", unsafe_allow_html=True)

    em_col,_ = st.columns([2,6])
    with em_col:
        edit_mode = st.toggle("✏️ Edit Mode", value=st.session_state["edit_mode"])
        st.session_state["edit_mode"] = edit_mode
    if edit_mode:
        st.markdown("<div class='edit-banner'>⚠️ <strong>Edit Mode ON.</strong> "
                    "Edits stored in session only — not saved to OneDrive.</div>",
                    unsafe_allow_html=True)

    # ── KPI Row ────────────────────────────────
    k1,k2,k3,k4,k5,k6 = st.columns(6)
    kpi_card(k1,"Total Projects",   total,             "in current filters",   accent=C["navy"])
    kpi_card(k2,"Delayed Projects", delayed_count,     "require attention",
             color_class="danger" if delayed_count else "",
             accent="#C0392B" if delayed_count else C["gray"])
    kpi_card(k3,"Active Projects",  active_count,      "in progress",          accent=C["deep_blue"])
    kpi_card(k4,"Teams Involved",   teams_count,       "across resource pool", accent=C["teal"])
    kpi_card(k5,"CDM Dependent",    cdm_yes_count,     "depend on CDM",
             color_class="warn" if cdm_yes_count else "", accent="#D97706")
    kpi_card(k6,"Unknown CDM",      cdm_unknown_count, "CDM status not set",
             color_class="warn" if cdm_unknown_count else "", accent="#D97706")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    if total == 0:
        st.markdown("<div class='empty-state'>No projects match the current filters.</div>",
                    unsafe_allow_html=True)
        st.stop()

    # ── Portfolio Analysis ──────────────────────
    st.markdown("<div class='section-title'>Portfolio Analysis</div>", unsafe_allow_html=True)
    st.markdown("")

    pa1,pa2 = st.columns(2)
    with pa1:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Projects by Team</div>",
                    unsafe_allow_html=True)
        team_drill_opts = ["All Teams"]
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            cr = res_df[res_df[pid_col_r].isin(ap)]
            if not cr.empty:
                tc = (cr.groupby(team_col_r)[pid_col_r].nunique().reset_index()
                        .rename(columns={team_col_r:"Team",pid_col_r:"Projects"})
                        .sort_values("Projects",ascending=True))
                team_drill_opts = ["All Teams"]+tc["Team"].tolist()
                fig = px.bar(tc,x="Projects",y="Team",orientation="h",color="Projects",
                             color_continuous_scale=[[0,C["light_blue"]],[1,C["deep_blue"]]],
                             template="plotly_white")
                fig = chart_layout(fig,height=300)
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig,use_container_width=True)
        else:
            st.warning("Resource data unavailable.")
        drill_team = st.selectbox("🔍 Drill into team", team_drill_opts, key="drill_team_sel")

    with pa2:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Projects by Status</div>",
                    unsafe_allow_html=True)
        status_drill_opts = ["All Statuses"]
        if status_col:
            sc = filtered[status_col].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            sc = sc.sort_values("Count",ascending=False)
            status_drill_opts = ["All Statuses"]+sc["Status"].tolist()
            fig2 = px.bar(sc,x="Status",y="Count",color="Status",
                          color_discrete_map=STATUS_COLORS,template="plotly_white")
            fig2 = chart_layout(fig2,height=300)
            fig2.update_layout(showlegend=False)
            st.plotly_chart(fig2,use_container_width=True)
        else:
            st.warning("Status column not found.")
        drill_status = st.selectbox("🔍 Drill into status", status_drill_opts, key="drill_status_sel")

    pa3,pa4 = st.columns(2)
    with pa3:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;margin-top:12px;'>Impact vs. Effort</div>",
                    unsafe_allow_html=True)
        scatter_effort_impact(filtered, height=290, show_legend=True)

    with pa4:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;margin-top:12px;'>Priority Distribution</div>",
                    unsafe_allow_html=True)
        if priority_col:
            pc = filtered[priority_col].value_counts().reset_index()
            pc.columns = ["Priority","Count"]
            fig4 = px.pie(pc,names="Priority",values="Count",
                          color_discrete_sequence=PALETTE,hole=0.50,template="plotly_white")
            fig4.update_traces(textposition="outside",textfont_size=10,
                               marker=dict(line=dict(color="white",width=2)))
            fig4.update_layout(height=260,margin=dict(t=8,b=8,l=8,r=8),showlegend=True,
                                legend=dict(font=dict(size=10)),paper_bgcolor="white",
                                font=dict(family="Inter, sans-serif"))
            st.plotly_chart(fig4,use_container_width=True)
        else:
            st.info("Priority column not found.")

    # Delayed table
    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                f"text-transform:uppercase;margin-bottom:8px;margin-top:4px;'>Delayed Projects</div>",
                unsafe_allow_html=True)
    if status_col:
        del_df = filtered[delayed_mask].copy()
        dcols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,cycle_col,
                              impact_col,effort_col,delayed_impact_col,CDM_COL,
                              dar_col,hd_col] if c]
        if not del_df.empty and dcols:
            st.dataframe(del_df[dcols].rename(columns={CDM_COL:"CDM Dependency"})
                         .reset_index(drop=True), use_container_width=True, hide_index=True)
        else:
            st.markdown(f"<div style='color:{C['teal']};font-size:12px;padding:6px 0;'>"
                        f"✅ No delayed projects in current view.</div>", unsafe_allow_html=True)

    # CDM breakdown
    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                f"text-transform:uppercase;margin-bottom:8px;margin-top:18px;'>CDM Dependency Breakdown</div>",
                unsafe_allow_html=True)
    cdm_counts = filtered[CDM_COL].value_counts().reset_index()
    cdm_counts.columns = ["CDM Status","Count"]
    fig_cdm = px.bar(cdm_counts,x="CDM Status",y="Count",color="CDM Status",
                     color_discrete_map={"Yes":"#D97706","No":C["teal"],"Unknown":C["gray"]},
                     template="plotly_white")
    fig_cdm = chart_layout(fig_cdm,height=200)
    fig_cdm.update_layout(showlegend=False)
    st.plotly_chart(fig_cdm,use_container_width=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Interactive Insights ────────────────────
    st.markdown("<div class='section-title'>Interactive Insights</div>", unsafe_allow_html=True)
    st.markdown("")
    ins1,ins2,ins3,ins4 = st.columns(4)
    with ins1:
        ins_teams = ["All Teams"]
        if team_col_p: ins_teams += sorted(filtered[team_col_p].dropna().unique().tolist())
        elif res_df is not None and team_col_r: ins_teams += sorted(res_df[team_col_r].dropna().unique().tolist())
        ins_team = st.selectbox("By Team", ins_teams)
    with ins2:
        ins_statuses = ["All Statuses"]
        if status_col: ins_statuses += sorted(filtered[status_col].dropna().unique().tolist())
        ins_status = st.selectbox("By Status", ins_statuses)
    with ins3:
        ins_cdm = st.selectbox("By CDM", ["All","Yes","No","Unknown"])
    with ins4:
        proj_opts = ["All Projects"]
        if proj_name_col: proj_opts += sorted(filtered[proj_name_col].dropna().unique().tolist())
        ins_proj = st.selectbox("By Project", proj_opts)

    ins_df = filtered.copy()
    if ins_team != "All Teams":
        if team_col_p and team_col_p in ins_df.columns:
            ins_df = ins_df[ins_df[team_col_p]==ins_team]
        elif res_df is not None and team_col_r and pid_col_r and proj_id_col:
            pids = res_df[res_df[team_col_r]==ins_team][pid_col_r].unique()
            ins_df = ins_df[ins_df[proj_id_col].isin(pids)]
    if ins_status != "All Statuses" and status_col:
        ins_df = ins_df[ins_df[status_col]==ins_status]
    if ins_cdm != "All":
        ins_df = ins_df[ins_df[CDM_COL]==ins_cdm]
    if ins_proj != "All Projects" and proj_name_col:
        ins_df = ins_df[ins_df[proj_name_col]==ins_proj]

    st.markdown(f"<div style='font-size:12px;color:#6B7280;margin-bottom:10px;'>"
                f"<strong>{len(ins_df)}</strong> projects match selected filters.</div>",
                unsafe_allow_html=True)

    if not ins_df.empty:
        il,ir = st.columns([3,2])
        with il:
            ins_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                                     cycle_col,impact_col,effort_col,CDM_COL,
                                     dar_col,cl_col,dc_col] if c]
            st.dataframe(ins_df[ins_cols].rename(columns={CDM_COL:"CDM Dependency"})
                         .reset_index(drop=True), use_container_width=True, hide_index=True, height=280)
        with ir:
            if res_df is not None and team_col_r and pid_col_r and proj_id_col:
                ap_i = ins_df[proj_id_col].dropna().unique()
                rf_i = res_df[res_df[pid_col_r].isin(ap_i)]
                if not rf_i.empty:
                    imp_t = rf_i[team_col_r].value_counts().head(8).reset_index()
                    imp_t.columns = ["Team","Projects"]
                    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                f"letter-spacing:0.06em;text-transform:uppercase;margin-bottom:6px;'>"
                                f"Impacted Teams</div>", unsafe_allow_html=True)
                    fig_it = px.bar(imp_t.sort_values("Projects",ascending=True),
                                    x="Projects",y="Team",orientation="h",color="Projects",
                                    color_continuous_scale=[[0,C["soft_green"]],[1,C["green"]]],
                                    template="plotly_white")
                    fig_it = chart_layout(fig_it,height=220)
                    fig_it.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(fig_it,use_container_width=True)
            if dep_df is not None and proj_id_col:
                dep_pid = get_col(dep_df,"Project ID","ProjectID","ID","Dependent Project ID")
                if dep_pid:
                    dm = dep_df[dep_df[dep_pid].isin(ins_df[proj_id_col].dropna().unique())]
                    if not dm.empty:
                        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                    f"letter-spacing:0.06em;text-transform:uppercase;"
                                    f"margin-bottom:6px;margin-top:10px;'>Related Dependencies</div>",
                                    unsafe_allow_html=True)
                        st.dataframe(dm.head(10).reset_index(drop=True),
                                     use_container_width=True, hide_index=True, height=140)
    else:
        st.markdown("<div class='empty-state'>No projects match this combination.</div>",
                    unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Risk Insights ───────────────────────────
    st.markdown("<div class='section-title'>Risk Insights</div>", unsafe_allow_html=True)
    st.markdown("")

    ri1,ri2 = st.columns([2,3])
    with ri1:
        r_shown=0
        if delayed_count:
            st.markdown(f"<div class='risk-item'>⚠️ <strong>{delayed_count}</strong> "
                        f"delayed project{'s' if delayed_count!=1 else ''}</div>",
                        unsafe_allow_html=True); r_shown+=1
        if status_col and impact_col:
            hid = filtered[delayed_mask & (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hid.empty:
                st.markdown(f"<div class='risk-item'>🔴 <strong>{len(hid)}</strong> "
                            f"delayed + high-impact</div>",unsafe_allow_html=True); r_shown+=1
        if effort_col and impact_col:
            hi_both = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi_both.empty:
                st.markdown(f"<div class='risk-item warn'>🔶 <strong>{len(hi_both)}</strong> "
                            f"high-effort + high-impact</div>",unsafe_allow_html=True); r_shown+=1
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            tdf = res_df[res_df[pid_col_r].isin(ap)]
            if not tdf.empty:
                grp=tdf.groupby(team_col_r)[pid_col_r].nunique()
                tt=grp.idxmax(); tv=int(grp.max())
                st.markdown(f"<div class='risk-item warn'>📌 <strong>{tt}</strong> — "
                            f"highest load ({tv} projects)</div>",unsafe_allow_html=True); r_shown+=1
        if cdm_yes_count:
            st.markdown(f"<div class='risk-item info'>🔗 <strong>{cdm_yes_count}</strong> "
                        f"CDM-dependent</div>",unsafe_allow_html=True); r_shown+=1
        if cdm_unknown_count:
            st.markdown(f"<div class='risk-item info'>❓ <strong>{cdm_unknown_count}</strong> "
                        f"unknown CDM dependency</div>",unsafe_allow_html=True); r_shown+=1
        if priority_col:
            high_pri = filtered[filtered[priority_col].astype(str).str.lower()
                                .str.contains("high|critical|p1",na=False)]
            if not high_pri.empty:
                st.markdown(f"<div class='risk-item info'>🔵 <strong>{len(high_pri)}</strong> "
                            f"high-priority project{'s' if len(high_pri)!=1 else ''}</div>",
                            unsafe_allow_html=True); r_shown+=1
        if dollars_at_risk_total is not None:
            st.markdown(f"<div class='risk-item warn'>💰 <strong>${dollars_at_risk_total:,.0f}</strong> "
                        f"estimated at risk</div>",unsafe_allow_html=True); r_shown+=1
        else:
            st.markdown("<div class='risk-item muted'>💰 Dollars at Risk: not yet captured</div>",
                        unsafe_allow_html=True)
        if hard_deadline_count:
            st.markdown(f"<div class='risk-item warn'>📅 <strong>{hard_deadline_count}</strong> "
                        f"hard deadline{'s' if hard_deadline_count!=1 else ''}</div>",
                        unsafe_allow_html=True); r_shown+=1
        else:
            st.markdown("<div class='risk-item muted'>📅 Hard Deadlines: not yet captured</div>",
                        unsafe_allow_html=True)
        if r_shown==0:
            st.markdown("<div class='risk-item ok'>✅ No critical risks identified.</div>",
                        unsafe_allow_html=True)

    with ri2:
        n_parts=[]
        if delayed_count: n_parts.append(f"{delayed_count} delayed project{'s' if delayed_count!=1 else ''}")
        if status_col and impact_col:
            hid2 = filtered[delayed_mask & (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hid2.empty: n_parts.append(f"{len(hid2)} with high impact score")
        if effort_col and impact_col:
            hi2 = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi2.empty: n_parts.append(f"{len(hi2)} requiring high effort and high impact")
        if cdm_yes_count: n_parts.append(f"{cdm_yes_count} dependent on CDM delivery")
        if cdm_unknown_count: n_parts.append(f"{cdm_unknown_count} with unresolved CDM status")
        if dollars_at_risk_total: n_parts.append(f"${dollars_at_risk_total:,.0f} estimated at risk")
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap3 = filtered[proj_id_col].dropna().unique()
            tdf3 = res_df[res_df[pid_col_r].isin(ap3)]
            if not tdf3.empty:
                grp3=tdf3.groupby(team_col_r)[pid_col_r].nunique()
                n_parts.append(f"{grp3.idxmax()} carrying the highest load at {int(grp3.max())} projects")

        heuristic_note = (" Recommendations based on current heuristics only — add Business Value, "
                          "Dollars at Risk, and Hard Deadline fields for stronger prioritisation."
                          if not dar_col and not bv_col and not hd_col else
                          " Some prioritisation fields are present; remaining recommendations "
                          "supplemented by heuristics where data is absent.")
        if n_parts:
            narrative = ("The current portfolio shows: " + "; ".join(n_parts) + ". " +
                         "Review prioritisation and resource allocation to reduce delivery risk." +
                         heuristic_note)
        else:
            narrative = "Portfolio appears on track with no major risks under current filters." + heuristic_note

        st.markdown(f"""
        <div class="insight-box">
          <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
            text-transform:uppercase;margin-bottom:8px;">Dynamic Risk Narrative</div>
          <div style="font-size:13px;color:#374151;line-height:1.6;">{narrative}</div>
        </div>""", unsafe_allow_html=True)

        if effort_col and impact_col:
            hi_risk = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)].copy()
            if not hi_risk.empty:
                st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                            f"letter-spacing:0.06em;text-transform:uppercase;"
                            f"margin-bottom:6px;margin-top:10px;'>High Effort + High Impact</div>",
                            unsafe_allow_html=True)
                hr_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                                        effort_col,impact_col,CDM_COL,dar_col] if c]
                st.dataframe(hi_risk[hr_cols].rename(columns={CDM_COL:"CDM Dependency"})
                             .reset_index(drop=True),
                             use_container_width=True, hide_index=True, height=160)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Decision Inputs Panel ───────────────────
    st.markdown("<div class='section-title'>Decision Inputs to Add</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:12px;color:#6B7280;margin-bottom:14px;'>"
                "These fields would significantly strengthen portfolio prioritisation. "
                "Add them to the source Excel file to unlock richer analysis.</div>",
                unsafe_allow_html=True)
    for fname, cands, why in FUTURE_FIELDS:
        actual = fc.get(fname)
        badge = (f"<span class='ff-badge-yes'>✓ Available</span>" if actual
                 else f"<span class='ff-badge-no'>Not yet captured</span>")
        st.markdown(f"""
        <div class="ff-card">
          <div class="ff-name">{fname}</div>
          <div class="ff-why">{why}</div>
          <div>{badge}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Project Explorer ───────────────────────
    st.markdown("<div class='section-title'>Project Explorer</div>", unsafe_allow_html=True)
    st.markdown("")

    if proj_name_col:
        proj_opts_ex = sorted(filtered[proj_name_col].dropna().unique().tolist())
    elif proj_id_col:
        proj_opts_ex = sorted(filtered[proj_id_col].dropna().astype(str).unique().tolist())
    else:
        proj_opts_ex = []

    if not proj_opts_ex:
        st.markdown("<div class='empty-state'>No projects available.</div>",unsafe_allow_html=True)
    else:
        sel_proj = st.selectbox("Select a project to explore", proj_opts_ex)
        prow_df = (filtered[filtered[proj_name_col]==sel_proj] if proj_name_col
                   else filtered[filtered[proj_id_col].astype(str)==sel_proj])

        if not prow_df.empty:
            row = prow_df.iloc[0]
            proj_status = row.get(status_col,"") if status_col else ""
            is_delayed  = "delay" in str(proj_status).lower()
            badge_h     = status_badge_html(proj_status) if proj_status else ""
            pid_disp    = (f"<span style='color:{C['gray']};font-size:12px;margin-left:8px;'>"
                           f"{row.get(proj_id_col,'')}</span>" if proj_id_col else "")
            cdm_val     = row.get(CDM_COL,"Unknown")
            cdm_color   = {"Yes":"#D97706","No":C["teal"],"Unknown":C["gray"]}.get(cdm_val,C["gray"])

            st.markdown(f"""
            <div class="detail-card" style="border-left:4px solid {'#C0392B' if is_delayed else C['deep_blue']};">
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <span style="font-size:16px;font-weight:700;color:{C['navy']};">{sel_proj}</span>
                {pid_disp} {badge_h}
                <span class='status-badge' style='background:#FEF3C7;color:{cdm_color};'>CDM: {cdm_val}</span>
              </div>
              <div style="margin-top:5px;font-size:11px;color:#C0392B;">
                {'⚠️ Flagged as delayed.' if is_delayed else ''}
              </div>
            </div>""", unsafe_allow_html=True)

            tab1,tab2,tab3,tab4,tab5 = st.tabs(
                ["Overview","Resources","Dependencies","Risk & Impact","Business & Decision"])

            # ── Tab 1: Overview ────────────────
            with tab1:
                meta = [c for c in [proj_id_col,owner_col,team_col_p,status_col,
                                     cycle_col,priority_col,effort_col,impact_col,CDM_COL] if c]
                if meta:
                    pairs = [(c,row.get(c,"—")) for c in meta]
                    half  = len(pairs)//2 + len(pairs)%2
                    m1,m2 = st.columns(2)
                    for cn,cv in pairs[:half]:
                        lbl = "CDM Dependency" if cn==CDM_COL else cn
                        m1.markdown(f"""
                        <div style="margin-bottom:12px;">
                          <div class="detail-label">{lbl}</div>
                          <div class="detail-value">{cv if cv==cv and str(cv).strip()!='' else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    for cn,cv in pairs[half:]:
                        lbl = "CDM Dependency" if cn==CDM_COL else cn
                        m2.markdown(f"""
                        <div style="margin-bottom:12px;">
                          <div class="detail-label">{lbl}</div>
                          <div class="detail-value">{cv if cv==cv and str(cv).strip()!='' else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                if notes_col and row.get(notes_col):
                    st.markdown(f"""
                    <div class="detail-card" style="background:#FFFBF0;border-left:3px solid #D97706;">
                      <div class="detail-label">Notes</div>
                      <div style="font-size:12px;color:#374151;margin-top:3px;">{row.get(notes_col)}</div>
                    </div>""", unsafe_allow_html=True)

                if edit_mode:
                    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                f"letter-spacing:0.06em;text-transform:uppercase;margin:14px 0 8px 0;'>"
                                f"Edit Core Fields</div>", unsafe_allow_html=True)
                    proj_key = str(row.get(proj_id_col, sel_proj))
                    edits = st.session_state["edits"].setdefault(proj_key, {})
                    e1,e2 = st.columns(2)
                    if status_col:
                        all_s = sorted(proj_df[status_col].dropna().unique().tolist())
                        cur_s = row.get(status_col)
                        edits["status"] = e1.selectbox("Status",all_s,
                            index=all_s.index(cur_s) if cur_s in all_s else 0,
                            key=f"es_{proj_key}")
                    if owner_col:
                        edits["owner"] = e2.text_input("Owner",
                            value=str(row.get(owner_col,"")), key=f"eo_{proj_key}")
                    if impact_col:
                        edits["impact"] = e1.text_input("Impact",
                            value=str(row.get(impact_col,"")), key=f"ei_{proj_key}")
                    if effort_col:
                        edits["effort"] = e2.text_input("Effort",
                            value=str(row.get(effort_col,"")), key=f"ef_{proj_key}")
                    if notes_col:
                        edits["notes"] = st.text_area("Notes",
                            value=str(row.get(notes_col,"")), key=f"en_{proj_key}")
                    st.session_state["edits"][proj_key] = edits
                    st.markdown("<div class='edit-banner' style='margin-top:8px;'>"
                                "Session edits stored. Writeback to OneDrive not yet implemented.</div>",
                                unsafe_allow_html=True)

            # ── Tab 2: Resources ───────────────
            with tab2:
                if res_df is not None and proj_id_col and pid_col_r:
                    pid_v = row.get(proj_id_col)
                    pr = res_df[res_df[pid_col_r]==pid_v]
                    if not pr.empty:
                        st.dataframe(pr.reset_index(drop=True),use_container_width=True,hide_index=True)
                        if team_col_r:
                            tl = pr[team_col_r].dropna().unique().tolist()
                            st.markdown(f"<div style='font-size:11px;color:{C['gray']};margin-top:6px;'>"
                                        f"Teams: {', '.join(str(t) for t in tl)}</div>",
                                        unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='empty-state'>No resources linked.</div>",
                                    unsafe_allow_html=True)
                else:
                    st.warning("Resource data or Project ID unavailable.")

            # ── Tab 3: Dependencies ────────────
            with tab3:
                if dep_df is not None:
                    dep_pid_col2 = get_col(dep_df,"Project ID","ProjectID","ID","Dependent Project ID")
                    if proj_id_col and dep_pid_col2:
                        pd_dep = dep_df[dep_df[dep_pid_col2]==row.get(proj_id_col)]
                        if not pd_dep.empty:
                            st.dataframe(pd_dep.reset_index(drop=True),
                                         use_container_width=True,hide_index=True)
                        else:
                            st.markdown("<div class='empty-state'>No dependencies recorded.</div>",
                                        unsafe_allow_html=True)
                    else:
                        st.warning("Project ID column missing in Dependencies.")
                else:
                    st.warning("Dependencies sheet not available.")

            # ── Tab 4: Risk & Impact ───────────
            with tab4:
                r1,r2 = st.columns(2)
                with r1:
                    if impact_col:
                        iv = pd.to_numeric(row.get(impact_col),errors="coerce")
                        ic = C["teal"] if pd.notna(iv) and iv>=3 else C["gray"]
                        st.markdown(f"""
                        <div class="detail-card">
                          <div class="detail-label">Impact Score</div>
                          <div class="kpi-value" style="font-size:28px;color:{ic};">
                            {iv if pd.notna(iv) else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    if effort_col:
                        ev = pd.to_numeric(row.get(effort_col),errors="coerce")
                        st.markdown(f"""
                        <div class="detail-card">
                          <div class="detail-label">Effort Score</div>
                          <div class="kpi-value" style="font-size:28px;color:{C['deep_blue']};">
                            {ev if pd.notna(ev) else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    st.markdown(f"""
                    <div class="detail-card">
                      <div class="detail-label">CDM Dependency</div>
                      <div style="font-size:16px;font-weight:600;color:{cdm_color};margin-top:2px;">
                        {cdm_val}</div>
                    </div>""", unsafe_allow_html=True)
                with r2:
                    if delayed_impact_col:
                        div_v = row.get(delayed_impact_col,"—")
                        st.markdown(f"""
                        <div class="detail-card" style="border-left:3px solid #C0392B;">
                          <div class="detail-label">If Delayed Impact</div>
                          <div style="font-size:13px;font-weight:500;color:#C0392B;margin-top:3px;">
                            {div_v if div_v==div_v else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    flag_bg  = "#FEF3F2" if is_delayed else "#F0FFF8"
                    flag_brd = "#C0392B" if is_delayed else C["teal"]
                    flag_clr = "#C0392B" if is_delayed else C["teal"]
                    flag_txt = "⚠️ Delay Flag Active" if is_delayed else "✅ No Delay Flag"
                    flag_sub = "Review ownership and blockers." if is_delayed else "Not currently delayed."
                    st.markdown(f"""
                    <div class="detail-card" style="background:{flag_bg};border-left:3px solid {flag_brd};">
                      <div style="font-size:12px;color:{flag_clr};font-weight:700;">{flag_txt}</div>
                      <div style="font-size:11px;color:#374151;margin-top:3px;">{flag_sub}</div>
                    </div>""", unsafe_allow_html=True)
                    if notes_col and row.get(notes_col):
                        st.markdown(f"""
                        <div class="detail-card" style="background:#FFFBF0;border-left:3px solid #D97706;">
                          <div class="detail-label">Risk Notes</div>
                          <div style="font-size:11px;color:#374151;margin-top:3px;">{row.get(notes_col)}</div>
                        </div>""", unsafe_allow_html=True)

            # ── Tab 5: Business & Decision ──────
            with tab5:
                st.markdown(f"<div style='font-size:11px;color:{C['gray']};margin-bottom:14px;line-height:1.5;'>"
                            f"These fields support stronger portfolio prioritisation. "
                            f"Fields marked <em>Not yet captured</em> are not present in the current "
                            f"source data — add them to the Excel file to enable richer analysis.</div>",
                            unsafe_allow_html=True)
                bd1,bd2 = st.columns(2)
                left_ff  = ["Business Value","Financial Value","Dollars at Risk",
                             "Estimated Capacity","Hard Deadline"]
                right_ff = ["Deadline Type","Dependency Criticality","Executive Sponsor",
                             "Confidence Level","Blocker Reason"]

                for fname in left_ff:
                    actual = fc.get(fname)
                    v = fval(row, actual)
                    ph_class = " ph" if v=="Not yet captured" else ""
                    bd1.markdown(f"""
                    <div style="margin-bottom:14px;">
                      <div class="detail-label">{fname}</div>
                      <div class="detail-value{ph_class}">{v}</div>
                    </div>""", unsafe_allow_html=True)

                for fname in right_ff:
                    actual = fc.get(fname)
                    v = fval(row, actual)
                    ph_class = " ph" if v=="Not yet captured" else ""
                    bd2.markdown(f"""
                    <div style="margin-bottom:14px;">
                      <div class="detail-label">{fname}</div>
                      <div class="detail-value{ph_class}">{v}</div>
                    </div>""", unsafe_allow_html=True)

                if edit_mode:
                    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                f"letter-spacing:0.06em;text-transform:uppercase;margin:16px 0 10px 0;'>"
                                f"Edit Decision Fields (Future-Ready)</div>", unsafe_allow_html=True)
                    proj_key = str(row.get(proj_id_col, sel_proj))
                    edits = st.session_state["edits"].setdefault(proj_key, {})
                    ff_fields_edit = [
                        ("Business Value",    "bv",  "text"),
                        ("Financial Value",   "fval", "text"),
                        ("Dollars at Risk",   "dar",  "text"),
                        ("Estimated Capacity","cap",  "text"),
                        ("Hard Deadline",     "hd",   "text"),
                        ("Deadline Type",     "dt",   "select",
                         ["","Regulatory","Contractual","Internal","Aspirational"]),
                        ("Dependency Criticality","dc","select",
                         ["","High","Medium","Low","None"]),
                        ("Executive Sponsor", "es",   "text"),
                        ("Confidence Level",  "cl",   "select",
                         ["","High","Medium","Low","Unknown"]),
                        ("Blocker Reason",    "br",   "text"),
                    ]
                    fe1,fe2 = st.columns(2)
                    for i,ff_def in enumerate(ff_fields_edit):
                        fname_f = ff_def[0]; fkey_f = ff_def[1]; ftype_f = ff_def[2]
                        actual_f = fc.get(fname_f)
                        cur_val  = str(row.get(actual_f,"")) if actual_f else ""
                        target   = fe1 if i%2==0 else fe2
                        label_f  = fname_f + ("" if actual_f else " (placeholder)")
                        if ftype_f=="select":
                            opts_f = ff_def[3]
                            idx_f  = opts_f.index(cur_val) if cur_val in opts_f else 0
                            edits[fkey_f] = target.selectbox(
                                label_f, opts_f, index=idx_f, key=f"ffe_{fkey_f}_{proj_key}")
                        else:
                            edits[fkey_f] = target.text_input(
                                label_f, value=cur_val, key=f"ffe_{fkey_f}_{proj_key}",
                                placeholder="Not yet captured")
                    st.session_state["edits"][proj_key] = edits
                    st.markdown("<div class='edit-banner' style='margin-top:10px;'>"
                                "Decision field edits stored in session only. "
                                "Writeback to OneDrive not yet implemented.</div>",
                                unsafe_allow_html=True)

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:36px 0 14px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      RevOps Program Dashboard · Working Team View · Refreshes every 60 s · Source: SharePoint
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  PROJECT CONFIGURATION TAB
# ══════════════════════════════════════════════════════════════
with _tab_config_obj:

    # ── helpers ───────────────────────────────────────────────
    def _safe_num(val, fallback=""):
        try:
            v = pd.to_numeric(val, errors="coerce")
            return "" if pd.isna(v) else v
        except Exception:
            return fallback

    # Work from session-state editable copies
    edit_proj = st.session_state["proj_edits"].copy()
    edit_vm   = st.session_state["vm_edits"].copy() if st.session_state["vm_edits"] is not None else None
    edit_res  = st.session_state["res_edits"].copy() if st.session_state["res_edits"] is not None else None
    edit_dep  = st.session_state["dep_edits"].copy() if st.session_state["dep_edits"] is not None else None

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{C['navy']} 0%,{C['deep_blue']} 100%);
      border-radius:12px;padding:16px 24px;color:white;margin-bottom:18px;">
      <div style="font-size:17px;font-weight:700;">Project Configuration</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.65);margin-top:2px;">
        Edit project fields in-session. Changes are not written back to OneDrive until writeback is implemented.
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div class='edit-banner'>⚠️ All edits are session-only and will not persist across refreshes.</div>",
                unsafe_allow_html=True)

    # ── Filters ──────────────────────────────────────────────
    st.markdown("<div class='section-title'>Filters</div>", unsafe_allow_html=True)
    cf1, cf2, cf3, cf4 = st.columns(4)

    func_opts = ["All"]
    if func_col and func_col in edit_proj.columns:
        func_opts += sorted(edit_proj[func_col].dropna().unique().tolist())
    elif team_col_p and team_col_p in edit_proj.columns:
        func_opts += sorted(edit_proj[team_col_p].dropna().unique().tolist())
    cf_func = cf1.selectbox("Function / Team", func_opts)

    type_opts = ["All"]
    if proj_type_col and proj_type_col in edit_proj.columns:
        type_opts += sorted(edit_proj[proj_type_col].dropna().unique().tolist())
    else:
        type_opts += ["Strategic","Sustaining"]
    cf_type = cf2.selectbox("Project Type", type_opts)

    band_opts = ["All","Top","Middle","Lower","Unranked"]
    cf_band = cf3.selectbox("Priority Band", band_opts)

    cf_search = cf4.text_input("Search Project Name", placeholder="Type to filter…")

    # apply config filters to edit_proj
    cfg_df = edit_proj.copy()
    if cf_func != "All":
        tcol = func_col or team_col_p
        if tcol and tcol in cfg_df.columns:
            cfg_df = cfg_df[cfg_df[tcol]==cf_func]
    if cf_type != "All" and proj_type_col and proj_type_col in cfg_df.columns:
        cfg_df = cfg_df[cfg_df[proj_type_col]==cf_type]
    if cf_band != "All":
        cfg_df = cfg_df[cfg_df["__band__"]==cf_band]
    if cf_search and proj_name_col and proj_name_col in cfg_df.columns:
        cfg_df = cfg_df[cfg_df[proj_name_col].astype(str).str.lower()
                        .str.contains(cf_search.lower(), na=False)]

    st.markdown(f"<div style='font-size:11px;color:{C['gray']};margin-bottom:10px;'>"
                f"<strong>{len(cfg_df)}</strong> projects shown</div>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Main editable table ───────────────────────────────────
    st.markdown("<div class='section-title'>Project Fields</div>", unsafe_allow_html=True)

    editable_fields = []
    display_rename  = {}

    # always show id + name
    if proj_id_col:   editable_fields.append(proj_id_col)
    if proj_name_col: editable_fields.append(proj_name_col)

    # conditionally add columns that exist
    for col_var, label in [
        (proj_type_col,     "Project Type"),
        (status_col,        "Status"),
        (priority_rank_col, "Priority Rank"),
        ("__band__",        "Priority Band"),
        (effort_col,        "Effort Score"),
        (impact_col,        "Impact Score"),
        (biz_val_col,       "Business Value ($)"),
        (dar_proj_col,      "Dollars at Risk ($)"),
        (owner_col,         "Owner"),
    ]:
        if col_var and col_var in cfg_df.columns:
            editable_fields.append(col_var)
            display_rename[col_var] = label

    # deduplicate preserving order
    seen = set(); deduped = []
    for c in editable_fields:
        if c not in seen: deduped.append(c); seen.add(c)
    editable_fields = deduped

    display_df = cfg_df[editable_fields].copy().rename(columns=display_rename)

    # column config for st.data_editor
    col_config = {}
    if "Priority Rank" in display_df.columns:
        col_config["Priority Rank"] = st.column_config.NumberColumn(
            "Priority Rank", min_value=1, step=1, help="Rank for Strategic projects only")
    if "Project Type" in display_df.columns:
        col_config["Project Type"] = st.column_config.SelectboxColumn(
            "Project Type", options=["Strategic","Sustaining","Unknown"])
    if "Status" in display_df.columns:
        all_s = sorted(proj_df[status_col].dropna().unique().tolist()) if status_col else []
        col_config["Status"] = st.column_config.SelectboxColumn("Status", options=all_s or None)
    if "Effort Score" in display_df.columns:
        col_config["Effort Score"] = st.column_config.SelectboxColumn(
            "Effort Score", options=["1","2","3","4","5","S","M","L"])
    if "Impact Score" in display_df.columns:
        col_config["Impact Score"] = st.column_config.NumberColumn(
            "Impact Score", min_value=1, max_value=5, step=1)
    if "Business Value ($)" in display_df.columns:
        col_config["Business Value ($)"] = st.column_config.NumberColumn(
            "Business Value ($)", min_value=0, format="$%d")
    if "Dollars at Risk ($)" in display_df.columns:
        col_config["Dollars at Risk ($)"] = st.column_config.NumberColumn(
            "Dollars at Risk ($)", min_value=0, format="$%d")
    if "Priority Band" in display_df.columns:
        col_config["Priority Band"] = st.column_config.SelectboxColumn(
            "Priority Band", options=["Top","Middle","Lower","Unranked"])

    edited_main = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config=col_config,
        key="cfg_main_editor",
    )

    if st.button("💾  Apply Project Field Edits", key="apply_proj_edits"):
        # Map renamed columns back to originals
        rev_rename = {v:k for k,v in display_rename.items()}
        edited_back = edited_main.rename(columns=rev_rename)
        for col in edited_back.columns:
            if col in edit_proj.columns:
                edit_proj.loc[cfg_df.index, col] = edited_back[col].values
        # Recompute band if rank changed
        edit_proj["__band__"] = assign_priority_band(edit_proj, priority_rank_col)
        st.session_state["proj_edits"] = edit_proj
        st.success("✅ Project fields updated in session.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Value Category mapping ────────────────────────────────
    st.markdown("<div class='section-title'>Value Category Mapping</div>", unsafe_allow_html=True)

    if edit_vm is None or vm_pid_col is None or vm_cat_col is None:
        st.markdown("<div class='ph-note'>ℹ️ Project_Value_Map sheet not found or missing expected columns "
                    "(Project ID, Value Category). Add this sheet to enable category editing.</div>",
                    unsafe_allow_html=True)
    else:
        # Show project selector scoped to current config filter
        vm_proj_ids = cfg_df[proj_id_col].dropna().unique().tolist() if proj_id_col else []
        if not vm_proj_ids:
            st.info("No projects visible under current filters.")
        else:
            vm_sel_pid = st.selectbox("Select Project to Edit Value Categories",
                                      ["— select —"] + [str(p) for p in vm_proj_ids],
                                      key="vm_proj_sel")
            if vm_sel_pid and vm_sel_pid != "— select —":
                cur_cats = edit_vm[edit_vm[vm_pid_col].astype(str)==vm_sel_pid][vm_cat_col].dropna().tolist()
                all_cats = (sorted(val_dict_df[vd_cat_col].dropna().unique().tolist())
                            if val_dict_df is not None and vd_cat_col else
                            sorted(edit_vm[vm_cat_col].dropna().unique().tolist()))
                new_cats = st.multiselect(f"Value Categories for project {vm_sel_pid}",
                                          all_cats, default=cur_cats, key="vm_cats_sel")
                if st.button("💾  Apply Category Changes", key="apply_vm"):
                    # Remove old rows for this pid and insert new ones
                    keep_mask = edit_vm[vm_pid_col].astype(str) != vm_sel_pid
                    kept = edit_vm[keep_mask].copy()
                    new_rows = pd.DataFrame({
                        vm_pid_col: [vm_sel_pid]*len(new_cats),
                        vm_cat_col: new_cats,
                    })
                    if vm_grp_col and val_dict_df is not None and vd_cat_col and vd_grp_col:
                        grp_map = val_dict_df.set_index(vd_cat_col)[vd_grp_col].to_dict()
                        new_rows[vm_grp_col] = new_rows[vm_cat_col].map(grp_map)
                    edit_vm_new = pd.concat([kept, new_rows], ignore_index=True)
                    st.session_state["vm_edits"] = edit_vm_new
                    st.success(f"✅ Value categories updated for project {vm_sel_pid}.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Resource / Function reassignment ─────────────────────
    st.markdown("<div class='section-title'>Resource & Function Assignment</div>",
                unsafe_allow_html=True)

    if edit_res is None or pid_col_r is None:
        st.markdown("<div class='ph-note'>ℹ️ Project_Resources sheet not found.</div>",
                    unsafe_allow_html=True)
    else:
        vm_proj_ids_res = cfg_df[proj_id_col].dropna().unique().tolist() if proj_id_col else []
        if vm_proj_ids_res:
            res_sel_pid = st.selectbox("Select Project to Edit Resources",
                                       ["— select —"] + [str(p) for p in vm_proj_ids_res],
                                       key="res_proj_sel")
            if res_sel_pid and res_sel_pid != "— select —":
                res_rows = edit_res[edit_res[pid_col_r].astype(str)==res_sel_pid].copy()
                all_res_cols = [c for c in edit_res.columns if c != pid_col_r]
                res_col_cfg = {}
                if team_col_r and team_col_r in edit_res.columns:
                    all_teams = sorted(edit_res[team_col_r].dropna().unique().tolist())
                    res_col_cfg[team_col_r] = st.column_config.SelectboxColumn(
                        team_col_r, options=all_teams)
                edited_res_rows = st.data_editor(
                    res_rows[all_res_cols] if all_res_cols else res_rows,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    column_config=res_col_cfg,
                    key="res_editor",
                )
                if st.button("💾  Apply Resource Changes", key="apply_res"):
                    keep_mask = edit_res[pid_col_r].astype(str) != res_sel_pid
                    kept_r = edit_res[keep_mask].copy()
                    edited_res_rows[pid_col_r] = res_sel_pid
                    new_res = pd.concat([kept_r, edited_res_rows], ignore_index=True)
                    st.session_state["res_edits"] = new_res
                    st.success(f"✅ Resources updated for project {res_sel_pid}.")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Dependency reassignment ───────────────────────────────
    st.markdown("<div class='section-title'>Dependency Management</div>", unsafe_allow_html=True)

    if edit_dep is None:
        st.markdown("<div class='ph-note'>ℹ️ Dependencies sheet not found.</div>",
                    unsafe_allow_html=True)
    else:
        dep_pid_col2 = get_col(edit_dep,"Project ID","ProjectID","ID","Dependent Project ID")
        vm_proj_ids_dep = cfg_df[proj_id_col].dropna().unique().tolist() if proj_id_col else []
        if dep_pid_col2 and vm_proj_ids_dep:
            dep_sel_pid = st.selectbox("Select Project to Edit Dependencies",
                                       ["— select —"] + [str(p) for p in vm_proj_ids_dep],
                                       key="dep_proj_sel")
            if dep_sel_pid and dep_sel_pid != "— select —":
                dep_rows = edit_dep[edit_dep[dep_pid_col2].astype(str)==dep_sel_pid].copy()
                all_pids = sorted(edit_proj[proj_id_col].dropna().astype(str).unique().tolist()) if proj_id_col else []
                dep_on_col = get_col(edit_dep,"Depends On","DependsOn","Dependency ID","dependency_id")
                dep_col_cfg = {}
                if dep_on_col and dep_on_col in edit_dep.columns:
                    dep_col_cfg[dep_on_col] = st.column_config.SelectboxColumn(
                        dep_on_col, options=all_pids,
                        help="Select by Project ID — use Project ID, not free text")
                edited_dep_rows = st.data_editor(
                    dep_rows,
                    use_container_width=True,
                    hide_index=True,
                    num_rows="dynamic",
                    column_config=dep_col_cfg,
                    key="dep_editor",
                )
                if st.button("💾  Apply Dependency Changes", key="apply_dep"):
                    keep_mask = edit_dep[dep_pid_col2].astype(str) != dep_sel_pid
                    kept_d = edit_dep[keep_mask].copy()
                    edited_dep_rows[dep_pid_col2] = dep_sel_pid
                    new_dep = pd.concat([kept_d, edited_dep_rows], ignore_index=True)
                    st.session_state["dep_edits"] = new_dep
                    st.success(f"✅ Dependencies updated for project {dep_sel_pid}.")
        elif dep_pid_col2 is None:
            st.warning("Dependencies sheet found but Project ID column is missing.")

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:32px 0 12px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      Project Configuration · Session edits only · Not persisted to OneDrive
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
#  EXECUTIVE DASHBOARD — Value + Resource + Dependency addons
#  (appended inside _tab_exec_obj via a second context block)
# ══════════════════════════════════════════════════════════════
with _tab_exec_obj:

    # Use session-state edited data where available
    _exec_proj = st.session_state["proj_edits"].copy()
    _exec_vm   = st.session_state["vm_edits"].copy() if st.session_state["vm_edits"] is not None else None
    _exec_res  = st.session_state["res_edits"].copy() if st.session_state["res_edits"] is not None else None
    _exec_dep  = st.session_state["dep_edits"].copy() if st.session_state["dep_edits"] is not None else None

    # apply same sidebar filters to exec proj
    _exec_filtered = filtered.copy()  # already filtered by sidebar

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.markdown("<div class='section-title'>Priority Bands</div>", unsafe_allow_html=True)

    # Priority Band grouping
    band_order = ["Top","Middle","Lower","Unranked"]
    band_colors = {"Top":C["teal"],"Middle":C["bright_blue"],"Lower":C["gray"],"Unranked":"#D1D5DB"}

    if "__band__" in _exec_filtered.columns:
        for band in band_order:
            band_df = _exec_filtered[_exec_filtered["__band__"]==band]
            if band_df.empty: continue
            with st.expander(f"**{band} Priority** — {len(band_df)} projects", expanded=(band=="Top")):
                band_cols = [c for c in [proj_id_col,proj_name_col,proj_type_col,owner_col,
                                          status_col,biz_val_col,dar_proj_col,effort_col,impact_col]
                             if c and c in band_df.columns]
                bd = band_df[band_cols].rename(columns={
                    biz_val_col:"Business Value ($)",dar_proj_col:"Dollars at Risk ($)"
                } if biz_val_col or dar_proj_col else {})
                st.dataframe(bd.reset_index(drop=True), use_container_width=True, hide_index=True)

    # Top 5 by Business Value
    if biz_val_col and biz_val_col in _exec_filtered.columns:
        st.markdown("<div class='section-title' style='margin-top:20px;'>Top 5 by Business Value</div>",
                    unsafe_allow_html=True)
        tv = _exec_filtered.copy()
        tv[biz_val_col] = pd.to_numeric(tv[biz_val_col], errors="coerce")
        top5 = tv.nlargest(5, biz_val_col)
        t5_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                                biz_val_col,"__band__"] if c and c in top5.columns]
        st.dataframe(top5[t5_cols].rename(columns={biz_val_col:"Business Value ($)",
                                                    "__band__":"Priority Band"})
                     .reset_index(drop=True), use_container_width=True, hide_index=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Value Breakdown from Project_Value_Map
    st.markdown("<div class='section-title'>Value Breakdown</div>", unsafe_allow_html=True)

    if _exec_vm is not None and vm_pid_col and vm_cat_col:
        # join to filtered project ids
        if proj_id_col:
            valid_pids = _exec_filtered[proj_id_col].dropna().unique()
            vm_f = _exec_vm[_exec_vm[vm_pid_col].isin(valid_pids)]
        else:
            vm_f = _exec_vm.copy()

        val_left, val_right = st.columns(2)

        with val_left:
            st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                        f"text-transform:uppercase;margin-bottom:8px;'>Projects by Value Category</div>",
                        unsafe_allow_html=True)
            cat_counts = vm_f.groupby(vm_cat_col)[vm_pid_col].nunique().reset_index()
            cat_counts.columns = ["Value Category","Projects"]
            cat_counts = cat_counts.sort_values("Projects",ascending=True)
            fig_vc = px.bar(cat_counts, x="Projects", y="Value Category", orientation="h",
                            color="Projects",
                            color_continuous_scale=[[0,C["light_blue"]],[1,C["teal"]]],
                            template="plotly_white")
            fig_vc = chart_layout(fig_vc, height=max(260, len(cat_counts)*28))
            fig_vc.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig_vc, use_container_width=True)

        with val_right:
            if vm_grp_col and vm_grp_col in vm_f.columns:
                st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                            f"text-transform:uppercase;margin-bottom:8px;'>Projects by Value Group</div>",
                            unsafe_allow_html=True)
                grp_counts = vm_f.groupby(vm_grp_col)[vm_pid_col].nunique().reset_index()
                grp_counts.columns = ["Value Group","Projects"]
                fig_vg = px.pie(grp_counts, names="Value Group", values="Projects",
                                color_discrete_sequence=PALETTE, hole=0.50, template="plotly_white")
                fig_vg.update_traces(textposition="outside", textfont_size=10,
                                     marker=dict(line=dict(color="white",width=2)))
                fig_vg.update_layout(height=280, margin=dict(t=8,b=8,l=8,r=8),
                                     showlegend=True, legend=dict(font=dict(size=10)),
                                     paper_bgcolor="white", font=dict(family="Inter, sans-serif"))
                st.plotly_chart(fig_vg, use_container_width=True)

            # Value Category Dictionary reference
            if val_dict_df is not None and vd_cat_col:
                st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                            f"text-transform:uppercase;margin-bottom:6px;margin-top:12px;'>Category Dictionary</div>",
                            unsafe_allow_html=True)
                st.dataframe(val_dict_df.head(20), use_container_width=True, hide_index=True, height=180)
    else:
        st.markdown("<div class='ph-note'>ℹ️ Project_Value_Map sheet not available. "
                    "Add it to enable value breakdown analysis.</div>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Resource Allocation
    st.markdown("<div class='section-title'>Resource Allocation</div>", unsafe_allow_html=True)

    if _exec_res is not None and team_col_r and pid_col_r:
        ra_left, ra_right = st.columns(2)
        valid_pids_r = _exec_filtered[proj_id_col].dropna().unique() if proj_id_col else []
        res_f = _exec_res[_exec_res[pid_col_r].isin(valid_pids_r)] if len(valid_pids_r) else _exec_res
        func_counts = res_f.groupby(team_col_r)[pid_col_r].nunique().reset_index()
        func_counts.columns = ["Function","Projects"]
        func_counts = func_counts.sort_values("Projects",ascending=False)
        with ra_left:
            st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                        f"text-transform:uppercase;margin-bottom:8px;'>Projects by Function</div>",
                        unsafe_allow_html=True)
            fig_ra = px.bar(func_counts.sort_values("Projects",ascending=True),
                            x="Projects",y="Function",orientation="h",
                            color="Projects",
                            color_continuous_scale=[[0,C["light_blue"]],[1,C["deep_blue"]]],
                            template="plotly_white")
            fig_ra = chart_layout(fig_ra, height=max(220, len(func_counts)*30))
            fig_ra.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig_ra, use_container_width=True)
        with ra_right:
            st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                        f"text-transform:uppercase;margin-bottom:8px;'>Function Load Table</div>",
                        unsafe_allow_html=True)
            avg_load = res_f.groupby(team_col_r)[pid_col_r].nunique().mean()
            func_counts["Load Status"] = func_counts["Projects"].apply(
                lambda x: "⚠️ Overloaded" if x > avg_load*1.3 else
                          ("✅ Normal" if x <= avg_load else "🔶 Elevated"))
            st.dataframe(func_counts, use_container_width=True, hide_index=True, height=260)
    else:
        st.markdown("<div class='ph-note'>ℹ️ Resource data unavailable.</div>",
                    unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Dependency Risk
    st.markdown("<div class='section-title'>Dependency Risk</div>", unsafe_allow_html=True)

    if _exec_dep is not None:
        dep_pid_col_e = get_col(_exec_dep,"Project ID","ProjectID","ID","Dependent Project ID")
        dep_on_col_e  = get_col(_exec_dep,"Depends On","DependsOn","Dependency ID","dependency_id")
        if dep_pid_col_e:
            valid_pids_d  = _exec_filtered[proj_id_col].dropna().unique() if proj_id_col else []
            dep_f = _exec_dep[_exec_dep[dep_pid_col_e].isin(valid_pids_d)]
            dep_l, dep_r = st.columns(2)
            with dep_l:
                dep_count = dep_f.groupby(dep_pid_col_e).size().reset_index()
                dep_count.columns = ["Project ID","Dependency Count"]
                dep_count = dep_count.sort_values("Dependency Count",ascending=False)
                st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                            f"text-transform:uppercase;margin-bottom:8px;'>Projects by Dependency Count</div>",
                            unsafe_allow_html=True)
                fig_dep = px.bar(dep_count.head(15).sort_values("Dependency Count",ascending=True),
                                 x="Dependency Count", y="Project ID", orientation="h",
                                 color="Dependency Count",
                                 color_continuous_scale=[[0,"#FEF3C7"],[0.5,"#E67E22"],[1,"#C0392B"]],
                                 template="plotly_white")
                fig_dep = chart_layout(fig_dep, height=260)
                fig_dep.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig_dep, use_container_width=True)
            with dep_r:
                # highlight bottlenecks: projects that are depended on by many others
                if dep_on_col_e and dep_on_col_e in dep_f.columns:
                    blocking = dep_f.groupby(dep_on_col_e).size().reset_index()
                    blocking.columns = ["Blocking Project ID","Blocked By N Projects"]
                    blocking = blocking.sort_values("Blocked By N Projects",ascending=False).head(10)
                    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                                f"text-transform:uppercase;margin-bottom:8px;'>Bottleneck Projects</div>",
                                unsafe_allow_html=True)
                    st.dataframe(blocking, use_container_width=True, hide_index=True, height=240)
                else:
                    st.dataframe(dep_count, use_container_width=True, hide_index=True, height=240)
        else:
            st.info("Dependencies sheet found but Project ID column missing.")
    else:
        st.markdown("<div class='ph-note'>ℹ️ Dependencies sheet not available.</div>",
                    unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # AI Insights (rule-based)
    st.markdown("<div class='section-title'>AI Insights</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;color:#6B7280;margin-bottom:12px;'>"
                "Rule-based insights derived from current portfolio data. "
                "No ML model — logic driven by thresholds and rankings.</div>",
                unsafe_allow_html=True)

    ins_left, ins_right = st.columns(2)

    with ins_left:
        # 1. Top 3 value-driving categories
        if _exec_vm is not None and vm_pid_col and vm_cat_col and proj_id_col:
            valid_pids_ai = _exec_filtered[proj_id_col].dropna().unique()
            vm_ai = _exec_vm[_exec_vm[vm_pid_col].isin(valid_pids_ai)]
            top_cats = (vm_ai.groupby(vm_cat_col)[vm_pid_col].nunique()
                        .sort_values(ascending=False).head(3))
            if not top_cats.empty:
                top_cat_lines = "; ".join([f"<strong>{c}</strong> ({n} projects)"
                                           for c,n in top_cats.items()])
                st.markdown(f"""
                <div class="insight-box">
                  <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
                    text-transform:uppercase;margin-bottom:6px;">Top Value-Driving Categories</div>
                  <div style="font-size:13px;color:#374151;">{top_cat_lines}</div>
                </div>""", unsafe_allow_html=True)

        # 2. Overloaded functions
        if _exec_res is not None and team_col_r and pid_col_r and proj_id_col:
            valid_pids_ai2 = _exec_filtered[proj_id_col].dropna().unique()
            res_ai = _exec_res[_exec_res[pid_col_r].isin(valid_pids_ai2)]
            if not res_ai.empty:
                func_ai = res_ai.groupby(team_col_r)[pid_col_r].nunique()
                mean_load = func_ai.mean()
                overloaded = func_ai[func_ai > mean_load * 1.3].sort_values(ascending=False)
                if not overloaded.empty:
                    ol_lines = "; ".join([f"<strong>{f}</strong> ({n} projects)"
                                          for f,n in overloaded.head(3).items()])
                    st.markdown(f"""
                    <div class="insight-box">
                      <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
                        text-transform:uppercase;margin-bottom:6px;">Overloaded Functions</div>
                      <div style="font-size:13px;color:#374151;">{ol_lines}</div>
                    </div>""", unsafe_allow_html=True)

    with ins_right:
        # 3. High-value projects not in Top Band
        if biz_val_col and biz_val_col in _exec_filtered.columns and "__band__" in _exec_filtered.columns:
            bv_vals = pd.to_numeric(_exec_filtered[biz_val_col], errors="coerce")
            if bv_vals.notna().any():
                bv_thresh = bv_vals.quantile(0.75) if bv_vals.notna().sum()>3 else bv_vals.max()
                hi_val_not_top = _exec_filtered[
                    (bv_vals >= bv_thresh) & (_exec_filtered["__band__"] != "Top")]
                if not hi_val_not_top.empty:
                    n_hv = len(hi_val_not_top)
                    names = ", ".join(hi_val_not_top[proj_name_col].head(3).tolist()) if proj_name_col else f"{n_hv} projects"
                    st.markdown(f"""
                    <div class="insight-box" style="border-left:3px solid #E67E22;">
                      <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
                        text-transform:uppercase;margin-bottom:6px;">High-Value Not in Top Band</div>
                      <div style="font-size:13px;color:#374151;">
                        <strong>{n_hv}</strong> high-value project{'s' if n_hv!=1 else ''} ranked outside Top:
                        {names}{'…' if n_hv>3 else ''}.
                        Consider re-evaluating priority bands.
                      </div>
                    </div>""", unsafe_allow_html=True)

        # 4. Dependencies but low priority
        if _exec_dep is not None and proj_id_col and "__band__" in _exec_filtered.columns:
            dep_pid_col_ins = get_col(_exec_dep,"Project ID","ProjectID","ID","Dependent Project ID")
            if dep_pid_col_ins:
                dep_pids = set(_exec_dep[dep_pid_col_ins].dropna().astype(str).unique())
                low_pri_dep = _exec_filtered[
                    (_exec_filtered[proj_id_col].astype(str).isin(dep_pids)) &
                    (_exec_filtered["__band__"].isin(["Lower","Unranked"]))]
                if not low_pri_dep.empty:
                    n_ld = len(low_pri_dep)
                    st.markdown(f"""
                    <div class="insight-box" style="border-left:3px solid #C0392B;">
                      <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
                        text-transform:uppercase;margin-bottom:6px;">Low-Priority Projects with Dependencies</div>
                      <div style="font-size:13px;color:#374151;">
                        <strong>{n_ld}</strong> project{'s' if n_ld!=1 else ''} {'have' if n_ld!=1 else 'has'}
                        dependencies but sit in Lower or Unranked bands.
                        These may create downstream bottlenecks.
                      </div>
                    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:32px 0 12px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      RevOps Program Dashboard · Executive Dashboard · Source: SharePoint · Refreshes every 60 s
    </div>""", unsafe_allow_html=True)
