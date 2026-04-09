import hashlib
import re
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# ─────────────────────────────────────────────
ONEDRIVE_FILE_URL = "https://emerson-my.sharepoint.com/:x:/p/savitri_lazarus/IQAQPOe1joHSTopYQHg4L61vAdgWzYvAdfVUHhZGNiI6TAM?e=YsNeJD"
# ─────────────────────────────────────────────

C = {
    "deep_blue":   "#004B8D",
    "green":       "#00573D",
    "navy":        "#1B2552",
    "bright_blue": "#1DB1DE",
    "soft_green":  "#7CCF8B",
    "teal":        "#00AD7C",
    "light_blue":  "#75D3EB",
    "gray":        "#9FA1A4",
    "black":       "#000000",
    "white":       "#FFFFFF",
}
PALETTE = [C["deep_blue"], C["bright_blue"], C["teal"], C["soft_green"],
           C["navy"], C["light_blue"], C["green"], C["gray"]]
STATUS_COLORS = {
    "Delayed":     "#C0392B",
    "At Risk":     "#E67E22",
    "On Track":    C["teal"],
    "Active":      C["deep_blue"],
    "In Progress": C["bright_blue"],
    "Complete":    C["soft_green"],
    "Completed":   C["soft_green"],
    "Not Started": C["gray"],
    "Planning":    C["light_blue"],
}

st.set_page_config(page_title="RevOps Program Dashboard",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{{font-family:'Inter',sans-serif;}}
.main{{background-color:#F7F8FA;}}
.block-container{{padding:1.6rem 2.2rem 2rem 2.2rem;max-width:1440px;}}
.kpi-wrap{{background:{C["white"]};border-radius:12px;padding:18px 20px 14px 20px;
  border:1px solid #E8ECF0;box-shadow:0 2px 8px rgba(0,0,0,0.05);
  min-height:108px;display:flex;flex-direction:column;justify-content:space-between;}}
.kpi-label{{font-size:10px;font-weight:700;letter-spacing:0.07em;text-transform:uppercase;
  color:{C["gray"]};margin-bottom:2px;}}
.kpi-value{{font-size:32px;font-weight:700;color:{C["navy"]};line-height:1;}}
.kpi-value.danger{{color:#C0392B;}} .kpi-value.success{{color:{C["teal"]};}}
.kpi-value.warn{{color:#D97706;}}
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
.detail-card{{background:{C["white"]};border-radius:10px;padding:16px 18px;
  border:1px solid #E8ECF0;margin-bottom:10px;}}
.detail-label{{font-size:10px;font-weight:700;letter-spacing:0.05em;text-transform:uppercase;
  color:{C["gray"]};margin-bottom:2px;}}
.detail-value{{font-size:13px;font-weight:500;color:{C["navy"]};}}
.status-badge{{display:inline-block;padding:2px 9px;border-radius:20px;
  font-size:10px;font-weight:700;letter-spacing:0.04em;}}
.empty-state{{text-align:center;padding:40px 24px;color:{C["gray"]};font-size:13px;}}
.edit-banner{{background:#FEF3C7;border:1px solid #F59E0B;border-radius:8px;
  padding:10px 16px;font-size:12px;color:#92400E;margin-bottom:14px;}}
.insight-box{{background:{C["white"]};border-radius:10px;padding:16px 18px;
  border:1px solid #E8ECF0;margin-bottom:12px;}}
[data-testid="stSidebar"]{{background:{C["white"]};border-right:1px solid #E8ECF0;}}
#MainMenu{{visibility:hidden;}}footer{{visibility:hidden;}}header{{visibility:hidden;}}
.stTabs [data-baseweb="tab-list"]{{gap:4px;background:#F0F2F5;border-radius:8px;padding:4px;}}
.stTabs [data-baseweb="tab"]{{border-radius:6px;padding:5px 14px;font-size:12px;font-weight:500;}}
.stTabs [aria-selected="true"]{{background:{C["white"]};color:{C["navy"]};}}
</style>
""", unsafe_allow_html=True)

# ─── UTILITIES ────────────────────────────────
def normalize_cols(df):
    df.columns = [c.strip() for c in df.columns]
    return df

def get_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    for c in candidates:
        for col in df.columns:
            if c.lower().replace(" ", "").replace("?", "") in col.lower().replace(" ", "").replace("?", ""):
                return col
    return None

def build_download_url(url):
    if "/:x:/p/" in url or "/:x:/s/" in url:
        sep = "&" if "?" in url else "?"
        return url + sep + "download=1"
    if "_layouts/15/Doc.aspx" in url:
        m = re.search(r'sourcedoc=%7B([^%]+)%7D', url, re.IGNORECASE)
        if m:
            base = url.split("/_layouts/")[0]
            return f"{base}/_layouts/15/download.aspx?UniqueId={m.group(1)}"
    if "1drv.ms" in url:
        try:
            r = requests.get(url, allow_redirects=True, timeout=15)
            sep = "&" if "?" in r.url else "?"
            return r.url + sep + "download=1"
        except Exception:
            pass
    return url + ("&" if "?" in url else "?") + "download=1"

def chart_layout(fig, height=300, legend=False):
    fig.update_layout(
        height=height, margin=dict(t=14, b=14, l=8, r=8),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Inter, sans-serif", size=11, color="#374151"),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02,
                    xanchor="right", x=1, font=dict(size=10)) if legend else {},
        xaxis=dict(gridcolor="#F0F2F5", linecolor="#E8ECF0", tickfont=dict(size=10)),
        yaxis=dict(gridcolor="#F0F2F5", linecolor="#E8ECF0", tickfont=dict(size=10)),
    )
    fig.update_traces(marker_line_width=0)
    return fig

def status_badge_html(s):
    cm = {
        "delayed":("#FEE2E2","#C0392B"), "at risk":("#FEF3C7","#D97706"),
        "on track":("#D1FAE5","#065F46"), "active":("#DBEAFE","#1E40AF"),
        "in progress":("#E0F2FE","#0369A1"), "complete":("#D1FAE5","#065F46"),
        "completed":("#D1FAE5","#065F46"), "not started":("#F3F4F6","#374151"),
        "planning":("#EDE9FE","#5B21B6"),
    }
    bg, fg = cm.get(str(s).lower(), ("#F3F4F6","#374151"))
    return f"<span class='status-badge' style='background:{bg};color:{fg};'>{s}</span>"

def normalize_cdm(val):
    if pd.isna(val) or str(val).strip() == "":
        return "Unknown"
    v = str(val).strip().lower()
    if v in ("yes","y","true","1"): return "Yes"
    if v in ("no","n","false","0"): return "No"
    return "Unknown"

def deterministic_jitter(series, scale=0.18):
    """Stable jitter using hash of string representation."""
    def _jitter(v):
        h = int(hashlib.md5(str(v).encode()).hexdigest(), 16)
        return ((h % 1000) / 1000.0 - 0.5) * 2 * scale
    return series.apply(_jitter)

def kpi_card(col, label, value, sub, color_class="", accent=None):
    a = accent or C["deep_blue"]
    col.markdown(f"""
    <div class="kpi-wrap">
      <div>
        <div class="kpi-accent-bar" style="background:{a};width:28px;"></div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color_class}">{value}</div>
      </div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

# ─── DATA LOAD ────────────────────────────────
@st.cache_data(ttl=60)
def load_data(url):
    try:
        dl = build_download_url(url)
        hdr = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(dl, headers=hdr, timeout=30, allow_redirects=True)
        r.raise_for_status()
        if "html" in r.headers.get("Content-Type","").lower():
            fb = url + ("&download=1" if "?" in url else "?download=1")
            r = requests.get(fb, headers=hdr, timeout=30, allow_redirects=True)
            r.raise_for_status()
        content = BytesIO(r.content)
        sheets = {}
        for s in ["Projects","Project_Resources","Dependencies"]:
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

proj_df = sheets.get("Projects")
res_df  = sheets.get("Project_Resources")
dep_df  = sheets.get("Dependencies")

for m in [s for s,d in sheets.items() if d is None]:
    st.warning(f"Sheet '{m}' could not be loaded.")

if proj_df is None:
    st.error("Projects sheet is required.")
    st.stop()

# ─── COLUMN MAP ───────────────────────────────
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
priority_rank_col  = get_col(proj_df,"Priority Rank","Rank","priority_rank")
cdm_col_raw        = get_col(proj_df,"Dependent on CDM Project?","Dependent on CDM Project",
                              "CDM Dependency","CDM Project","CDM","DependentonCDMProject")

CDM_COL = "__cdm__"
proj_df[CDM_COL] = proj_df[cdm_col_raw].apply(normalize_cdm) if cdm_col_raw else "Unknown"

team_col_r = None
pid_col_r  = None
if res_df is not None:
    team_col_r = get_col(res_df,"Team","team","Department","Resource Team")
    pid_col_r  = get_col(res_df,"Project ID","ProjectID","ID")

# ─── SESSION STATE ────────────────────────────
for k,v in [("view","Executive Summary"),("edit_mode",False),("edits",{}),
             ("drill_team",None),("drill_status",None),("drill_cdm",None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ─── SIDEBAR ──────────────────────────────────
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

    sel_teams_sb = []
    if team_col_p:
        teams_all = sorted(proj_df[team_col_p].dropna().unique().tolist())
        sel_teams_sb = st.multiselect("Team", teams_all, default=[])
        if sel_teams_sb:
            base = base[base[team_col_p].isin(sel_teams_sb)]

    if status_col:
        statuses_all = sorted(proj_df[status_col].dropna().unique().tolist())
        sel_status_sb = st.multiselect("Status", statuses_all, default=[])
        if sel_status_sb:
            base = base[base[status_col].isin(sel_status_sb)]

    if cycle_col:
        cycles_all = sorted(proj_df[cycle_col].dropna().unique().tolist())
        sel_cycles_sb = st.multiselect("Cycle", cycles_all, default=[])
        if sel_cycles_sb:
            base = base[base[cycle_col].isin(sel_cycles_sb)]

    if priority_col:
        pris_all = sorted(proj_df[priority_col].dropna().unique().tolist())
        sel_pris_sb = st.multiselect("Priority Type", pris_all, default=[])
        if sel_pris_sb:
            base = base[base[priority_col].isin(sel_pris_sb)]

    sel_cdm_sb = st.multiselect("CDM Dependency", ["Yes","No","Unknown"], default=[],
                                 help="Filter by Dependent on CDM Project status")
    if sel_cdm_sb:
        base = base[base[CDM_COL].isin(sel_cdm_sb)]

    st.markdown("<hr style='border:none;border-top:1px solid #E8ECF0;margin:14px 0;'>",
                unsafe_allow_html=True)
    st.caption(f"**{len(base)}** of {len(proj_df)} projects shown")
    st.caption("Auto-refreshes every 60 s")

filtered = base.copy()

# ─── SHARED METRICS ───────────────────────────
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

# ─── SCATTER HELPER (jitter) ──────────────────
def scatter_effort_impact(df, height=300, show_legend=True):
    if not effort_col or not impact_col:
        st.warning("Effort / Impact columns not found.")
        return
    keep = [c for c in [proj_id_col, proj_name_col, owner_col, status_col,
                         effort_col, impact_col, CDM_COL] if c]
    sdf = df[keep].copy()
    sdf[effort_col] = pd.to_numeric(sdf[effort_col], errors="coerce")
    sdf[impact_col] = pd.to_numeric(sdf[impact_col], errors="coerce")
    sdf = sdf.dropna(subset=[effort_col, impact_col]).reset_index(drop=True)
    if sdf.empty:
        st.markdown("<div class='empty-state'>No numeric data for scatter plot.</div>",
                    unsafe_allow_html=True)
        return

    # deterministic jitter using row index + value hash
    seed_e = sdf.apply(lambda r: f"{r.name}_{r[effort_col]}", axis=1)
    seed_i = sdf.apply(lambda r: f"{r.name}_{r[impact_col]}", axis=1)
    sdf["__x__"] = sdf[effort_col] + deterministic_jitter(seed_e, scale=0.15)
    sdf["__y__"] = sdf[impact_col] + deterministic_jitter(seed_i, scale=0.15)

    custom_cols = [c for c in [proj_id_col, proj_name_col, owner_col,
                                effort_col, impact_col, CDM_COL] if c]

    fig = px.scatter(
        sdf, x="__x__", y="__y__",
        color=status_col if status_col else None,
        color_discrete_map=STATUS_COLORS,
        custom_data=custom_cols,
        template="plotly_white",
        opacity=0.78,
    )
    # Build hover template
    ht_lines = []
    for i, c in enumerate(custom_cols):
        label = "CDM" if c == CDM_COL else c
        ht_lines.append(f"<b>{label}:</b> %{{customdata[{i}]}}")
    hover_template = "<br>".join(ht_lines) + "<extra></extra>"
    fig.update_traces(
        marker=dict(size=11, line=dict(width=1.5, color="white")),
        hovertemplate=hover_template,
    )
    fig.update_layout(
        xaxis_title=effort_col or "Effort",
        yaxis_title=impact_col or "Impact",
    )
    fig = chart_layout(fig, height=height, legend=show_legend)
    st.plotly_chart(fig, use_container_width=True)


# ─── VIEW TOGGLE ──────────────────────────────
t_col, _ = st.columns([2.4, 5])
with t_col:
    view = st.radio("", ["Executive Summary","Working Team View"],
                    horizontal=True, label_visibility="collapsed",
                    index=0 if st.session_state["view"]=="Executive Summary" else 1)
    st.session_state["view"] = view


# ══════════════════════════════════════════════
#  EXECUTIVE SUMMARY VIEW
# ══════════════════════════════════════════════
if view == "Executive Summary":

    st.markdown(f"""
    <div class="exec-header">
      <h1>RevOps Program Dashboard</h1>
      <div class="subtitle">Resource load, project risk, and dependency visibility</div>
      <div class="dynamic">{dynamic_summary}</div>
    </div>""", unsafe_allow_html=True)

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    kpi_card(k1,"Total Projects",   total,             "in current filters",     accent=C["navy"])
    kpi_card(k2,"Delayed Projects", delayed_count,     "require attention",
             color_class="danger" if delayed_count else "",
             accent="#C0392B" if delayed_count else C["gray"])
    kpi_card(k3,"Active Projects",  active_count,      "in progress",            accent=C["deep_blue"])
    kpi_card(k4,"Teams Involved",   teams_count,       "across resource pool",   accent=C["teal"])
    kpi_card(k5,"CDM Dependent",    cdm_yes_count,     "depend on CDM",
             color_class="warn" if cdm_yes_count else "", accent="#D97706")
    kpi_card(k6,"Unknown CDM",      cdm_unknown_count, "CDM status not set",
             color_class="warn" if cdm_unknown_count else "", accent="#D97706")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    if total == 0:
        st.markdown("<div class='empty-state'>No projects match the current filters.</div>",
                    unsafe_allow_html=True)
        st.stop()

    st.markdown("<div class='section-title'>Executive Overview</div>", unsafe_allow_html=True)
    st.markdown("")

    ch1, ch2 = st.columns(2)
    with ch1:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Top Teams by Project Load</div>",
                    unsafe_allow_html=True)
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            cr = res_df[res_df[pid_col_r].isin(ap)]
            if not cr.empty:
                tc = (cr.groupby(team_col_r)[pid_col_r].nunique()
                        .reset_index().rename(columns={team_col_r:"Team", pid_col_r:"Projects"})
                        .sort_values("Projects", ascending=True).tail(10))
                fig = px.bar(tc, x="Projects", y="Team", orientation="h",
                             color="Projects",
                             color_continuous_scale=[[0,C["light_blue"]],[1,C["deep_blue"]]],
                             template="plotly_white")
                fig = chart_layout(fig, height=270)
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.markdown("<div class='empty-state'>No resource data.</div>", unsafe_allow_html=True)
        else:
            st.warning("Resource data unavailable.")

    with ch2:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Portfolio Status</div>",
                    unsafe_allow_html=True)
        if status_col:
            sc = (filtered[status_col].value_counts().reset_index()
                  .rename(columns={status_col:"Status","count":"Count"}))
            if "Count" not in sc.columns:
                sc.columns = ["Status","Count"]
            fig2 = px.bar(sc.sort_values("Count",ascending=False),
                          x="Status", y="Count", color="Status",
                          color_discrete_map=STATUS_COLORS, template="plotly_white")
            fig2 = chart_layout(fig2, height=270)
            fig2.update_layout(showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)
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
                            f"delayed + high-impact</div>", unsafe_allow_html=True); shown+=1
        if effort_col and impact_col:
            hi_both = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi_both.empty:
                st.markdown(f"<div class='risk-item warn'>🔶 <strong>{len(hi_both)}</strong> "
                            f"high-effort + high-impact</div>", unsafe_allow_html=True); shown+=1
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            tdf = res_df[res_df[pid_col_r].isin(ap)]
            if not tdf.empty:
                grp = tdf.groupby(team_col_r)[pid_col_r].nunique()
                tt = grp.idxmax(); tv = int(grp.max())
                st.markdown(f"<div class='risk-item warn'>📌 <strong>{tt}</strong> — "
                            f"highest load ({tv} projects)</div>",
                            unsafe_allow_html=True); shown+=1
        if cdm_yes_count:
            st.markdown(f"<div class='risk-item info'>🔗 <strong>{cdm_yes_count}</strong> "
                        f"project{'s' if cdm_yes_count!=1 else ''} CDM-dependent</div>",
                        unsafe_allow_html=True); shown+=1
        if cdm_unknown_count:
            st.markdown(f"<div class='risk-item info'>❓ <strong>{cdm_unknown_count}</strong> "
                        f"unknown CDM dependency</div>", unsafe_allow_html=True); shown+=1
        if shown==0:
            st.markdown("<div class='risk-item ok'>✅ No critical risks identified.</div>",
                        unsafe_allow_html=True)

        # Narrative
        narrative_parts = []
        if delayed_count: narrative_parts.append(f"{delayed_count} delayed project{'s' if delayed_count!=1 else ''}")
        if cdm_yes_count: narrative_parts.append(f"{cdm_yes_count} CDM-dependent")
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap2 = filtered[proj_id_col].dropna().unique()
            tdf2 = res_df[res_df[pid_col_r].isin(ap2)]
            if not tdf2.empty:
                grp2 = tdf2.groupby(team_col_r)[pid_col_r].nunique()
                narrative_parts.append(f"{grp2.idxmax()} carrying the most load")
        if narrative_parts:
            narrative = "Key portfolio risks: " + "; ".join(narrative_parts) + "."
            st.markdown(f"<div style='margin-top:10px;font-size:11px;color:#374151;"
                        f"line-height:1.5;'>{narrative}</div>", unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Spotlight table
    st.markdown("<div class='section-title'>Project Spotlight</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:12px;color:#6B7280;margin-bottom:10px;'>"
                "Top projects ranked by delay status, impact, and effort.</div>",
                unsafe_allow_html=True)
    spot = filtered.copy()
    spot["__score__"] = 0
    if status_col:
        spot["__score__"] += spot[status_col].str.lower().str.contains("delay",na=False).astype(int)*10
    if impact_col:
        spot["__score__"] += pd.to_numeric(spot[impact_col],errors="coerce").fillna(0)
    if effort_col:
        spot["__score__"] += pd.to_numeric(spot[effort_col],errors="coerce").fillna(0)*0.5
    spot = spot.sort_values("__score__",ascending=False).head(15)
    sp_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                            cycle_col,impact_col,effort_col,CDM_COL,delayed_impact_col] if c]
    disp = spot[sp_cols].rename(columns={CDM_COL:"CDM Dependency"}).reset_index(drop=True)
    st.dataframe(disp, use_container_width=True, hide_index=True)

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:36px 0 14px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      RevOps Program Dashboard · Executive View · Refreshes every 60 s · Source: SharePoint
    </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════
#  WORKING TEAM VIEW
# ══════════════════════════════════════════════
else:
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{C['green']} 0%,{C['teal']} 100%);
      border-radius:12px;padding:18px 24px;color:white;margin-bottom:20px;">
      <div style="font-size:18px;font-weight:700;">RevOps Working Team View</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.65);margin-top:3px;">{dynamic_summary}</div>
    </div>""", unsafe_allow_html=True)

    em_col, _ = st.columns([2,6])
    with em_col:
        edit_mode = st.toggle("✏️ Edit Mode", value=st.session_state["edit_mode"])
        st.session_state["edit_mode"] = edit_mode
    if edit_mode:
        st.markdown("<div class='edit-banner'>⚠️ <strong>Edit Mode ON.</strong> "
                    "Edits are stored in session only — not saved to OneDrive.</div>",
                    unsafe_allow_html=True)

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    kpi_card(k1,"Total Projects",   total,             "in current filters",    accent=C["navy"])
    kpi_card(k2,"Delayed Projects", delayed_count,     "require attention",
             color_class="danger" if delayed_count else "",
             accent="#C0392B" if delayed_count else C["gray"])
    kpi_card(k3,"Active Projects",  active_count,      "in progress",           accent=C["deep_blue"])
    kpi_card(k4,"Teams Involved",   teams_count,       "across resource pool",  accent=C["teal"])
    kpi_card(k5,"CDM Dependent",    cdm_yes_count,     "depend on CDM",
             color_class="warn" if cdm_yes_count else "", accent="#D97706")
    kpi_card(k6,"Unknown CDM",      cdm_unknown_count, "CDM status not set",
             color_class="warn" if cdm_unknown_count else "", accent="#D97706")

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    if total == 0:
        st.markdown("<div class='empty-state'>No projects match the current filters.</div>",
                    unsafe_allow_html=True)
        st.stop()

    # ── Portfolio Analysis ────────────────────
    st.markdown("<div class='section-title'>Portfolio Analysis</div>", unsafe_allow_html=True)
    st.markdown("")

    pa1, pa2 = st.columns(2)
    with pa1:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Projects by Team</div>",
                    unsafe_allow_html=True)
        team_drill_opts = ["All Teams"]
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            cr = res_df[res_df[pid_col_r].isin(ap)]
            if not cr.empty:
                tc = (cr.groupby(team_col_r)[pid_col_r].nunique()
                        .reset_index().rename(columns={team_col_r:"Team",pid_col_r:"Projects"})
                        .sort_values("Projects",ascending=True))
                team_drill_opts = ["All Teams"] + tc["Team"].tolist()
                fig = px.bar(tc, x="Projects", y="Team", orientation="h",
                             color="Projects",
                             color_continuous_scale=[[0,C["light_blue"]],[1,C["deep_blue"]]],
                             template="plotly_white")
                fig = chart_layout(fig, height=300)
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Resource data unavailable.")
        drill_team = st.selectbox("🔍 Drill into team", team_drill_opts, key="drill_team_sel")
        st.session_state["drill_team"] = None if drill_team=="All Teams" else drill_team

    with pa2:
        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                    f"text-transform:uppercase;margin-bottom:8px;'>Projects by Status</div>",
                    unsafe_allow_html=True)
        status_drill_opts = ["All Statuses"]
        if status_col:
            sc = (filtered[status_col].value_counts().reset_index()
                  .rename(columns={status_col:"Status","count":"Count"}))
            if "Count" not in sc.columns: sc.columns = ["Status","Count"]
            sc = sc.sort_values("Count",ascending=False)
            status_drill_opts = ["All Statuses"] + sc["Status"].tolist()
            fig2 = px.bar(sc, x="Status", y="Count", color="Status",
                          color_discrete_map=STATUS_COLORS, template="plotly_white")
            fig2 = chart_layout(fig2, height=300)
            fig2.update_layout(showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("Status column not found.")
        drill_status = st.selectbox("🔍 Drill into status", status_drill_opts, key="drill_status_sel")
        st.session_state["drill_status"] = None if drill_status=="All Statuses" else drill_status

    pa3, pa4 = st.columns(2)
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
            pc = (filtered[priority_col].value_counts().reset_index()
                  .rename(columns={priority_col:"Priority","count":"Count"}))
            if "Count" not in pc.columns: pc.columns = ["Priority","Count"]
            fig4 = px.pie(pc, names="Priority", values="Count",
                          color_discrete_sequence=PALETTE, hole=0.50, template="plotly_white")
            fig4.update_traces(textposition="outside", textfont_size=10,
                               marker=dict(line=dict(color="white",width=2)))
            fig4.update_layout(height=260, margin=dict(t=8,b=8,l=8,r=8),
                                showlegend=True, legend=dict(font=dict(size=10)),
                                paper_bgcolor="white", font=dict(family="Inter, sans-serif"))
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.info("Priority column not found.")

    # Delayed table
    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;"
                f"text-transform:uppercase;margin-bottom:8px;margin-top:4px;'>Delayed Projects</div>",
                unsafe_allow_html=True)
    if status_col:
        del_df = filtered[delayed_mask].copy()
        dcols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                              cycle_col,impact_col,effort_col,delayed_impact_col,CDM_COL] if c]
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
    cdm_color_map = {"Yes":"#D97706","No":C["teal"],"Unknown":C["gray"]}
    fig_cdm = px.bar(cdm_counts, x="CDM Status", y="Count", color="CDM Status",
                     color_discrete_map=cdm_color_map, template="plotly_white")
    fig_cdm = chart_layout(fig_cdm, height=220)
    fig_cdm.update_layout(showlegend=False)
    st.plotly_chart(fig_cdm, use_container_width=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Interactive Insights ──────────────────
    st.markdown("<div class='section-title'>Interactive Insights</div>", unsafe_allow_html=True)
    st.markdown("")

    ins1, ins2, ins3, ins4 = st.columns(4)
    with ins1:
        ins_teams = ["All Teams"]
        if team_col_p:
            ins_teams += sorted(filtered[team_col_p].dropna().unique().tolist())
        elif res_df is not None and team_col_r:
            ins_teams += sorted(res_df[team_col_r].dropna().unique().tolist())
        ins_team = st.selectbox("Filter by Team", ins_teams)
    with ins2:
        ins_statuses = ["All Statuses"]
        if status_col:
            ins_statuses += sorted(filtered[status_col].dropna().unique().tolist())
        ins_status = st.selectbox("Filter by Status", ins_statuses)
    with ins3:
        ins_cdm = st.selectbox("Filter by CDM", ["All","Yes","No","Unknown"])
    with ins4:
        proj_opts = ["All Projects"]
        if proj_name_col:
            proj_opts += sorted(filtered[proj_name_col].dropna().unique().tolist())
        ins_proj = st.selectbox("Filter by Project", proj_opts)

    # apply interactive filters
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
                f"<strong>{len(ins_df)}</strong> projects match the selected filters.</div>",
                unsafe_allow_html=True)

    if not ins_df.empty:
        ins_left, ins_right = st.columns([3,2])
        with ins_left:
            ins_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                                     cycle_col,impact_col,effort_col,CDM_COL] if c]
            st.dataframe(ins_df[ins_cols].rename(columns={CDM_COL:"CDM Dependency"})
                         .reset_index(drop=True),
                         use_container_width=True, hide_index=True, height=280)

        with ins_right:
            # Impacted teams
            if res_df is not None and team_col_r and pid_col_r and proj_id_col:
                ap_ins = ins_df[proj_id_col].dropna().unique()
                rf_ins = res_df[res_df[pid_col_r].isin(ap_ins)]
                if not rf_ins.empty and team_col_r:
                    imp_teams = rf_ins[team_col_r].value_counts().head(8).reset_index()
                    imp_teams.columns = ["Team","Projects"]
                    st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                f"letter-spacing:0.06em;text-transform:uppercase;margin-bottom:6px;'>"
                                f"Impacted Teams</div>", unsafe_allow_html=True)
                    fig_it = px.bar(imp_teams.sort_values("Projects",ascending=True),
                                    x="Projects", y="Team", orientation="h",
                                    color="Projects",
                                    color_continuous_scale=[[0,C["soft_green"]],[1,C["green"]]],
                                    template="plotly_white")
                    fig_it = chart_layout(fig_it, height=220)
                    fig_it.update_layout(coloraxis_showscale=False)
                    st.plotly_chart(fig_it, use_container_width=True)

            # Dependencies preview
            if dep_df is not None and proj_id_col:
                dep_pid = get_col(dep_df,"Project ID","ProjectID","ID","Dependent Project ID")
                if dep_pid:
                    ap_ins2 = ins_df[proj_id_col].dropna().unique()
                    dep_match = dep_df[dep_df[dep_pid].isin(ap_ins2)]
                    if not dep_match.empty:
                        st.markdown(f"<div style='font-size:11px;font-weight:700;color:{C['gray']};"
                                    f"letter-spacing:0.06em;text-transform:uppercase;"
                                    f"margin-bottom:6px;margin-top:10px;'>Related Dependencies</div>",
                                    unsafe_allow_html=True)
                        st.dataframe(dep_match.head(10).reset_index(drop=True),
                                     use_container_width=True, hide_index=True, height=140)

    else:
        st.markdown("<div class='empty-state'>No projects match this combination of filters.</div>",
                    unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Risk Insights ─────────────────────────
    st.markdown("<div class='section-title'>Risk Insights</div>", unsafe_allow_html=True)
    st.markdown("")

    ri1, ri2 = st.columns([2,3])
    with ri1:
        r_shown = 0
        if delayed_count:
            st.markdown(f"<div class='risk-item'>⚠️ <strong>{delayed_count}</strong> "
                        f"delayed project{'s' if delayed_count!=1 else ''}</div>",
                        unsafe_allow_html=True); r_shown+=1
        if status_col and impact_col:
            hid = filtered[delayed_mask & (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hid.empty:
                st.markdown(f"<div class='risk-item'>🔴 <strong>{len(hid)}</strong> "
                            f"delayed + high-impact</div>", unsafe_allow_html=True); r_shown+=1
        if effort_col and impact_col:
            hi_both = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi_both.empty:
                st.markdown(f"<div class='risk-item warn'>🔶 <strong>{len(hi_both)}</strong> "
                            f"high-effort + high-impact</div>", unsafe_allow_html=True); r_shown+=1
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap = filtered[proj_id_col].dropna().unique()
            tdf = res_df[res_df[pid_col_r].isin(ap)]
            if not tdf.empty:
                grp = tdf.groupby(team_col_r)[pid_col_r].nunique()
                tt = grp.idxmax(); tv = int(grp.max())
                st.markdown(f"<div class='risk-item warn'>📌 <strong>{tt}</strong> — "
                            f"highest load ({tv} projects)</div>", unsafe_allow_html=True); r_shown+=1
        if cdm_yes_count:
            st.markdown(f"<div class='risk-item info'>🔗 <strong>{cdm_yes_count}</strong> "
                        f"CDM-dependent project{'s' if cdm_yes_count!=1 else ''}</div>",
                        unsafe_allow_html=True); r_shown+=1
        if cdm_unknown_count:
            st.markdown(f"<div class='risk-item info'>❓ <strong>{cdm_unknown_count}</strong> "
                        f"unknown CDM dependency</div>", unsafe_allow_html=True); r_shown+=1
        if priority_col:
            high_pri = filtered[filtered[priority_col].astype(str).str.lower().str.contains(
                "high|critical|p1", na=False)]
            if not high_pri.empty:
                st.markdown(f"<div class='risk-item info'>🔵 <strong>{len(high_pri)}</strong> "
                            f"high-priority project{'s' if len(high_pri)!=1 else ''}</div>",
                            unsafe_allow_html=True); r_shown+=1
        if r_shown==0:
            st.markdown("<div class='risk-item ok'>✅ No critical risks identified.</div>",
                        unsafe_allow_html=True)

    with ri2:
        # Narrative
        n_parts = []
        if delayed_count: n_parts.append(f"{delayed_count} delayed project{'s' if delayed_count!=1 else ''}")
        if status_col and impact_col:
            hid2 = filtered[delayed_mask & (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hid2.empty: n_parts.append(f"{len(hid2)} with high impact")
        if effort_col and impact_col:
            hi2 = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)]
            if not hi2.empty: n_parts.append(f"{len(hi2)} requiring both high effort and high impact")
        if cdm_yes_count: n_parts.append(f"{cdm_yes_count} dependent on CDM delivery")
        if cdm_unknown_count: n_parts.append(f"{cdm_unknown_count} with unresolved CDM dependency status")
        if res_df is not None and team_col_r and pid_col_r and proj_id_col:
            ap3 = filtered[proj_id_col].dropna().unique()
            tdf3 = res_df[res_df[pid_col_r].isin(ap3)]
            if not tdf3.empty:
                grp3 = tdf3.groupby(team_col_r)[pid_col_r].nunique()
                n_parts.append(f"{grp3.idxmax()} carrying the highest team load at {int(grp3.max())} projects")
        if n_parts:
            narrative = ("The current portfolio shows: " + "; ".join(n_parts) + ". "
                         "Review prioritisation and resource allocation to reduce delivery risk.")
        else:
            narrative = "Portfolio appears on track with no major risks flagged under current filters."
        st.markdown(f"""
        <div class="insight-box">
          <div style="font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
            text-transform:uppercase;margin-bottom:8px;">Dynamic Risk Narrative</div>
          <div style="font-size:13px;color:#374151;line-height:1.6;">{narrative}</div>
        </div>""", unsafe_allow_html=True)

        # Effort/Impact risk table
        if effort_col and impact_col:
            hi_risk = filtered[
                (pd.to_numeric(filtered[effort_col],errors="coerce")>=4) &
                (pd.to_numeric(filtered[impact_col],errors="coerce")>=4)
            ].copy()
            if not hi_risk.empty:
                st.markdown(f"""
                <div style='font-size:11px;font-weight:700;color:{C['gray']};letter-spacing:0.06em;
                  text-transform:uppercase;margin-bottom:6px;margin-top:10px;'>
                  High Effort + High Impact Projects</div>""", unsafe_allow_html=True)
                hr_cols = [c for c in [proj_id_col,proj_name_col,owner_col,status_col,
                                        effort_col,impact_col,CDM_COL] if c]
                st.dataframe(hi_risk[hr_cols].rename(columns={CDM_COL:"CDM Dependency"})
                             .reset_index(drop=True),
                             use_container_width=True, hide_index=True, height=160)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # ── Project Explorer ──────────────────────
    st.markdown("<div class='section-title'>Project Explorer</div>", unsafe_allow_html=True)
    st.markdown("")

    if proj_name_col:
        proj_opts = sorted(filtered[proj_name_col].dropna().unique().tolist())
    elif proj_id_col:
        proj_opts = sorted(filtered[proj_id_col].dropna().astype(str).unique().tolist())
    else:
        proj_opts = []

    if not proj_opts:
        st.markdown("<div class='empty-state'>No projects available.</div>", unsafe_allow_html=True)
    else:
        sel_proj = st.selectbox("Select a project to explore", proj_opts,
                                help="Search or select a project for full detail.")
        prow_df = (filtered[filtered[proj_name_col]==sel_proj] if proj_name_col
                   else filtered[filtered[proj_id_col].astype(str)==sel_proj])

        if not prow_df.empty:
            row = prow_df.iloc[0]
            proj_status = row.get(status_col,"") if status_col else ""
            is_delayed  = "delay" in str(proj_status).lower()
            badge       = status_badge_html(proj_status) if proj_status else ""
            pid_disp    = (f"<span style='color:{C['gray']};font-size:12px;margin-left:8px;'>"
                           f"{row.get(proj_id_col,'')}</span>" if proj_id_col else "")
            cdm_val     = row.get(CDM_COL,"Unknown")
            cdm_color   = {"Yes":"#D97706","No":C["teal"],"Unknown":C["gray"]}.get(cdm_val,C["gray"])

            st.markdown(f"""
            <div class="detail-card" style="border-left:4px solid {'#C0392B' if is_delayed else C['deep_blue']};">
              <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <span style="font-size:16px;font-weight:700;color:{C['navy']};">{sel_proj}</span>
                {pid_disp} {badge}
                <span class='status-badge' style='background:#FEF3C7;color:{cdm_color};'>
                  CDM: {cdm_val}</span>
              </div>
              <div style="margin-top:5px;font-size:11px;color:#C0392B;">
                {'⚠️ Flagged as delayed.' if is_delayed else ''}
              </div>
            </div>""", unsafe_allow_html=True)

            tab1, tab2, tab3, tab4 = st.tabs(
                ["Overview","Resources","Dependencies","Risk & Impact"])

            with tab1:
                meta = [c for c in [proj_id_col,owner_col,team_col_p,status_col,
                                     cycle_col,priority_col,effort_col,impact_col,CDM_COL] if c]
                if meta:
                    pairs = [(c, row.get(c,"—")) for c in meta]
                    half  = len(pairs)//2 + len(pairs)%2
                    m1,m2 = st.columns(2)
                    for cn,cv in pairs[:half]:
                        lbl = "CDM Dependency" if cn==CDM_COL else cn
                        m1.markdown(f"""
                        <div style="margin-bottom:12px;">
                          <div class="detail-label">{lbl}</div>
                          <div class="detail-value">{cv if cv==cv and cv!='' else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    for cn,cv in pairs[half:]:
                        lbl = "CDM Dependency" if cn==CDM_COL else cn
                        m2.markdown(f"""
                        <div style="margin-bottom:12px;">
                          <div class="detail-label">{lbl}</div>
                          <div class="detail-value">{cv if cv==cv and cv!='' else '—'}</div>
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
                                f"Edit Fields</div>", unsafe_allow_html=True)
                    proj_key = str(row.get(proj_id_col, sel_proj))
                    edits = st.session_state["edits"].setdefault(proj_key, {})
                    e1,e2 = st.columns(2)
                    if status_col:
                        all_s = sorted(proj_df[status_col].dropna().unique().tolist())
                        edits["status"] = e1.selectbox(
                            "Status", all_s,
                            index=all_s.index(row.get(status_col)) if row.get(status_col) in all_s else 0,
                            key=f"edit_status_{proj_key}")
                    if owner_col:
                        edits["owner"] = e2.text_input(
                            "Owner", value=str(row.get(owner_col,"")),
                            key=f"edit_owner_{proj_key}")
                    if impact_col:
                        edits["impact"] = e1.text_input(
                            "Impact", value=str(row.get(impact_col,"")),
                            key=f"edit_impact_{proj_key}")
                    if effort_col:
                        edits["effort"] = e2.text_input(
                            "Effort", value=str(row.get(effort_col,"")),
                            key=f"edit_effort_{proj_key}")
                    if notes_col:
                        edits["notes"] = st.text_area(
                            "Notes", value=str(row.get(notes_col,"")),
                            key=f"edit_notes_{proj_key}")
                    st.session_state["edits"][proj_key] = edits
                    st.markdown("<div class='edit-banner' style='margin-top:8px;'>"
                                "Session edits stored. Writeback to OneDrive not yet implemented.</div>",
                                unsafe_allow_html=True)

            with tab2:
                if res_df is not None and proj_id_col and pid_col_r:
                    pid_v = row.get(proj_id_col)
                    pr = res_df[res_df[pid_col_r]==pid_v]
                    if not pr.empty:
                        st.dataframe(pr.reset_index(drop=True),
                                     use_container_width=True, hide_index=True)
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

            with tab3:
                if dep_df is not None:
                    dep_pid_col = get_col(dep_df,"Project ID","ProjectID","ID","Dependent Project ID")
                    if proj_id_col and dep_pid_col:
                        pid_v = row.get(proj_id_col)
                        pd_dep = dep_df[dep_df[dep_pid_col]==pid_v]
                        if not pd_dep.empty:
                            st.dataframe(pd_dep.reset_index(drop=True),
                                         use_container_width=True, hide_index=True)
                        else:
                            st.markdown("<div class='empty-state'>No dependencies recorded.</div>",
                                        unsafe_allow_html=True)
                    else:
                        st.warning("Project ID column missing in Dependencies.")
                else:
                    st.warning("Dependencies sheet not available.")

            with tab4:
                r1,r2 = st.columns(2)
                with r1:
                    if impact_col:
                        iv = pd.to_numeric(row.get(impact_col), errors="coerce")
                        ic = C["teal"] if pd.notna(iv) and iv>=3 else C["gray"]
                        st.markdown(f"""
                        <div class="detail-card">
                          <div class="detail-label">Impact Score</div>
                          <div class="kpi-value" style="font-size:28px;color:{ic};">
                            {iv if pd.notna(iv) else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    if effort_col:
                        ev = pd.to_numeric(row.get(effort_col), errors="coerce")
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
                        div = row.get(delayed_impact_col,"—")
                        st.markdown(f"""
                        <div class="detail-card" style="border-left:3px solid #C0392B;">
                          <div class="detail-label">If Delayed Impact</div>
                          <div style="font-size:13px;font-weight:500;color:#C0392B;margin-top:3px;">
                            {div if div==div else '—'}</div>
                        </div>""", unsafe_allow_html=True)
                    if is_delayed:
                        st.markdown(f"""
                        <div class="detail-card" style="background:#FEF3F2;border-left:3px solid #C0392B;">
                          <div style="font-size:12px;color:#C0392B;font-weight:700;">⚠️ Delay Flag Active</div>
                          <div style="font-size:11px;color:#374151;margin-top:3px;">
                            Review ownership and blockers.</div>
                        </div>""", unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div class="detail-card" style="background:#F0FFF8;border-left:3px solid {C['teal']};">
                          <div style="font-size:12px;color:{C['teal']};font-weight:700;">✅ No Delay Flag</div>
                          <div style="font-size:11px;color:#374151;margin-top:3px;">Not currently delayed.</div>
                        </div>""", unsafe_allow_html=True)
                    if notes_col and row.get(notes_col):
                        st.markdown(f"""
                        <div class="detail-card" style="background:#FFFBF0;border-left:3px solid #D97706;">
                          <div class="detail-label">Risk Notes</div>
                          <div style="font-size:11px;color:#374151;margin-top:3px;">{row.get(notes_col)}</div>
                        </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <hr style='border:none;border-top:1px solid #E8ECF0;margin:36px 0 14px 0;'>
    <div style='font-size:10px;color:{C["gray"]};text-align:center;padding-bottom:10px;'>
      RevOps Program Dashboard · Working Team View · Refreshes every 60 s · Source: SharePoint
    </div>""", unsafe_allow_html=True)
