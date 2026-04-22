"""
RevOps Program Dashboard  —  v3
Read-only live feed · refreshes every 12 hours · click any widget to drill in.
"""
import re, hashlib
from collections import defaultdict
from io import BytesIO
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# ── CONFIG ────────────────────────────────────────────────────
ONEDRIVE_FILE_URL = (
    "https://emerson-my.sharepoint.com/:x:/p/savitri_lazarus/"
    "IQAQPOe1joHSTopYQHg4L61vAdgWzYvAdfVUHhZGNiI6TAM?e=YsNeJD"
)
REFRESH_TTL = 43200   # 12 h = twice daily
EDIT_LINK   = ONEDRIVE_FILE_URL

# ── PALETTE ───────────────────────────────────────────────────
C = dict(
    navy="#1B2552", blue="#004B8D", teal="#00AD7C", lblue="#1DB1DE",
    sgreen="#7CCF8B", lgreen="#75D3EB", green="#00573D", gray="#9FA1A4",
    red="#C0392B", amber="#D97706", white="#FFFFFF", bg="#F4F6FA",
)
STATUS_COLOR = {
    "Delayed":C["red"],"At Risk":C["amber"],"On Track":C["teal"],
    "Active":C["blue"],"In Progress":C["lblue"],"Complete":C["sgreen"],
    "Completed":C["sgreen"],"Not Started":C["gray"],"Planning":C["lgreen"],
}
BAND_COLOR = {
    "Top Priority":C["teal"],"Middle Priority":C["blue"],
    "Lower Priority":C["lblue"],"N/A":C["gray"],
}

# ── PAGE ──────────────────────────────────────────────────────
st.set_page_config(page_title="RevOps Program Dashboard",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{{font-family:'DM Sans',sans-serif;background:{C['bg']};}}
.main{{background:{C['bg']};}}
.block-container{{padding:1.2rem 1.8rem 2rem;max-width:1520px;}}
.kpi{{background:{C['white']};border-radius:10px;padding:14px 16px 10px;
  border:1px solid #E2E8F2;box-shadow:0 1px 6px rgba(0,0,0,0.05);}}
.kpi-val{{font-size:28px;font-weight:700;line-height:1;margin:5px 0 3px;}}
.kpi-lbl{{font-size:10px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:{C['gray']};}}
.kpi-sub{{font-size:10px;color:{C['gray']};margin-top:2px;}}
.kpi-bar{{height:3px;border-radius:2px;margin-bottom:6px;}}
.sec{{font-size:11px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;
  color:{C['navy']};border-bottom:2px solid {C['blue']};padding-bottom:5px;
  display:inline-block;margin-bottom:12px;}}
.badge{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700;}}
.drill-panel{{background:{C['white']};border:2px solid {C['blue']};
  border-radius:10px;padding:16px 18px;margin-top:14px;}}
.drill-hdr{{font-size:13px;font-weight:700;color:{C['navy']};margin-bottom:12px;}}
.drill-row{{background:#FAFBFD;border-left:3px solid {C['blue']};
  border-radius:0 6px 6px 0;padding:8px 12px;margin-bottom:5px;}}
.drill-row.delayed{{border-left-color:{C['red']};}}
.drill-row.cdm{{border-left-color:{C['amber']};}}
.edit-cta{{background:#EFF6FF;border:1px solid {C['lblue']};border-radius:8px;
  padding:10px 16px;font-size:12px;color:{C['navy']};margin:10px 0;}}
hr.slim{{border:none;border-top:1px solid #E2E8F2;margin:16px 0;}}
.detail-lbl{{font-size:10px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;
  color:{C['gray']};margin-bottom:2px;}}
.detail-val{{font-size:13px;font-weight:500;color:{C['navy']};margin-bottom:10px;}}
.detail-val.ph{{color:{C['gray']};font-style:italic;font-weight:400;}}
.stTabs [data-baseweb="tab-list"]{{gap:4px;background:#EAEEF4;border-radius:8px;padding:4px;}}
.stTabs [data-baseweb="tab"]{{border-radius:6px;font-size:12px;font-weight:500;padding:5px 14px;}}
.stTabs [aria-selected="true"]{{background:{C['white']};color:{C['navy']};}}
#MainMenu{{visibility:hidden;}}footer{{visibility:hidden;}}header{{visibility:hidden;}}
</style>
""", unsafe_allow_html=True)

# ── HELPERS ───────────────────────────────────────────────────
def nc(df, *candidates):
    for c in candidates:
        if c in df.columns: return c
    for c in candidates:
        for col in df.columns:
            if c.lower().replace(" ","").replace("?","") in \
               col.lower().replace(" ","").replace("?",""): return col
    return None

def build_dl_url(url):
    if "/:x:/p/" in url or "/:x:/s/" in url:
        return url + ("&" if "?" in url else "?") + "download=1"
    if "_layouts/15/Doc.aspx" in url:
        m = re.search(r'sourcedoc=%7B([^%]+)%7D', url, re.I)
        if m:
            return url.split("/_layouts/")[0] + \
                   f"/_layouts/15/download.aspx?UniqueId={m.group(1)}"
    return url + ("&" if "?" in url else "?") + "download=1"

def norm_cdm(v):
    if pd.isna(v) or str(v).strip() == "": return "Unknown"
    s = str(v).strip().upper()
    if s in ("Y","YES","TRUE","1"): return "Yes"
    if s in ("N","NO","FALSE","0"): return "No"
    return "Unknown"

def sbadge(s, sz=10):
    cm = {
        "delayed":("#FEE2E2","#C0392B"), "at risk":("#FEF3C7","#D97706"),
        "on track":("#D1FAE5","#065F46"), "active":("#DBEAFE","#1E40AF"),
        "in progress":("#E0F2FE","#0369A1"), "complete":("#D1FAE5","#065F46"),
        "completed":("#D1FAE5","#065F46"), "not started":("#F3F4F6","#374151"),
        "planning":("#EDE9FE","#5B21B6"),
    }
    bg,fg = cm.get(str(s).lower(),("#F3F4F6","#374151"))
    return f"<span class='badge' style='background:{bg};color:{fg};font-size:{sz}px'>{s}</span>"

def fmt_money(v):
    try:
        n = float(v)
        if n>=1e6: return f"${n/1e6:.1f}M"
        if n>=1e3: return f"${n/1e3:.0f}K"
        return f"${n:.0f}"
    except: return "—"

def chart_base(fig, h=240):
    fig.update_layout(
        height=h, margin=dict(t=10,b=10,l=6,r=6),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="DM Sans",size=11,color="#374151"),
        showlegend=False,
        xaxis=dict(gridcolor="#F0F4F8",linecolor="#E2E8F2",tickfont=dict(size=10)),
        yaxis=dict(gridcolor="#F0F4F8",linecolor="#E2E8F2",tickfont=dict(size=10)),
    )
    fig.update_traces(marker_line_width=0)
    return fig

# ── DATA LOAD (12-hour cache) ─────────────────────────────────
@st.cache_data(ttl=REFRESH_TTL, show_spinner=False)
def load_all(url):
    try:
        r = requests.get(build_dl_url(url),
                         headers={"User-Agent":"Mozilla/5.0"},
                         timeout=30, allow_redirects=True)
        r.raise_for_status()
        if "html" in r.headers.get("Content-Type","").lower():
            r = requests.get(url+("&download=1" if "?" in url else "?download=1"),
                             headers={"User-Agent":"Mozilla/5.0"},
                             timeout=30, allow_redirects=True)
            r.raise_for_status()
        buf = BytesIO(r.content)
        out = {}
        for s in ["Projects","Project_Resources","Dependencies",
                  "Project_Value_Map","Value_Category_Dictionary"]:
            try:
                df = pd.read_excel(buf, sheet_name=s, engine="openpyxl")
                df.columns = [c.strip() for c in df.columns]
                out[s] = df
            except: out[s] = None
        return out, None, datetime.now()
    except Exception as e:
        return None, str(e), datetime.now()

with st.spinner("Loading…"):
    sheets, err, loaded_at = load_all(ONEDRIVE_FILE_URL)

if err:
    st.error(f"Load failed: {err}")
    st.info("Ensure the SharePoint link is set to 'Anyone with the link can view'.")
    st.stop()

proj_df = sheets["Projects"]
res_df  = sheets["Project_Resources"]
dep_df  = sheets["Dependencies"]
vm_df   = sheets["Project_Value_Map"]

if proj_df is None:
    st.error("Projects sheet missing."); st.stop()

# ── COLUMN DETECTION ──────────────────────────────────────────
pid_c    = nc(proj_df,"Project ID","ProjectID","ID")
name_c   = nc(proj_df,"Project","Project Name","Name")
rank_c   = nc(proj_df,"Priority Rank","PriorityRank","Rank")
type_c   = nc(proj_df,"Priority Type","ProjectType","Type")
strat_c  = nc(proj_df,"Strategic Priority","FLMC Tag","Strategic")
owner_c  = nc(proj_df,"Owner","owner")
core_c   = nc(proj_df,"Core Team","CoreTeam","Requested By")
status_c = nc(proj_df,"Status","status")
cycle_c  = nc(proj_df,"Cycle","cycle")
effort_c = nc(proj_df,"Effort","effort")
impact_c = nc(proj_df,"Impact","impact")
invest_c = nc(proj_df,"Investment","investment")
delay_c  = nc(proj_df,"Delayed Flag","Delayed","delay_flag")
deli_c   = nc(proj_df,"If Delayed Impact","Delayed Impact")
band_c   = nc(proj_df,"Priority Band","Band")
cdm_src  = nc(proj_df,"CDM Dependency Flag","CDM Dependency","CDM")
bizprog_c= nc(proj_df,"Business Program","BizProg")
bv_c     = nc(proj_df,"Business Value ($)","Business Value","BizValue")
dar_c    = nc(proj_df,"Dollars at Risk ($)","Dollars at Risk","DAR")
rawval_c = nc(proj_df,"Raw Value Description","RawValue")
valgrp_c = nc(proj_df,"Value Groups","ValueGroups")

res_pid_c  = nc(res_df,"Project ID","ProjectID","ID") if res_df is not None else None
res_team_c = nc(res_df,"Team","team") if res_df is not None else None
dep_pid_c  = nc(dep_df,"Project ID","ProjectID","ID") if dep_df is not None else None
dep_on_c   = nc(dep_df,"Depends On Project ID","DependsOn","dependency") if dep_df is not None else None

# normalise CDM into a stable column
CDM = "__cdm__"
proj_df[CDM] = proj_df[cdm_src].apply(norm_cdm) if cdm_src else "Unknown"

# resource & dependency lookups
proj_teams = defaultdict(list)
if res_df is not None and res_pid_c and res_team_c:
    for _, r in res_df.iterrows():
        if pd.notna(r[res_pid_c]) and pd.notna(r[res_team_c]):
            proj_teams[str(r[res_pid_c])].append(str(r[res_team_c]))

proj_deps = defaultdict(list)
if dep_df is not None and dep_pid_c and dep_on_c:
    for _, r in dep_df.iterrows():
        if pd.notna(r[dep_pid_c]) and pd.notna(r[dep_on_c]):
            for d in str(r[dep_on_c]).replace(";",",").split(","):
                if d.strip(): proj_deps[str(r[dep_pid_c])].append(d.strip())

# restrict to RevOps owner
revops_df = (proj_df[proj_df[owner_c].str.strip()=="RevOps"].copy()
             if owner_c else proj_df.copy())

# ── PRECOMPUTED METRICS ───────────────────────────────────────
total       = len(revops_df)
delayed_m   = (revops_df[delay_c].str.strip().str.upper()=="Y"
               if delay_c else pd.Series([False]*total, index=revops_df.index))
delayed_n   = int(delayed_m.sum())
strategic_n = int((revops_df[type_c].str.strip()=="Strategic").sum()) if type_c else 0
sustaining_n= int((revops_df[type_c].str.strip()=="Sustaining").sum()) if type_c else 0
cdm_yes_n   = int((revops_df[CDM]=="Yes").sum())
flmc_n      = int(revops_df[strat_c].str.contains("FLMC",na=False).sum()) if strat_c else 0
top_total   = int((revops_df[band_c]=="Top Priority").sum()) if band_c else 0
top_delayed = int(((revops_df[band_c]=="Top Priority") & delayed_m).sum()) if band_c else 0
strat_del   = int(((revops_df[type_c]=="Strategic") & delayed_m).sum()) if type_c else 0
pct_delayed = round(delayed_n/total*100) if total else 0

# ── SESSION STATE ─────────────────────────────────────────────
# Store a filter SPEC (dict) — never a DataFrame — so session state serialises cleanly.
# spec keys: "kind" + relevant value
# e.g. {"kind":"status","value":"Delayed"} or {"kind":"all"} or {"kind":"cdm","value":"Yes"}
if "drill_spec" not in st.session_state:
    st.session_state["drill_spec"] = None
if "drill_label" not in st.session_state:
    st.session_state["drill_label"] = ""
if "selected_pid" not in st.session_state:
    st.session_state["selected_pid"] = None

def set_drill(spec: dict, label: str):
    st.session_state["drill_spec"]  = spec
    st.session_state["drill_label"] = label

def clear_drill():
    st.session_state["drill_spec"]  = None
    st.session_state["drill_label"] = ""

def apply_spec(df, spec):
    """Return subset of df matching spec."""
    if spec is None: return None
    k = spec.get("kind","")
    v = spec.get("value","")
    if k == "all":          return df
    if k == "status"  and status_c: return df[df[status_c]==v]
    if k == "band"    and band_c:   return df[df[band_c]==v]
    if k == "cdm":          return df[df[CDM]==v]
    if k == "type"    and type_c:   return df[df[type_c]==v]
    if k == "delayed" and delay_c:  return df[df[delay_c].str.strip().str.upper()=="Y"]
    if k == "flmc"    and strat_c:  return df[df[strat_c].str.contains("FLMC",na=False)]
    if k == "strat_delayed" and type_c and delay_c:
        return df[(df[type_c]=="Strategic") &
                  (df[delay_c].str.strip().str.upper()=="Y")]
    if k == "top_priority" and band_c:
        return df[df[band_c]=="Top Priority"]
    if k == "team" and pid_c:
        pids = set(str(p) for p,teams in proj_teams.items() if v in teams)
        return df[df[pid_c].astype(str).isin(pids)]
    if k == "value_group" and pid_c and vm_df is not None:
        vm_pid = nc(vm_df,"Project ID","ProjectID","ID")
        vm_grp = nc(vm_df,"Value Group","Group")
        if vm_pid and vm_grp:
            pids = set(vm_df[vm_df[vm_grp]==v][vm_pid].astype(str).unique())
            return df[df[pid_c].astype(str).isin(pids)]
    return df

def render_drill_panel():
    """Render the drill-down panel from session state spec."""
    spec  = st.session_state.get("drill_spec")
    label = st.session_state.get("drill_label","")
    if spec is None: return

    df = apply_spec(revops_df, spec)
    if df is None or df.empty:
        st.info("No projects match this selection.")
        return

    del_n = int((df[delay_c].str.upper()=="Y").sum()) if delay_c else 0
    ok_n  = int((df[status_c].str.lower()=="on track").sum()) if status_c else 0

    hdr_col, btn_col = st.columns([8,1])
    with hdr_col:
        st.markdown(f"""
        <div style='font-size:13px;font-weight:700;color:{C['navy']};margin-bottom:10px;'>
          📂 {label}
          <span style='font-size:11px;font-weight:400;color:{C['gray']};margin-left:10px;'>
            {len(df)} project{'s' if len(df)!=1 else ''}
            {'  ·  ⚠ '+str(del_n)+' delayed' if del_n else ''}
            {'  ·  ✓ '+str(ok_n)+' on track' if ok_n else ''}
          </span>
        </div>""", unsafe_allow_html=True)
    with btn_col:
        if st.button("✕ Clear", key="clear_drill_btn"):
            clear_drill()
            st.rerun()

    for _, row in df.iterrows():
        pid   = str(row[pid_c])   if pid_c   else "—"
        pname = str(row[name_c])  if name_c  else "—"
        stat  = str(row[status_c])if status_c else "—"
        band  = str(row[band_c])  if band_c  else ""
        cdm   = row[CDM]
        what  = (str(row[rawval_c]) if rawval_c and pd.notna(row[rawval_c])
                 and str(row[rawval_c]) not in ("None","nan","") else "")
        is_del= str(row[delay_c]).upper()=="Y" if delay_c else False
        bc    = BAND_COLOR.get(band, C["gray"])
        border= C["red"] if is_del else (C["amber"] if cdm=="Yes" else C["blue"])

        st.markdown(f"""
        <div style='border-left:3px solid {border};background:#FAFBFD;
          border-radius:0 7px 7px 0;padding:9px 14px;margin-bottom:5px;'>
          <div style='display:flex;align-items:center;gap:8px;flex-wrap:wrap;'>
            <span style='font-family:DM Mono;font-size:10px;color:{C['gray']};'>{pid}</span>
            <span style='font-size:12px;font-weight:600;color:{C['navy']};'>{pname}</span>
            {sbadge(stat,10)}
            <span style='font-size:9px;font-weight:600;color:{bc};'>
              {band.replace(" Priority","")}</span>
            {'<span class="badge" style="background:#FEF3C7;color:#D97706;font-size:9px;">⚠ CDM</span>'
             if cdm=="Yes" else ""}
          </div>
          {f'<div style="font-size:10px;color:#555;margin-top:3px;">{what[:110]}{"…" if len(what)>110 else ""}</div>'
           if what else ""}
        </div>""", unsafe_allow_html=True)

# ── PAGE HEADER ───────────────────────────────────────────────
hc, rc = st.columns([8,1])
with hc:
    st.markdown(f"""
    <div style='display:flex;align-items:baseline;gap:12px;'>
      <span style='font-size:20px;font-weight:700;color:{C["navy"]};'>
        RevOps Program Dashboard</span>
      <span style='font-size:12px;color:{C["gray"]};'>
        FY26 · Owner = RevOps · Read-only live view</span>
    </div>""", unsafe_allow_html=True)
with rc:
    if st.button("↺ Refresh", help="Force reload from OneDrive"):
        st.cache_data.clear(); st.rerun()

st.markdown(
    f"<div style='font-size:10px;color:{C['gray']};text-align:right;padding:2px 0 4px;'>"
    f"Last loaded {loaded_at.strftime('%b %d %Y %I:%M %p')} · "
    f"auto-refresh every 12 h · "
    f"<a href='{EDIT_LINK}' target='_blank' "
    f"style='color:{C['blue']};font-weight:600;'>✏️ Edit in Excel →</a></div>",
    unsafe_allow_html=True)

st.markdown("<hr class='slim'>", unsafe_allow_html=True)

vc,_ = st.columns([3,5])
with vc:
    view = st.radio("",
                    ["📋 Executive Summary","📊 Portfolio Detail","🔍 Project Explorer"],
                    horizontal=True, label_visibility="collapsed", key="main_view")

st.markdown("<hr class='slim'>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# ██  EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════
if view == "📋 Executive Summary":

    # KPI cards — one button per card sets drill spec
    KPI = [
        ("Total Projects",    str(total),        "Owner = RevOps",
         C["navy"],  {"kind":"all"},             "All RevOps Projects"),
        ("Strategic",         str(strategic_n),  "Project type: Strategic",
         C["teal"],  {"kind":"type","value":"Strategic"}, "Strategic Projects"),
        ("Sustaining",        str(sustaining_n), "Project type: Sustaining",
         C["lblue"], {"kind":"type","value":"Sustaining"},"Sustaining Projects"),
        ("Delayed",           f"⚠ {delayed_n}",  f"{pct_delayed}% of portfolio",
         C["red"],   {"kind":"delayed"},         "Delayed Projects"),
        ("CDM Dependent",     str(cdm_yes_n),    "Blocked pending P11",
         C["amber"], {"kind":"cdm","value":"Yes"},"CDM-Dependent Projects"),
        ("FLMC SoaP Aligned", str(flmc_n),       "FLMC Strategy on a Page",
         C["navy"],  {"kind":"flmc"},            "FLMC SoaP Projects"),
    ]
    k = st.columns(6)
    for col,(lbl,val,sub,color,spec,dlbl) in zip(k, KPI):
        col.markdown(f"""
        <div class='kpi'>
          <div class='kpi-bar' style='background:{color}'></div>
          <div class='kpi-lbl'>{lbl}</div>
          <div class='kpi-val' style='color:{color}'>{val}</div>
          <div class='kpi-sub'>{sub}</div>
        </div>""", unsafe_allow_html=True)
        if col.button(f"View →", key=f"kpi_{lbl}", use_container_width=True):
            set_drill(spec, dlbl)
            st.rerun()

    # Cost of inaction strip
    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='background:linear-gradient(135deg,{C['navy']} 0%,{C['blue']} 100%);
      border-radius:10px;padding:12px 18px 4px;margin-bottom:2px;'>
      <div style='font-size:10px;font-weight:700;letter-spacing:.09em;color:rgba(255,255,255,.55);
        text-transform:uppercase;margin-bottom:8px;'>
        COST OF INACTION — IF NO DECISIONS ARE MADE THIS QUARTER</div>
    </div>""", unsafe_allow_html=True)

    IA = [
        (f"{pct_delayed}%","of RevOps portfolio delayed",  f"{delayed_n} of {total} programs",
         C["red"],  {"kind":"delayed"},                    "Delayed Projects"),
        (f"1 of {top_total}","Top Priority on track",       f"{top_delayed} of {top_total} already delayed",
         C["amber"],{"kind":"top_priority"},               "Top Priority Projects"),
        (str(cdm_yes_n),"programs blocked on CDM",          "Cannot start or advance",
         C["amber"],{"kind":"cdm","value":"Yes"},          "CDM-Dependent Projects"),
        (str(strat_del),"Strategic programs slipping",       f"{strat_del} of {strategic_n} Strategic delayed",
         C["red"],  {"kind":"strat_delayed"},              "Delayed Strategic Projects"),
    ]
    ia_cols = st.columns(4)
    for col,(val,lbl,sub,color,spec,dlbl) in zip(ia_cols, IA):
        col.markdown(f"""
        <div style='background:{C['white']};border-radius:8px;padding:12px 14px;
          border:1px solid #E2E8F2;border-top:3px solid {color};margin:0 2px;'>
          <div style='font-size:24px;font-weight:700;color:{color};line-height:1;'>{val}</div>
          <div style='font-size:11px;font-weight:600;color:{C['navy']};margin-top:5px;'>{lbl}</div>
          <div style='font-size:10px;color:{C['gray']};margin-top:2px;'>{sub}</div>
        </div>""", unsafe_allow_html=True)
        if col.button("See projects →", key=f"ia_{lbl[:18]}", use_container_width=True):
            set_drill(spec, dlbl)
            st.rerun()

    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # Charts with drill selectors
    st.markdown("<div class='sec'>Executive Overview</div>", unsafe_allow_html=True)
    ch1, ch2, ch3 = st.columns(3)

    with ch1:
        st.caption("**Projects by Status**")
        if status_c:
            sc = revops_df[status_c].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            fig = px.bar(sc, x="Status", y="Count", color="Status",
                         color_discrete_map=STATUS_COLOR, template="plotly_white")
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_base(fig,210), use_container_width=True)
            opts = ["— select status —"] + sc["Status"].tolist()
            sel  = st.selectbox("", opts, key="exec_status_sel",
                                label_visibility="collapsed")
            if sel != "— select status —":
                set_drill({"kind":"status","value":sel}, f"Status: {sel}")
                st.rerun()

    with ch2:
        st.caption("**Projects by Priority Band**")
        if band_c:
            bc = revops_df[band_c].value_counts().reset_index()
            bc.columns = ["Band","Count"]
            fig2 = go.Figure(go.Bar(
                x=bc["Count"], y=bc["Band"], orientation="h",
                marker_color=[BAND_COLOR.get(b,C["gray"]) for b in bc["Band"]]))
            st.plotly_chart(chart_base(fig2,210), use_container_width=True)
            opts2 = ["— select band —"] + bc["Band"].tolist()
            sel2  = st.selectbox("", opts2, key="exec_band_sel",
                                 label_visibility="collapsed")
            if sel2 != "— select band —":
                set_drill({"kind":"band","value":sel2}, f"Priority Band: {sel2}")
                st.rerun()

    with ch3:
        st.caption("**CDM Dependency**")
        cv = revops_df[CDM].value_counts().reset_index()
        cv.columns = ["CDM","Count"]
        fig3 = px.pie(cv, names="CDM", values="Count", color="CDM",
                      color_discrete_map={"Yes":C["amber"],"No":C["teal"],"Unknown":C["gray"]},
                      hole=0.55, template="plotly_white")
        fig3.update_traces(textposition="outside", textfont_size=10,
                           marker=dict(line=dict(color="white",width=2)))
        fig3.update_layout(height=210, margin=dict(t=6,b=6,l=6,r=6),
                           showlegend=True, legend=dict(font=dict(size=10)),
                           paper_bgcolor="white")
        st.plotly_chart(fig3, use_container_width=True)
        opts3 = ["— select CDM —","Yes","No","Unknown"]
        sel3  = st.selectbox("", opts3, key="exec_cdm_sel",
                             label_visibility="collapsed")
        if sel3 != "— select CDM —":
            set_drill({"kind":"cdm","value":sel3}, f"CDM Dependency: {sel3}")
            st.rerun()

    # CDM callout
    st.markdown(f"""
    <div style='background:#FFF8F0;border:1px solid {C['amber']};
      border-left:4px solid {C['amber']};border-radius:8px;padding:12px 16px;margin:6px 0;'>
      <div style='font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;
        color:{C['amber']};margin-bottom:5px;'>CDM AS A CASE STUDY IN DELAY RISK</div>
      <div style='font-size:12px;color:#1a1a1a;line-height:1.55;'>
        <strong>P11 (CDM DAUT ID Replacement) is currently Delayed.</strong>
        Upstream dependency for 8 projects: P27, P20 (CPQ), P29 (DataCloud), P7 (Forecasting),
        P15 (MTM Expansion), P22 (Sales Planning), P38 (SYSS/MAS), P60 (AFAG).<br>
        <span style='color:{C['red']};font-weight:600;'>Every month CDM slips, it cascades
        across quoting, forecasting, pricing, and customer data integrity.</span>
      </div>
    </div>""", unsafe_allow_html=True)
    if st.button("See all CDM-dependent projects →", key="cdm_cta"):
        set_drill({"kind":"cdm","value":"Yes"}, "CDM-Dependent Projects")
        st.rerun()

    # Drill panel (appears here when active, else show spotlight)
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)
    if st.session_state["drill_spec"] is not None:
        render_drill_panel()
    else:
        st.markdown("<div class='sec'>Project Spotlight — Top & Delayed</div>",
                    unsafe_allow_html=True)
        st.caption("Click any KPI card or chart selector above to drill into that segment.")
        spot = revops_df.copy()
        spot["__s"] = 0
        if delay_c:  spot["__s"] += (spot[delay_c].str.upper()=="Y").astype(int)*10
        if impact_c: spot["__s"] += pd.to_numeric(spot[impact_c],errors="coerce").fillna(0)
        spot = spot.sort_values("__s",ascending=False).head(20)
        dcols = [c for c in [pid_c,name_c,band_c,status_c,type_c,CDM,bv_c,dar_c] if c]
        disp  = spot[dcols].rename(columns={CDM:"CDM Dep"})
        if bv_c  and bv_c  in disp.columns: disp=disp.rename(columns={bv_c:"Biz Value ($)"})
        if dar_c and dar_c in disp.columns: disp=disp.rename(columns={dar_c:"$ at Risk"})
        st.dataframe(disp.reset_index(drop=True),
                     use_container_width=True, hide_index=True, height=300)

    st.markdown(f"""
    <div class='edit-cta'>✏️ <strong>To update any field</strong> — edit directly in Excel.
    Dashboard auto-refreshes every 12 hours.
    &nbsp;<a href='{EDIT_LINK}' target='_blank'
      style='color:{C['blue']};font-weight:700;'>Open Excel →</a></div>""",
    unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# ██  PORTFOLIO DETAIL
# ═══════════════════════════════════════════════════════════════
elif view == "📊 Portfolio Detail":

    # Filters
    st.markdown("<div class='sec'>Filters</div>", unsafe_allow_html=True)
    f1,f2,f3,f4,f5 = st.columns(5)

    b_opts = ["All"] + (sorted(b for b in revops_df[band_c].dropna().unique() if b)
                        if band_c else [])
    s_opts = ["All"] + (sorted(s for s in revops_df[status_c].dropna().unique() if s)
                        if status_c else [])
    t_opts = ["All"] + (sorted(t for t in revops_df[type_c].dropna().unique() if t)
                        if type_c else [])
    c_opts = ["All"] + (sorted(c for c in revops_df[core_c].dropna().unique() if c)
                        if core_c else [])

    f_band   = f1.selectbox("Priority Band",  b_opts,                  key="pd_band")
    f_status = f2.selectbox("Status",         s_opts,                  key="pd_status")
    f_type   = f3.selectbox("Project Type",   t_opts,                  key="pd_type")
    f_cdm    = f4.selectbox("CDM Dependency", ["All","Yes","No","Unknown"], key="pd_cdm")
    f_core   = f5.selectbox("Requested By",   c_opts,                  key="pd_core")

    filt = revops_df.copy()
    if f_band   != "All" and band_c:   filt = filt[filt[band_c]==f_band]
    if f_status != "All" and status_c: filt = filt[filt[status_c]==f_status]
    if f_type   != "All" and type_c:   filt = filt[filt[type_c]==f_type]
    if f_cdm    != "All":              filt = filt[filt[CDM]==f_cdm]
    if f_core   != "All" and core_c:   filt = filt[filt[core_c]==f_core]

    st.caption(f"**{len(filt)}** of {total} RevOps projects")
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)

    # Charts
    ch1, ch2 = st.columns(2)
    with ch1:
        st.caption("**Status breakdown**")
        if status_c and not filt.empty:
            sc2 = filt[status_c].value_counts().reset_index()
            sc2.columns = ["Status","Count"]
            fig = px.bar(sc2.sort_values("Count",ascending=False),
                         x="Status", y="Count", color="Status",
                         color_discrete_map=STATUS_COLOR, template="plotly_white")
            fig.update_layout(showlegend=False)
            st.plotly_chart(chart_base(fig,200), use_container_width=True)
            ss = st.selectbox("Drill into status →",
                              ["— select —"]+sc2["Status"].tolist(), key="pd_s_drill")
            if ss != "— select —":
                set_drill({"kind":"status","value":ss}, f"Status: {ss}")
                st.rerun()

    with ch2:
        st.caption("**Priority Band**")
        if band_c and not filt.empty:
            bc2 = filt[band_c].value_counts().reset_index()
            bc2.columns = ["Band","Count"]
            fig2 = go.Figure(go.Bar(
                x=bc2["Count"], y=bc2["Band"], orientation="h",
                marker_color=[BAND_COLOR.get(b,C["gray"]) for b in bc2["Band"]]))
            st.plotly_chart(chart_base(fig2,200), use_container_width=True)
            bs = st.selectbox("Drill into band →",
                              ["— select —"]+bc2["Band"].tolist(), key="pd_b_drill")
            if bs != "— select —":
                set_drill({"kind":"band","value":bs}, f"Priority Band: {bs}")
                st.rerun()

    # Effort vs Impact
    if effort_c and impact_c and not filt.empty:
        st.caption("**Effort vs Impact**")
        sdf = filt[[c for c in [pid_c,name_c,effort_c,impact_c,status_c] if c]].copy()
        sdf[effort_c] = pd.to_numeric(sdf[effort_c], errors="coerce")
        sdf[impact_c] = pd.to_numeric(sdf[impact_c], errors="coerce")
        sdf = sdf.dropna(subset=[effort_c,impact_c])
        if not sdf.empty:
            def jit(s,sc=0.12):
                return s.apply(lambda v: ((int(hashlib.md5(str(v).encode()).hexdigest(),16)
                                           %1000)/1000-.5)*2*sc)
            sdf["__x"] = sdf[effort_c]+jit(sdf[effort_c])
            sdf["__y"] = sdf[impact_c]+jit(sdf[impact_c])
            fig3 = px.scatter(sdf, x="__x", y="__y",
                              color=status_c, color_discrete_map=STATUS_COLOR,
                              hover_data={c:True for c in [name_c,pid_c] if c},
                              template="plotly_white", opacity=0.8)
            fig3.update_traces(marker=dict(size=11,line=dict(width=1.5,color="white")))
            fig3.update_layout(xaxis_title="Effort",yaxis_title="Impact",showlegend=True,
                               legend=dict(orientation="h",y=1.1,font=dict(size=10)))
            st.plotly_chart(chart_base(fig3,230), use_container_width=True)

    # Resource load
    if res_df is not None and res_pid_c and res_team_c and pid_c:
        st.caption("**Resource / Team Load**")
        fpids = set(filt[pid_c].dropna().astype(str).unique())
        rf = res_df[res_df[res_pid_c].astype(str).isin(fpids)]
        if not rf.empty:
            tc = rf.groupby(res_team_c)[res_pid_c].nunique().reset_index()
            tc.columns = ["Team","Projects"]
            tc = tc.sort_values("Projects",ascending=True)
            avg = tc["Projects"].mean()
            tc["Color"] = tc["Projects"].apply(
                lambda x: C["red"] if x>avg*1.4 else (C["amber"] if x>avg else C["teal"]))
            fig4 = go.Figure(go.Bar(
                x=tc["Projects"], y=tc["Team"], orientation="h",
                marker_color=tc["Color"].tolist()))
            st.plotly_chart(chart_base(fig4,max(200,len(tc)*26)),use_container_width=True)
            st.caption("🔴 Overloaded  🟠 Elevated  🟢 Normal")
            ts = st.selectbox("Drill into team →",
                              ["— select —"]+sorted(tc["Team"].tolist()), key="pd_t_drill")
            if ts != "— select —":
                set_drill({"kind":"team","value":ts}, f"Team: {ts}")
                st.rerun()

    # Value Groups
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)
    st.markdown("<div class='sec'>Value Groups</div>", unsafe_allow_html=True)
    if vm_df is not None:
        vm_pid = nc(vm_df,"Project ID","ProjectID","ID")
        vm_grp = nc(vm_df,"Value Group","Group")
        if vm_pid and vm_grp and pid_c:
            fpids2 = set(filt[pid_c].dropna().astype(str).unique())
            vmf = vm_df[vm_df[vm_pid].isin(fpids2)]
            gc = vmf.groupby(vm_grp)[vm_pid].nunique().reset_index()
            gc.columns = ["Group","N"]
            grp_colors = [C["teal"],C["blue"],C["lblue"],C["navy"],C["green"],C["amber"],C["gray"]]
            vcols = st.columns(max(1,len(gc)))
            for i,(col,(_,row)) in enumerate(zip(vcols, gc.iterrows())):
                clr = grp_colors[i%len(grp_colors)]
                col.markdown(f"""
                <div style='background:{C['white']};border-radius:8px;padding:10px 12px;
                  border:1px solid #E2E8F2;border-top:3px solid {clr};text-align:center;'>
                  <div style='font-size:20px;font-weight:700;color:{clr};'>{int(row['N'])}</div>
                  <div style='font-size:10px;font-weight:600;color:{C['navy']};margin-top:2px;'>
                    {row['Group']}</div>
                </div>""", unsafe_allow_html=True)
                if col.button("View →", key=f"vg_{row['Group']}", use_container_width=True):
                    set_drill({"kind":"value_group","value":row["Group"]},
                              f"Value Group: {row['Group']}")
                    st.rerun()

    # Drill panel
    if st.session_state["drill_spec"] is not None:
        st.markdown("<hr class='slim'>", unsafe_allow_html=True)
        render_drill_panel()

    # Full table
    st.markdown("<hr class='slim'>", unsafe_allow_html=True)
    st.markdown("<div class='sec'>All Projects</div>", unsafe_allow_html=True)
    sc = [c for c in [pid_c,name_c,band_c,type_c,status_c,core_c,CDM,
                      effort_c,impact_c,bv_c,dar_c] if c]
    disp3 = filt[sc].rename(columns={CDM:"CDM Dep"})
    if bv_c  and bv_c  in disp3.columns: disp3=disp3.rename(columns={bv_c:"Biz Value ($)"})
    if dar_c and dar_c in disp3.columns: disp3=disp3.rename(columns={dar_c:"$ at Risk"})
    st.dataframe(disp3.reset_index(drop=True),use_container_width=True,
                 hide_index=True, height=380)

    st.markdown(f"""
    <div class='edit-cta'>✏️ <strong>To update any field</strong> — edit in Excel.
    &nbsp;<a href='{EDIT_LINK}' target='_blank'
      style='color:{C['blue']};font-weight:700;'>Open Excel →</a></div>""",
    unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# ██  PROJECT EXPLORER
# ═══════════════════════════════════════════════════════════════
else:
    st.markdown("<div class='sec'>Project Explorer</div>", unsafe_allow_html=True)

    s1,s2,s3 = st.columns([3,2,2])
    srch  = s1.text_input("Search name or ID", placeholder="Type to filter…", key="srch")
    bflt  = s2.selectbox("Band",   ["All"]+(sorted(b for b in revops_df[band_c].dropna().unique() if b) if band_c else []))
    sflt  = s3.selectbox("Status", ["All"]+(sorted(s for s in revops_df[status_c].dropna().unique() if s) if status_c else []))

    exp = revops_df.copy()
    if srch and pid_c and name_c:
        q = srch.lower()
        exp = exp[exp[pid_c].astype(str).str.lower().str.contains(q,na=False) |
                  exp[name_c].astype(str).str.lower().str.contains(q,na=False)]
    if bflt != "All" and band_c:   exp = exp[exp[band_c]==bflt]
    if sflt != "All" and status_c: exp = exp[exp[status_c]==sflt]

    if band_c:
        bo = {"Top Priority":0,"Middle Priority":1,"Lower Priority":2,"N/A":3}
        exp = exp.copy()
        exp["__bs"] = exp[band_c].map(bo).fillna(4)
        exp["__ds"] = (exp[delay_c].str.upper()=="Y").astype(int)*-1 if delay_c else 0
        exp = exp.sort_values(["__bs","__ds"])

    st.caption(f"**{len(exp)}** projects")
    lc, rc2 = st.columns([2,3])

    with lc:
        if exp.empty:
            st.info("No projects match.")
        else:
            for _,row in exp.iterrows():
                pid  = str(row[pid_c]) if pid_c else "—"
                pn   = str(row[name_c]) if name_c else "—"
                band = str(row[band_c]) if band_c else ""
                stat = str(row[status_c]) if status_c else ""
                cdm  = row[CDM]
                bc   = BAND_COLOR.get(band,C["gray"])
                if st.button(f"[{pid}]  {pn[:40]}{'…' if len(pn)>40 else ''}",
                             key=f"xp_{pid}", use_container_width=True):
                    st.session_state["selected_pid"] = pid
                st.markdown(
                    f"<div style='font-size:10px;color:{C['gray']};margin:-5px 0 5px 6px;'>"
                    f"<span style='color:{bc};font-weight:600;'>{band.replace(' Priority','')}</span>"
                    f"  ·  {sbadge(stat,10)}  ·  "
                    f"{'<span style=\"color:#D97706\">⚠CDM</span>' if cdm=='Yes' else '—'}"
                    f"</div>", unsafe_allow_html=True)

    with rc2:
        sel = st.session_state.get("selected_pid")
        if not sel and not exp.empty and pid_c:
            sel = str(exp.iloc[0][pid_c])
            st.session_state["selected_pid"] = sel

        if sel and pid_c:
            m = revops_df[revops_df[pid_c].astype(str)==sel]
            if not m.empty:
                row   = m.iloc[0]
                pname = str(row[name_c]) if name_c else sel
                pstat = str(row[status_c]) if status_c else "—"
                pband = str(row[band_c])   if band_c  else "—"
                ptype = str(row[type_c])   if type_c  else "—"
                is_del= str(row[delay_c]).strip().upper()=="Y" if delay_c else False
                pcdm  = row[CDM]
                acc   = C["red"] if is_del else (C["amber"] if pcdm=="Yes" else C["blue"])

                st.markdown(f"""
                <div style='background:linear-gradient(135deg,{C['navy']} 0%,{acc} 100%);
                  border-radius:10px;padding:16px 20px;color:white;margin-bottom:12px;'>
                  <div style='font-family:DM Mono;font-size:10px;
                    color:rgba(255,255,255,.55);'>{sel}</div>
                  <div style='font-size:16px;font-weight:700;margin:4px 0;'>{pname}</div>
                  <div style='display:flex;gap:8px;flex-wrap:wrap;margin-top:5px;'>
                    {sbadge(pstat,11)}
                    <span class='badge' style='background:rgba(255,255,255,.15);
                      color:white;font-size:10px;'>{pband}</span>
                    <span class='badge' style='background:rgba(255,255,255,.15);
                      color:white;font-size:10px;'>{ptype}</span>
                    {'<span class="badge" style="background:#FEF3C7;color:#D97706;font-size:10px;">⚠ CDM</span>' if pcdm=="Yes" else ""}
                    {'<span class="badge" style="background:#FEE2E2;color:#C0392B;font-size:10px;">⚠ Delayed</span>' if is_del else ""}
                  </div>
                </div>""", unsafe_allow_html=True)

                t1,t2,t3,t4 = st.tabs(["Overview","Resources","Dependencies","Risk & Value"])

                def dv(lbl,col_k,fb="Not yet captured"):
                    v = (str(row[col_k]) if col_k and col_k in row.index
                         and pd.notna(row[col_k])
                         and str(row[col_k]) not in ("None","nan","") else None)
                    ph = " ph" if not v else ""
                    return (f"<div class='detail-lbl'>{lbl}</div>"
                            f"<div class='detail-val{ph}'>{v or fb}</div>")

                with t1:
                    c1,c2 = st.columns(2)
                    with c1:
                        st.markdown(dv("Business Program",bizprog_c)+dv("Core Team",core_c)+
                                    dv("Strategic Priority",strat_c)+dv("Cycle",cycle_c)+
                                    dv("Investment",invest_c), unsafe_allow_html=True)
                    with c2:
                        st.markdown(dv("Priority Rank",rank_c)+dv("Effort",effort_c)+
                                    dv("Impact",impact_c)+dv("If Delayed Impact",deli_c)+
                                    dv("Value Groups",valgrp_c), unsafe_allow_html=True)
                    rw = (str(row[rawval_c]) if rawval_c and pd.notna(row[rawval_c])
                          and str(row[rawval_c]) not in ("None","nan","") else None)
                    if rw:
                        st.markdown(f"""
                        <div style='background:#F8FAFC;border-radius:6px;padding:10px 14px;
                          border-left:3px solid {C['blue']};margin-top:4px;'>
                          <div class='detail-lbl'>What This Project Delivers</div>
                          <div style='font-size:12px;color:#374151;margin-top:3px;
                            line-height:1.5;'>{rw}</div>
                        </div>""", unsafe_allow_html=True)

                with t2:
                    teams = proj_teams.get(sel,[])
                    if teams:
                        st.markdown(f"**{len(teams)} teams:**")
                        for t in sorted(teams): st.markdown(f"- {t}")
                    else: st.info("No resource data.")

                with t3:
                    deps = proj_deps.get(sel,[])
                    if deps:
                        st.markdown(f"**Depends on {len(deps)}:**")
                        for d in deps:
                            dm = revops_df[revops_df[pid_c].astype(str)==d.strip()] if pid_c else pd.DataFrame()
                            dn = str(dm.iloc[0][name_c]) if not dm.empty and name_c else d
                            ds = str(dm.iloc[0][status_c]) if not dm.empty and status_c else "—"
                            st.markdown(f"- **{d}** — {dn} &nbsp; {sbadge(ds)}", unsafe_allow_html=True)
                    else: st.info("No dependencies.")
                    blocking = [p for p,dl in proj_deps.items() if sel in dl]
                    if blocking:
                        st.markdown(f"**{len(blocking)} depend on this:**")
                        for b in blocking:
                            bm = revops_df[revops_df[pid_c].astype(str)==b] if pid_c else pd.DataFrame()
                            bn = str(bm.iloc[0][name_c]) if not bm.empty and name_c else b
                            st.markdown(f"- **{b}** — {bn}")

                with t4:
                    r1,r2 = st.columns(2)
                    with r1:
                        for lbl2,ck,clr in [("Effort",effort_c,C["blue"]),
                                             ("Impact",impact_c,C["teal"])]:
                            v2 = pd.to_numeric(row[ck],errors="coerce") if ck else None
                            dsp = str(int(v2)) if v2 is not None and not pd.isna(v2) else "—"
                            cl2 = clr if v2 is not None and not pd.isna(v2) else C["gray"]
                            st.markdown(f"""
                            <div style='background:{C['white']};border-radius:8px;
                              padding:12px;border:1px solid #E2E8F2;margin-bottom:8px;'>
                              <div class='detail-lbl'>{lbl2} Score</div>
                              <div style='font-size:26px;font-weight:700;color:{cl2};'>{dsp}</div>
                            </div>""", unsafe_allow_html=True)
                    with r2:
                        for lbl3,ck2,clr2 in [("Business Value ($)",bv_c,C["teal"]),
                                               ("Dollars at Risk ($)",dar_c,C["red"])]:
                            v3 = row[ck2] if ck2 else None
                            ph2= not (v3 and pd.notna(v3) and str(v3) not in ("None","nan",""))
                            dsp2 = fmt_money(v3) if not ph2 else "Not yet captured"
                            cl3  = clr2 if not ph2 else C["gray"]
                            st.markdown(f"""
                            <div style='background:{C['white']};border-radius:8px;
                              padding:12px;border:1px solid #E2E8F2;margin-bottom:8px;'>
                              <div class='detail-lbl'>{lbl3}</div>
                              <div style='font-size:{"18px" if not ph2 else "12px"};
                                font-weight:700;color:{cl3};
                                {"font-style:italic;" if ph2 else ""}'>{dsp2}</div>
                            </div>""", unsafe_allow_html=True)

                    flag_bg  = "#FEF3F2" if is_del else ("#FFF8F0" if pcdm=="Yes" else "#F0FFF8")
                    flag_brd = C["red"]  if is_del else (C["amber"] if pcdm=="Yes" else C["teal"])
                    flag_ttl = ("⚠ Delay Flag Active" if is_del
                                else ("CDM Dependency" if pcdm=="Yes"
                                      else "✓ No Active Risk Flags"))
                    flag_txt = ("Review owner accountability and blockers."
                                if is_del else
                                ("On track now but blocked pending P11." if pcdm=="Yes"
                                 else ""))
                    st.markdown(f"""
                    <div style='background:{flag_bg};border-left:3px solid {flag_brd};
                      border-radius:0 6px 6px 0;padding:10px 14px;font-size:12px;'>
                      <strong style='color:{flag_brd};'>{flag_ttl}</strong>
                      {"<br><span style='color:#374151;'>"+flag_txt+"</span>" if flag_txt else ""}
                    </div>""", unsafe_allow_html=True)

                st.markdown(f"""
                <div class='edit-cta'>To update <strong>{pname}</strong> — edit in Excel.
                &nbsp;<a href='{EDIT_LINK}' target='_blank'
                  style='color:{C['blue']};font-weight:700;'>Open Excel →</a></div>""",
                unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────
st.markdown(f"""
<hr class='slim'>
<div style='display:flex;justify-content:space-between;font-size:10px;
  color:{C['gray']};padding-bottom:8px;'>
  <span>RevOps Program Dashboard · FY26 · Emerson · Read-only</span>
  <span>Auto-refreshes every 12 hours ·
    <a href='{EDIT_LINK}' target='_blank'
      style='color:{C['blue']};text-decoration:none;font-weight:600;'>
      Edit source in Excel →</a></span>
</div>""", unsafe_allow_html=True)
