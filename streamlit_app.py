import re
from io import BytesIO

import pandas as pd
import plotly.express as px
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
PALETTE = [
    C["deep_blue"], C["bright_blue"], C["teal"], C["soft_green"],
    C["navy"], C["light_blue"], C["green"], C["gray"],
]
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

st.set_page_config(
    page_title="RevOps Program Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
    .main {{ background-color: #F7F8FA; }}
    .block-container {{ padding: 1.8rem 2.2rem 2rem 2.2rem; max-width: 1400px; }}
    .kpi-wrap {{
        background: {C["white"]};
        border-radius: 12px;
        padding: 20px 22px 16px 22px;
        border: 1px solid #E8ECF0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }}
    .kpi-label {{
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        color: {C["gray"]};
        margin-bottom: 4px;
    }}
    .kpi-value {{
        font-size: 34px;
        font-weight: 700;
        color: {C["navy"]};
        line-height: 1;
    }}
    .kpi-value.danger {{ color: #C0392B; }}
    .kpi-value.success {{ color: {C["teal"]}; }}
    .kpi-sub {{ font-size: 11px; color: {C["gray"]}; margin-top: 6px; }}
    .kpi-accent-bar {{ height: 3px; border-radius: 2px; margin-bottom: 10px; }}
    .section-title {{
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: {C["navy"]};
        margin-bottom: 16px;
        padding-bottom: 8px;
        border-bottom: 2px solid {C["deep_blue"]};
        display: inline-block;
    }}
    .section-divider {{
        border: none;
        border-top: 1px solid #E8ECF0;
        margin: 32px 0 28px 0;
    }}
    .exec-summary {{
        background: linear-gradient(135deg, {C["navy"]} 0%, {C["deep_blue"]} 100%);
        border-radius: 12px;
        padding: 22px 28px;
        color: white;
        margin-bottom: 28px;
    }}
    .exec-summary h1 {{
        font-size: 22px;
        font-weight: 700;
        color: white !important;
        margin: 0 0 4px 0;
        letter-spacing: -0.3px;
    }}
    .exec-summary .subtitle {{
        font-size: 13px;
        color: rgba(255,255,255,0.65);
        margin-bottom: 14px;
        font-weight: 400;
    }}
    .exec-summary .dynamic {{
        font-size: 14px;
        color: rgba(255,255,255,0.9);
        font-weight: 400;
        border-top: 1px solid rgba(255,255,255,0.15);
        padding-top: 12px;
        margin-top: 4px;
    }}
    .risk-item {{
        background: #FEF3F2;
        border-left: 3px solid #C0392B;
        border-radius: 0 6px 6px 0;
        padding: 10px 14px;
        margin-bottom: 8px;
        font-size: 13px;
        color: #1a1a1a;
    }}
    .risk-item.warn {{ background: #FFF8F0; border-left-color: #E67E22; }}
    .risk-item.info {{ background: #F0F7FF; border-left-color: {C["deep_blue"]}; }}
    .detail-card {{
        background: {C["white"]};
        border-radius: 10px;
        padding: 18px 20px;
        border: 1px solid #E8ECF0;
        margin-bottom: 12px;
    }}
    .detail-label {{
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        color: {C["gray"]};
        margin-bottom: 3px;
    }}
    .detail-value {{ font-size: 14px; font-weight: 500; color: {C["navy"]}; }}
    .status-badge {{
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.04em;
    }}
    .empty-state {{
        text-align: center;
        padding: 48px 24px;
        color: {C["gray"]};
        font-size: 14px;
    }}
    [data-testid="stSidebar"] {{
        background: {C["white"]};
        border-right: 1px solid #E8ECF0;
    }}
    #MainMenu {{ visibility: hidden; }}
    footer {{ visibility: hidden; }}
    header {{ visibility: hidden; }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        background: #F0F2F5;
        border-radius: 8px;
        padding: 4px;
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 13px;
        font-weight: 500;
    }}
    .stTabs [aria-selected="true"] {{
        background: {C["white"]};
        color: {C["navy"]};
    }}
</style>
""", unsafe_allow_html=True)


# ── UTILITIES ─────────────────────────────────
def normalize_cols(df):
    df.columns = [c.strip() for c in df.columns]
    return df


def get_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def build_download_url(url):
    if "/:x:/p/" in url or "/:x:/s/" in url:
        sep = "&" if "?" in url else "?"
        return url + sep + "download=1"
    if "_layouts/15/Doc.aspx" in url:
        match = re.search(r'sourcedoc=%7B([^%]+)%7D', url, re.IGNORECASE)
        if match:
            guid = match.group(1)
            base = url.split("/_layouts/")[0]
            return f"{base}/_layouts/15/download.aspx?UniqueId={guid}"
    if "1drv.ms" in url:
        try:
            r = requests.get(url, allow_redirects=True, timeout=15)
            resolved = r.url
            sep = "&" if "?" in resolved else "?"
            return resolved + sep + "download=1"
        except Exception:
            pass
    sep = "&" if "?" in url else "?"
    return url + sep + "download=1"


def chart_layout(fig, height=300, legend=False):
    fig.update_layout(
        height=height,
        margin=dict(t=16, b=16, l=8, r=8),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(family="Inter, sans-serif", size=12, color="#374151"),
        showlegend=legend,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
            font=dict(size=11),
        ) if legend else {},
        xaxis=dict(gridcolor="#F0F2F5", linecolor="#E8ECF0", tickfont=dict(size=11)),
        yaxis=dict(gridcolor="#F0F2F5", linecolor="#E8ECF0", tickfont=dict(size=11)),
    )
    fig.update_traces(marker_line_width=0)
    return fig


def status_badge_html(status):
    color_map = {
        "delayed":     ("#FEE2E2", "#C0392B"),
        "at risk":     ("#FEF3C7", "#D97706"),
        "on track":    ("#D1FAE5", "#065F46"),
        "active":      ("#DBEAFE", "#1E40AF"),
        "in progress": ("#E0F2FE", "#0369A1"),
        "complete":    ("#D1FAE5", "#065F46"),
        "completed":   ("#D1FAE5", "#065F46"),
        "not started": ("#F3F4F6", "#374151"),
        "planning":    ("#EDE9FE", "#5B21B6"),
    }
    key = str(status).lower() if status else ""
    bg, fg = color_map.get(key, ("#F3F4F6", "#374151"))
    return f"<span class='status-badge' style='background:{bg};color:{fg};'>{status}</span>"


# ── DATA LOADING ──────────────────────────────
@st.cache_data(ttl=60)
def load_data(url):
    try:
        download_url = build_download_url(url)
        headers = {"User-Agent": "Mozilla/5.0"}
        r = requests.get(download_url, headers=headers, timeout=30, allow_redirects=True)
        r.raise_for_status()
        content_type = r.headers.get("Content-Type", "")
        if "html" in content_type.lower():
            fallback = url + ("&download=1" if "?" in url else "?download=1")
            r = requests.get(fallback, headers=headers, timeout=30, allow_redirects=True)
            r.raise_for_status()
        content = BytesIO(r.content)
        sheets = {}
        for sheet in ["Projects", "Project_Resources", "Dependencies"]:
            try:
                df = pd.read_excel(content, sheet_name=sheet, engine="openpyxl")
                sheets[sheet] = normalize_cols(df)
            except Exception:
                sheets[sheet] = None
        return sheets, None
    except Exception as e:
        return None, str(e)


# ── LOAD DATA ─────────────────────────────────
with st.spinner(""):
    sheets, err = load_data(ONEDRIVE_FILE_URL)

if err:
    st.error(f"**Data load failed:** {err}")
    st.info("Ensure the SharePoint link is set to 'Anyone with the link can view'.")
    st.stop()

proj_df = sheets.get("Projects")
res_df  = sheets.get("Project_Resources")
dep_df  = sheets.get("Dependencies")

for m in [s for s, d in sheets.items() if d is None]:
    st.warning(f"Sheet '{m}' could not be loaded.")

if proj_df is None:
    st.error("Projects sheet is required. Cannot render dashboard.")
    st.stop()

# ── COLUMN MAP ────────────────────────────────
owner_col          = get_col(proj_df, "Owner", "owner", "PM", "Project Owner")
team_col_p         = get_col(proj_df, "Team", "team", "Department")
status_col         = get_col(proj_df, "Status", "status", "Project Status")
cycle_col          = get_col(proj_df, "Cycle", "cycle", "Sprint", "Quarter")
priority_col       = get_col(proj_df, "Priority", "priority", "Priority Type")
effort_col         = get_col(proj_df, "Effort", "effort", "Effort Score")
impact_col         = get_col(proj_df, "Impact", "impact", "Impact Score")
proj_id_col        = get_col(proj_df, "Project ID", "ProjectID", "ID", "project_id")
proj_name_col      = get_col(proj_df, "Project", "Project Name", "project", "Name")
delayed_impact_col = get_col(proj_df, "If Delayed Impact", "Delayed Impact", "delay_impact", "Impact If Delayed")
notes_col          = get_col(proj_df, "Notes", "notes", "Risk Notes", "Delay Notes", "Comments")

# ── SIDEBAR ───────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style='margin-bottom:20px;'>
      <div style='font-size:15px;font-weight:700;color:{C["navy"]};'>RevOps Dashboard</div>
      <div style='font-size:11px;color:{C["gray"]};margin-top:2px;'>Filter controls</div>
    </div>
    """, unsafe_allow_html=True)

    filtered = proj_df.copy()
    sel_owners = []

    if owner_col:
        owners = sorted(proj_df[owner_col].dropna().unique().tolist())
        default_owners = ["RevOps"] if "RevOps" in owners else owners
        sel_owners = st.multiselect("Owner", owners, default=default_owners)
        if sel_owners:
            filtered = filtered[filtered[owner_col].isin(sel_owners)]

    if team_col_p:
        teams = sorted(proj_df[team_col_p].dropna().unique().tolist())
        sel_teams = st.multiselect("Team", teams, default=[])
        if sel_teams:
            filtered = filtered[filtered[team_col_p].isin(sel_teams)]

    if status_col:
        statuses = sorted(proj_df[status_col].dropna().unique().tolist())
        sel_status = st.multiselect("Status", statuses, default=[])
        if sel_status:
            filtered = filtered[filtered[status_col].isin(sel_status)]

    if cycle_col:
        cycles = sorted(proj_df[cycle_col].dropna().unique().tolist())
        sel_cycles = st.multiselect("Cycle", cycles, default=[])
        if sel_cycles:
            filtered = filtered[filtered[cycle_col].isin(sel_cycles)]

    if priority_col:
        priorities = sorted(proj_df[priority_col].dropna().unique().tolist())
        sel_priorities = st.multiselect("Priority Type", priorities, default=[])
        if sel_priorities:
            filtered = filtered[filtered[priority_col].isin(sel_priorities)]

    st.markdown("<hr style='border:none;border-top:1px solid #E8ECF0;margin:16px 0;'>", unsafe_allow_html=True)
    st.caption(f"**{len(filtered)}** of {len(proj_df)} projects shown")
    st.caption("Auto-refreshes every 60 seconds")

# ── DERIVED METRICS ───────────────────────────
total = len(filtered)

delayed_mask = pd.Series([False] * len(filtered), index=filtered.index)
if status_col:
    delayed_mask = filtered[status_col].str.lower().str.contains("delay", na=False)
delayed_count = int(delayed_mask.sum())

active_count = 0
if status_col:
    active_count = int(filtered[status_col].str.lower().str.contains(
        "active|in progress|in-progress", na=False, regex=True).sum())

teams_count   = 0
team_col_r    = None
pid_col_r     = None
if res_df is not None:
    team_col_r = get_col(res_df, "Team", "team", "Department", "Resource Team")
    pid_col_r  = get_col(res_df, "Project ID", "ProjectID", "ID")
    if proj_id_col and pid_col_r and proj_id_col in filtered.columns:
        active_pids = filtered[proj_id_col].dropna().unique()
        res_filtered_global = res_df[res_df[pid_col_r].isin(active_pids)]
        if team_col_r:
            teams_count = res_filtered_global[team_col_r].nunique()

avg_impact = None
avg_effort = None
if impact_col:
    vals = pd.to_numeric(filtered[impact_col], errors="coerce")
    if vals.notna().any():
        avg_impact = round(float(vals.mean()), 1)
if effort_col:
    vals = pd.to_numeric(filtered[effort_col], errors="coerce")
    if vals.notna().any():
        avg_effort = round(float(vals.mean()), 1)

# ── DYNAMIC SUMMARY ───────────────────────────
if owner_col and sel_owners:
    owner_label = f"{', '.join(sel_owners)} " if len(sel_owners) <= 2 else "filtered "
else:
    owner_label = "RevOps " if (owner_col and "RevOps" in proj_df[owner_col].values) else ""

summary_parts = [f"<strong>{total}</strong> {owner_label}projects tracked"]
if teams_count:
    summary_parts.append(f"across <strong>{teams_count}</strong> teams")
if delayed_count:
    summary_parts.append(
        f"with <strong>{delayed_count}</strong> delayed program{'s' if delayed_count != 1 else ''} requiring attention"
    )
else:
    summary_parts.append("with no delays flagged")
dynamic_summary = ", ".join(summary_parts[:2])
if len(summary_parts) > 2:
    dynamic_summary += f", {summary_parts[2]}"
dynamic_summary += "."

# ── HEADER ────────────────────────────────────
st.markdown(f"""
<div class="exec-summary">
  <h1>RevOps Program Dashboard</h1>
  <div class="subtitle">Resource load, project risk, and dependency visibility</div>
  <div class="dynamic">{dynamic_summary}</div>
</div>
""", unsafe_allow_html=True)


# ── KPI ROW ───────────────────────────────────
def kpi_card(col, label, value, sub, color_class="", accent_color=None):
    accent = accent_color or C["deep_blue"]
    col.markdown(f"""
    <div class="kpi-wrap">
      <div>
        <div class="kpi-accent-bar" style="background:{accent};width:32px;"></div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color_class}">{value}</div>
      </div>
      <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


k1, k2, k3, k4, k5, k6 = st.columns(6)
kpi_card(k1, "Total Projects",   total,         "in current filters",   accent_color=C["navy"])
kpi_card(k2, "Delayed Projects", delayed_count, "require attention",
         color_class="danger" if delayed_count else "",
         accent_color="#C0392B" if delayed_count else C["gray"])
kpi_card(k3, "Active Projects",  active_count,  "in progress",          accent_color=C["deep_blue"])
kpi_card(k4, "Teams Involved",   teams_count,   "across resource pool", accent_color=C["teal"])
kpi_card(k5, "Avg Impact",
         avg_impact if avg_impact is not None else "—",
         "mean impact score",
         color_class="success" if avg_impact else "",
         accent_color=C["soft_green"])
kpi_card(k6, "Avg Effort",
         avg_effort if avg_effort is not None else "—",
         "mean effort score",
         accent_color=C["bright_blue"])

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# SECTION 1 — EXECUTIVE OVERVIEW
# ═══════════════════════════════════════════════
st.markdown("<div class='section-title'>Executive Overview</div>", unsafe_allow_html=True)
st.markdown("")

if total == 0:
    st.markdown("<div class='empty-state'>No projects match the current filter selection.</div>",
                unsafe_allow_html=True)
else:
    ov_left, ov_mid, ov_right = st.columns([2, 2, 1.4])

    with ov_left:
        st.markdown(
            f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
            f"text-transform:uppercase;margin-bottom:10px;'>Projects by Team</div>",
            unsafe_allow_html=True,
        )
        if res_df is not None and team_col_r and pid_col_r:
            if proj_id_col and proj_id_col in filtered.columns:
                active_pids = filtered[proj_id_col].dropna().unique()
                chart_res = res_df[res_df[pid_col_r].isin(active_pids)]
            else:
                chart_res = res_df.copy()
            if not chart_res.empty:
                tc = chart_res.groupby(team_col_r)[pid_col_r].nunique().reset_index()
                tc.columns = ["Team", "Projects"]
                tc = tc.sort_values("Projects", ascending=True).tail(10)
                fig = px.bar(
                    tc, x="Projects", y="Team", orientation="h",
                    color="Projects",
                    color_continuous_scale=[[0, C["light_blue"]], [1, C["deep_blue"]]],
                    template="plotly_white",
                )
                fig = chart_layout(fig, height=280)
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.markdown("<div class='empty-state'>No resource data for filtered projects.</div>",
                            unsafe_allow_html=True)
        else:
            st.warning("Resource data unavailable.")

    with ov_mid:
        st.markdown(
            f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
            f"text-transform:uppercase;margin-bottom:10px;'>Projects by Status</div>",
            unsafe_allow_html=True,
        )
        if status_col:
            sc = filtered[status_col].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            sc = sc.sort_values("Count", ascending=False)
            fig2 = px.bar(
                sc, x="Status", y="Count",
                color="Status", color_discrete_map=STATUS_COLORS,
                template="plotly_white",
            )
            fig2 = chart_layout(fig2, height=280)
            fig2.update_layout(showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("Status column not found.")

    with ov_right:
        st.markdown(
            f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
            f"text-transform:uppercase;margin-bottom:10px;'>Top Risks</div>",
            unsafe_allow_html=True,
        )
        risks_shown = 0

        if delayed_count > 0:
            st.markdown(
                f"<div class='risk-item'>⚠️ <strong>{delayed_count}</strong> "
                f"delayed project{'s' if delayed_count != 1 else ''} flagged</div>",
                unsafe_allow_html=True,
            )
            risks_shown += 1

        if res_df is not None and team_col_r and pid_col_r and proj_id_col and proj_id_col in filtered.columns:
            active_pids = filtered[proj_id_col].dropna().unique()
            top_team_df = res_df[res_df[pid_col_r].isin(active_pids)]
            if not top_team_df.empty:
                grouped = top_team_df.groupby(team_col_r)[pid_col_r].nunique()
                top_team = grouped.idxmax()
                top_team_count = grouped.max()
                st.markdown(
                    f"<div class='risk-item warn'>📌 <strong>{top_team}</strong> carries "
                    f"{top_team_count} projects — highest load</div>",
                    unsafe_allow_html=True,
                )
                risks_shown += 1

        if status_col and impact_col:
            hi_delayed = filtered[
                delayed_mask & (pd.to_numeric(filtered[impact_col], errors="coerce") >= 4)
            ]
            if not hi_delayed.empty:
                n = len(hi_delayed)
                st.markdown(
                    f"<div class='risk-item'>🔴 <strong>{n}</strong> delayed project"
                    f"{'s' if n != 1 else ''} with high impact score</div>",
                    unsafe_allow_html=True,
                )
                risks_shown += 1

        if priority_col:
            high_pri = filtered[
                filtered[priority_col].astype(str).str.lower().str.contains(
                    "high|critical|p1", na=False
                )
            ]
            if not high_pri.empty:
                st.markdown(
                    f"<div class='risk-item info'>🔵 <strong>{len(high_pri)}</strong> "
                    f"high-priority project{'s' if len(high_pri) != 1 else ''} in portfolio</div>",
                    unsafe_allow_html=True,
                )
                risks_shown += 1

        if risks_shown == 0:
            st.markdown(
                f"<div style='color:{C['teal']};font-size:13px;padding:12px 0;'>"
                f"✅ No critical risks identified with current filters.</div>",
                unsafe_allow_html=True,
            )

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# SECTION 2 — PORTFOLIO ANALYSIS
# ═══════════════════════════════════════════════
st.markdown("<div class='section-title'>Portfolio Analysis</div>", unsafe_allow_html=True)
st.markdown("")

if total == 0:
    st.markdown("<div class='empty-state'>No projects to analyze.</div>", unsafe_allow_html=True)
else:
    pa_left, pa_right = st.columns([3, 2])

    with pa_left:
        st.markdown(
            f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
            f"text-transform:uppercase;margin-bottom:10px;'>Impact vs. Effort</div>",
            unsafe_allow_html=True,
        )
        if effort_col and impact_col:
            keep = [c for c in [effort_col, impact_col, status_col, proj_name_col, owner_col] if c]
            scatter_df = filtered[keep].copy()
            scatter_df[effort_col] = pd.to_numeric(scatter_df[effort_col], errors="coerce")
            scatter_df[impact_col] = pd.to_numeric(scatter_df[impact_col], errors="coerce")
            scatter_df = scatter_df.dropna(subset=[effort_col, impact_col])
            if not scatter_df.empty:
                hover = {}
                if proj_name_col:
                    hover[proj_name_col] = True
                if owner_col:
                    hover[owner_col] = True
                fig3 = px.scatter(
                    scatter_df, x=effort_col, y=impact_col,
                    color=status_col if status_col else None,
                    color_discrete_map=STATUS_COLORS,
                    hover_data=hover,
                    template="plotly_white",
                    opacity=0.85,
                )
                fig3.update_traces(marker=dict(size=12, line=dict(width=1.5, color="white")))
                fig3 = chart_layout(fig3, height=320, legend=True)
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.markdown(
                    "<div class='empty-state'>Insufficient numeric data for scatter plot.</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.warning(f"Effort or Impact columns not found. Detected: {list(proj_df.columns)}")

    with pa_right:
        if priority_col:
            st.markdown(
                f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
                f"text-transform:uppercase;margin-bottom:10px;'>Priority Distribution</div>",
                unsafe_allow_html=True,
            )
            pri_counts = filtered[priority_col].value_counts().reset_index()
            pri_counts.columns = ["Priority", "Count"]
            fig4 = px.pie(
                pri_counts, names="Priority", values="Count",
                color_discrete_sequence=PALETTE,
                hole=0.52,
                template="plotly_white",
            )
            fig4.update_traces(
                textposition="outside",
                textfont_size=11,
                marker=dict(line=dict(color="white", width=2)),
            )
            fig4.update_layout(
                height=200,
                margin=dict(t=8, b=8, l=8, r=8),
                showlegend=True,
                legend=dict(font=dict(size=11), orientation="v"),
                paper_bgcolor="white",
                font=dict(family="Inter, sans-serif"),
            )
            st.plotly_chart(fig4, use_container_width=True)

        st.markdown(
            f"<div style='font-size:12px;font-weight:600;color:{C['gray']};letter-spacing:0.05em;"
            f"text-transform:uppercase;margin-top:16px;margin-bottom:10px;'>Delayed Projects</div>",
            unsafe_allow_html=True,
        )
        if status_col:
            delayed_df = filtered[delayed_mask].copy()
            table_cols = [
                c for c in [proj_id_col, proj_name_col, owner_col, status_col,
                             cycle_col, impact_col, delayed_impact_col] if c
            ]
            if not delayed_df.empty and table_cols:
                st.dataframe(
                    delayed_df[table_cols].reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                    height=200,
                )
            elif delayed_df.empty:
                st.markdown(
                    f"<div style='color:{C['teal']};font-size:13px;padding:8px 0;'>"
                    f"✅ No delayed projects in view.</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.warning("Delayed projects found but columns are missing.")
        else:
            st.warning("Status column unavailable.")

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# SECTION 3 — PROJECT EXPLORER
# ═══════════════════════════════════════════════
st.markdown("<div class='section-title'>Project Explorer</div>", unsafe_allow_html=True)
st.markdown("")

if proj_name_col:
    proj_options = sorted(filtered[proj_name_col].dropna().unique().tolist())
elif proj_id_col:
    proj_options = sorted(filtered[proj_id_col].dropna().astype(str).unique().tolist())
else:
    proj_options = []

if not proj_options:
    st.markdown(
        "<div class='empty-state'>No projects available with current filters.</div>",
        unsafe_allow_html=True,
    )
else:
    selected_proj = st.selectbox(
        "Select a project to explore",
        proj_options,
        help="Search or select a project to view full details, resources, and dependencies.",
    )

    if proj_name_col:
        proj_row = filtered[filtered[proj_name_col] == selected_proj]
    else:
        proj_row = filtered[filtered[proj_id_col].astype(str) == selected_proj]

    if not proj_row.empty:
        row = proj_row.iloc[0]
        proj_status = row.get(status_col, "") if status_col else ""
        is_delayed  = "delay" in str(proj_status).lower()
        badge_html  = status_badge_html(proj_status) if proj_status else ""
        pid_display = (
            f"<span style='color:{C['gray']};font-size:13px;margin-left:10px;'>"
            f"{row.get(proj_id_col, '')}</span>"
            if proj_id_col else ""
        )

        st.markdown(f"""
        <div class="detail-card" style="border-left: 4px solid {'#C0392B' if is_delayed else C['deep_blue']};">
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <span style="font-size:17px;font-weight:700;color:{C['navy']};">{selected_proj}</span>
            {pid_display}
            {badge_html}
          </div>
          <div style="margin-top:6px;font-size:12px;color:#C0392B;">
            {'⚠️ This project is flagged as delayed.' if is_delayed else ''}
          </div>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4 = st.tabs(["Overview", "Resources", "Dependencies", "Risk & Impact"])

        with tab1:
            meta_cols = [
                c for c in [proj_id_col, owner_col, team_col_p, status_col,
                             cycle_col, priority_col, effort_col, impact_col] if c
            ]
            if meta_cols:
                pairs = [(col, row.get(col, "—")) for col in meta_cols]
                half  = len(pairs) // 2 + len(pairs) % 2
                m1, m2 = st.columns(2)
                for col_name, val in pairs[:half]:
                    m1.markdown(f"""
                    <div style="margin-bottom:14px;">
                      <div class="detail-label">{col_name}</div>
                      <div class="detail-value">{val if (val == val and val != '') else '—'}</div>
                    </div>""", unsafe_allow_html=True)
                for col_name, val in pairs[half:]:
                    m2.markdown(f"""
                    <div style="margin-bottom:14px;">
                      <div class="detail-label">{col_name}</div>
                      <div class="detail-value">{val if (val == val and val != '') else '—'}</div>
                    </div>""", unsafe_allow_html=True)
            if notes_col and row.get(notes_col):
                st.markdown(f"""
                <div class="detail-card" style="margin-top:8px;background:#FFFBF0;border-left:3px solid #D97706;">
                  <div class="detail-label">Notes</div>
                  <div style="font-size:13px;color:#374151;margin-top:4px;">{row.get(notes_col)}</div>
                </div>""", unsafe_allow_html=True)

        with tab2:
            if res_df is not None and proj_id_col and pid_col_r:
                proj_pid = row.get(proj_id_col)
                proj_res = res_df[res_df[pid_col_r] == proj_pid]
                if not proj_res.empty:
                    st.dataframe(proj_res.reset_index(drop=True),
                                 use_container_width=True, hide_index=True)
                    if team_col_r:
                        team_list = proj_res[team_col_r].dropna().unique().tolist()
                        st.markdown(
                            f"<div style='font-size:12px;color:{C['gray']};margin-top:8px;'>"
                            f"Teams: {', '.join(str(t) for t in team_list)}</div>",
                            unsafe_allow_html=True,
                        )
                else:
                    st.markdown(
                        "<div class='empty-state'>No resources linked to this project.</div>",
                        unsafe_allow_html=True,
                    )
            else:
                st.warning("Resource data unavailable or Project ID not found.")

        with tab3:
            if dep_df is not None:
                dep_pid_col = get_col(dep_df, "Project ID", "ProjectID", "ID", "Dependent Project ID")
                if proj_id_col and dep_pid_col:
                    proj_pid  = row.get(proj_id_col)
                    proj_deps = dep_df[dep_df[dep_pid_col] == proj_pid]
                    if not proj_deps.empty:
                        st.dataframe(proj_deps.reset_index(drop=True),
                                     use_container_width=True, hide_index=True)
                    else:
                        st.markdown(
                            "<div class='empty-state'>No dependencies recorded for this project.</div>",
                            unsafe_allow_html=True,
                        )
                else:
                    st.warning("Cannot match dependencies — Project ID column missing.")
            else:
                st.warning("Dependencies sheet not available.")

        with tab4:
            r1, r2 = st.columns(2)
            with r1:
                if impact_col:
                    impact_val     = pd.to_numeric(row.get(impact_col), errors="coerce")
                    impact_display = impact_val if pd.notna(impact_val) else "—"
                    impact_color   = C["teal"] if pd.notna(impact_val) and impact_val >= 3 else C["gray"]
                    st.markdown(f"""
                    <div class="detail-card">
                      <div class="detail-label">Impact Score</div>
                      <div class="kpi-value" style="font-size:28px;color:{impact_color};">{impact_display}</div>
                    </div>""", unsafe_allow_html=True)
                if effort_col:
                    effort_val     = pd.to_numeric(row.get(effort_col), errors="coerce")
                    effort_display = effort_val if pd.notna(effort_val) else "—"
                    st.markdown(f"""
                    <div class="detail-card">
                      <div class="detail-label">Effort Score</div>
                      <div class="kpi-value" style="font-size:28px;color:{C['deep_blue']};">{effort_display}</div>
                    </div>""", unsafe_allow_html=True)
            with r2:
                if delayed_impact_col:
                    di_val = row.get(delayed_impact_col, "—")
                    st.markdown(f"""
                    <div class="detail-card" style="border-left:3px solid #C0392B;">
                      <div class="detail-label">If Delayed Impact</div>
                      <div style="font-size:14px;font-weight:500;color:#C0392B;margin-top:4px;">
                        {di_val if di_val == di_val else '—'}
                      </div>
                    </div>""", unsafe_allow_html=True)
                if is_delayed:
                    st.markdown(f"""
                    <div class="detail-card" style="background:#FEF3F2;border-left:3px solid #C0392B;">
                      <div style="font-size:13px;color:#C0392B;font-weight:600;">⚠️ Delay Flag Active</div>
                      <div style="font-size:12px;color:#374151;margin-top:4px;">
                        This project is currently marked as Delayed. Review ownership and blockers.
                      </div>
                    </div>""", unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="detail-card" style="background:#F0FFF8;border-left:3px solid {C['teal']};">
                      <div style="font-size:13px;color:{C['teal']};font-weight:600;">✅ No Delay Flag</div>
                      <div style="font-size:12px;color:#374151;margin-top:4px;">
                        Project is not currently flagged as delayed.
                      </div>
                    </div>""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────
st.markdown(f"""
<hr style='border:none;border-top:1px solid #E8ECF0;margin:40px 0 16px 0;'>
<div style='font-size:11px;color:{C["gray"]};text-align:center;padding-bottom:12px;'>
  RevOps Program Dashboard · Data refreshes every 60 seconds · Source: SharePoint
</div>
""", unsafe_allow_html=True)
