import streamlit as st
import pandas as pd
import plotly.express as px
import requests
from io import BytesIO

# ─────────────────────────────────────────────
ONEDRIVE_FILE_URL = "https://emerson-my.sharepoint.com/:x:/p/savitri_lazarus/IQB7_WEjDxxfQZDKz88rVLHpAWN5slfhpjAks8zRAvuDIlY?e=WcgHiV"
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="RevOps Program Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main { background-color: #f8f9fb; }
    .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 20px 24px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        text-align: center;
    }
    .kpi-label { font-size: 13px; color: #6b7280; font-weight: 500; margin-bottom: 4px; }
    .kpi-value { font-size: 32px; font-weight: 700; color: #111827; }
    .kpi-sub { font-size: 12px; color: #9ca3af; margin-top: 2px; }
    h1 { color: #111827 !important; }
    .section-header {
        font-size: 15px;
        font-weight: 600;
        color: #374151;
        margin-bottom: 8px;
        padding-bottom: 6px;
        border-bottom: 1px solid #e5e7eb;
    }
</style>
""", unsafe_allow_html=True)


def normalize_cols(df):
    df.columns = [c.strip() for c in df.columns]
    return df


def get_col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


@st.cache_data(ttl=60)
def load_data(url):
    try:
        # Convert OneDrive share link to direct download link
        if "1drv.ms" in url or "sharepoint.com" in url or "onedrive.live.com" in url:
            if "1drv.ms" in url:
                r = requests.get(url, allow_redirects=True, timeout=15)
                direct_url = r.url
                if "download=1" not in direct_url:
                    direct_url = direct_url.replace("redir?", "download?").replace("embed?", "download?")
                    if "?" in direct_url:
                        direct_url += "&download=1"
                    else:
                        direct_url += "?download=1"
                url = direct_url
            elif "sharepoint.com" in url:
                url = url.replace("/:x:/", "/:x:/").rstrip("/")
                if "download=1" not in url:
                    url = url + ("&" if "?" in url else "?") + "download=1"

        r = requests.get(url, timeout=20)
        r.raise_for_status()
        content = BytesIO(r.content)
        sheets = {}
        for sheet in ["Projects", "Project_Resources", "Dependencies"]:
            try:
                df = pd.read_excel(content, sheet_name=sheet)
                sheets[sheet] = normalize_cols(df)
            except Exception:
                sheets[sheet] = None
        return sheets, None
    except Exception as e:
        return None, str(e)


# ── HEADER ───────────────────────────────────
st.markdown("## RevOps Program Dashboard")
st.markdown(
    "<span style='color:#6b7280;font-size:15px;'>Resource load, project risk, and dependency visibility</span>",
    unsafe_allow_html=True,
)
st.markdown("<div style='margin-bottom:4px'></div>", unsafe_allow_html=True)

if ONEDRIVE_FILE_URL == "PASTE_LINK_HERE":
    st.error("⚠️ Set your OneDrive URL in `ONEDRIVE_FILE_URL` at the top of this file.")
    st.stop()

with st.spinner("Loading data…"):
    sheets, err = load_data(ONEDRIVE_FILE_URL)

if err:
    st.error(f"Failed to load data: {err}")
    st.stop()

proj_df = sheets.get("Projects")
res_df = sheets.get("Project_Resources")
dep_df = sheets.get("Dependencies")

missing = [s for s, d in sheets.items() if d is None]
if missing:
    for m in missing:
        st.warning(f"Sheet '{m}' not found or could not be read.")

if proj_df is None:
    st.error("Cannot render dashboard without the Projects sheet.")
    st.stop()

# ── COLUMN HELPERS ───────────────────────────
owner_col = get_col(proj_df, "Owner", "owner", "PM", "Project Owner")
team_col_p = get_col(proj_df, "Team", "team", "Department")
status_col = get_col(proj_df, "Status", "status", "Project Status")
cycle_col = get_col(proj_df, "Cycle", "cycle", "Sprint", "Quarter")
priority_col = get_col(proj_df, "Priority", "priority", "Priority Type")
effort_col = get_col(proj_df, "Effort", "effort", "Effort Score")
impact_col = get_col(proj_df, "Impact", "impact", "Impact Score")
proj_id_col = get_col(proj_df, "Project ID", "ProjectID", "ID", "project_id")
proj_name_col = get_col(proj_df, "Project", "Project Name", "project", "Name")
delayed_impact_col = get_col(proj_df, "If Delayed Impact", "Delayed Impact", "delay_impact", "Impact If Delayed")

# ── SIDEBAR FILTERS ──────────────────────────
st.sidebar.markdown("### Filters")

filtered = proj_df.copy()

# Default to RevOps owner if available
if owner_col:
    owners = sorted(proj_df[owner_col].dropna().unique().tolist())
    default_owners = ["RevOps"] if "RevOps" in owners else owners
    sel_owners = st.sidebar.multiselect("Owner", owners, default=default_owners)
    if sel_owners:
        filtered = filtered[filtered[owner_col].isin(sel_owners)]

if team_col_p:
    teams = sorted(proj_df[team_col_p].dropna().unique().tolist())
    sel_teams = st.sidebar.multiselect("Team", teams, default=[])
    if sel_teams:
        filtered = filtered[filtered[team_col_p].isin(sel_teams)]

if status_col:
    statuses = sorted(proj_df[status_col].dropna().unique().tolist())
    sel_status = st.sidebar.multiselect("Status", statuses, default=[])
    if sel_status:
        filtered = filtered[filtered[status_col].isin(sel_status)]

if cycle_col:
    cycles = sorted(proj_df[cycle_col].dropna().unique().tolist())
    sel_cycles = st.sidebar.multiselect("Cycle", cycles, default=[])
    if sel_cycles:
        filtered = filtered[filtered[cycle_col].isin(sel_cycles)]

if priority_col:
    priorities = sorted(proj_df[priority_col].dropna().unique().tolist())
    sel_priorities = st.sidebar.multiselect("Priority Type", priorities, default=[])
    if sel_priorities:
        filtered = filtered[filtered[priority_col].isin(sel_priorities)]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing {len(filtered)} of {len(proj_df)} projects · Auto-refresh every 60s")

# ── KPI CARDS ────────────────────────────────
total = len(filtered)

delayed_count = 0
if status_col:
    delayed_count = filtered[status_col].str.lower().str.contains("delay", na=False).sum()

active_count = 0
if status_col:
    active_count = filtered[status_col].str.lower().str.contains("active|in progress|in-progress", na=False, regex=True).sum()

teams_count = 0
if res_df is not None:
    team_col_r = get_col(res_df, "Team", "team", "Department", "Resource Team")
    pid_col_r = get_col(res_df, "Project ID", "ProjectID", "ID")
    if proj_id_col and pid_col_r and proj_id_col in filtered.columns:
        active_pids = filtered[proj_id_col].dropna().unique()
        res_filtered = res_df[res_df[pid_col_r].isin(active_pids)]
        if team_col_r:
            teams_count = res_filtered[team_col_r].nunique()

k1, k2, k3, k4 = st.columns(4)

def kpi(col, label, value, sub=""):
    col.markdown(
        f"""<div class='kpi-card'>
        <div class='kpi-label'>{label}</div>
        <div class='kpi-value'>{value}</div>
        <div class='kpi-sub'>{sub}</div>
        </div>""",
        unsafe_allow_html=True,
    )

kpi(k1, "Total Projects", total, "in current filters")
kpi(k2, "Delayed Projects", delayed_count, "⚠️ needs attention" if delayed_count else "on track")
kpi(k3, "Active Projects", active_count, "in progress")
kpi(k4, "Teams Involved", teams_count, "across resource pool")

st.markdown("<div style='margin-top:24px'></div>", unsafe_allow_html=True)

# ── CHARTS ROW 1 ─────────────────────────────
c1, c2 = st.columns(2)

with c1:
    with st.container():
        st.markdown("<div class='section-header'>Projects by Team (Resource Pool)</div>", unsafe_allow_html=True)
        if res_df is not None:
            team_col_r = get_col(res_df, "Team", "team", "Department", "Resource Team")
            pid_col_r = get_col(res_df, "Project ID", "ProjectID", "ID")
            proj_id_r = get_col(res_df, "Project ID", "ProjectID", "ID")
            if team_col_r and pid_col_r:
                if proj_id_col and proj_id_col in filtered.columns:
                    active_pids = filtered[proj_id_col].dropna().unique()
                    chart_df = res_df[res_df[pid_col_r].isin(active_pids)]
                else:
                    chart_df = res_df.copy()
                team_counts = chart_df.groupby(team_col_r)[pid_col_r].nunique().reset_index()
                team_counts.columns = ["Team", "Projects"]
                team_counts = team_counts.sort_values("Projects", ascending=True)
                fig = px.bar(
                    team_counts, x="Projects", y="Team", orientation="h",
                    color="Projects", color_continuous_scale="Blues",
                    template="plotly_white",
                )
                fig.update_layout(
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=280,
                    showlegend=False,
                    coloraxis_showscale=False,
                    font=dict(family="Inter, sans-serif", size=12),
                    plot_bgcolor="white",
                )
                fig.update_traces(marker_line_width=0)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Missing Team or Project ID column in Project_Resources.")
        else:
            st.warning("Project_Resources sheet not available.")

with c2:
    with st.container():
        st.markdown("<div class='section-header'>Projects by Status</div>", unsafe_allow_html=True)
        if status_col:
            status_counts = filtered[status_col].value_counts().reset_index()
            status_counts.columns = ["Status", "Count"]
            color_map = {
                "Delayed": "#ef4444", "At Risk": "#f97316", "On Track": "#22c55e",
                "Active": "#3b82f6", "Complete": "#8b5cf6", "In Progress": "#06b6d4",
            }
            fig2 = px.bar(
                status_counts, x="Status", y="Count",
                color="Status", color_discrete_map=color_map,
                template="plotly_white",
            )
            fig2.update_layout(
                margin=dict(t=10, b=10, l=10, r=10),
                height=280,
                showlegend=False,
                font=dict(family="Inter, sans-serif", size=12),
                plot_bgcolor="white",
            )
            fig2.update_traces(marker_line_width=0)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.warning("Status column not found in Projects.")

# ── SCATTER ──────────────────────────────────
st.markdown("<div class='section-header'>Effort vs. Impact</div>", unsafe_allow_html=True)
if effort_col and impact_col:
    scatter_df = filtered[[c for c in [effort_col, impact_col, status_col, proj_name_col, owner_col] if c]].dropna(subset=[effort_col, impact_col])
    hover_data = {}
    if proj_name_col:
        hover_data[proj_name_col] = True
    if owner_col:
        hover_data[owner_col] = True
    fig3 = px.scatter(
        scatter_df,
        x=effort_col, y=impact_col,
        color=status_col if status_col else None,
        hover_data=hover_data,
        template="plotly_white",
        size_max=14,
        opacity=0.8,
    )
    fig3.update_traces(marker=dict(size=11, line=dict(width=1, color="white")))
    fig3.update_layout(
        margin=dict(t=10, b=20, l=10, r=10),
        height=320,
        font=dict(family="Inter, sans-serif", size=12),
        plot_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig3, use_container_width=True)
else:
    st.warning(f"Effort or Impact column not found. Found columns: {list(proj_df.columns)}")

# ── DELAYED PROJECTS TABLE ────────────────────
st.markdown("<div class='section-header'>⚠️ Delayed Projects</div>", unsafe_allow_html=True)
if status_col:
    delayed_df = filtered[filtered[status_col].str.lower().str.contains("delay", na=False)]
    table_cols = [c for c in [proj_id_col, proj_name_col, owner_col, status_col, cycle_col, impact_col, delayed_impact_col] if c]
    if not delayed_df.empty and table_cols:
        display_delayed = delayed_df[table_cols].reset_index(drop=True)
        st.dataframe(
            display_delayed,
            use_container_width=True,
            hide_index=True,
        )
    elif delayed_df.empty:
        st.success("No delayed projects in current filter set.")
    else:
        st.warning("Delayed projects exist but expected columns not found.")
else:
    st.warning("Status column unavailable — cannot compute delayed projects.")

# ── PROJECT DETAIL ────────────────────────────
st.markdown("---")
st.markdown("## Project Detail")

if proj_name_col and proj_id_col:
    proj_options = filtered[proj_name_col].dropna().unique().tolist()
elif proj_name_col:
    proj_options = filtered[proj_name_col].dropna().unique().tolist()
elif proj_id_col:
    proj_options = filtered[proj_id_col].dropna().unique().tolist()
else:
    proj_options = []

if proj_options:
    selected_proj = st.selectbox("Select a project", sorted(proj_options))

    if proj_name_col and selected_proj:
        proj_row = filtered[filtered[proj_name_col] == selected_proj]
    elif proj_id_col and selected_proj:
        proj_row = filtered[filtered[proj_id_col] == selected_proj]
    else:
        proj_row = pd.DataFrame()

    if not proj_row.empty:
        pd1, pd2 = st.columns([1, 2])

        with pd1:
            st.markdown("<div class='section-header'>Project Info</div>", unsafe_allow_html=True)
            display_row = proj_row.iloc[0]
            detail_cols = [c for c in [proj_id_col, proj_name_col, owner_col, status_col, cycle_col, priority_col, effort_col, impact_col] if c]
            for col in detail_cols:
                val = display_row.get(col, "—")
                st.markdown(f"**{col}:** {val}")

        with pd2:
            # Resources
            if res_df is not None:
                st.markdown("<div class='section-header'>Teams & Resources</div>", unsafe_allow_html=True)
                pid_col_r = get_col(res_df, "Project ID", "ProjectID", "ID")
                if proj_id_col and pid_col_r:
                    proj_pid = display_row.get(proj_id_col)
                    proj_res = res_df[res_df[pid_col_r] == proj_pid]
                    if not proj_res.empty:
                        st.dataframe(proj_res.reset_index(drop=True), use_container_width=True, hide_index=True)
                    else:
                        st.info("No resources linked to this project.")
                else:
                    st.warning("Cannot match resources — missing Project ID.")
            else:
                st.warning("Project_Resources sheet not available.")

            # Dependencies
            if dep_df is not None:
                st.markdown("<div class='section-header' style='margin-top:16px'>Dependencies</div>", unsafe_allow_html=True)
                dep_pid_col = get_col(dep_df, "Project ID", "ProjectID", "ID", "Dependent Project ID")
                if proj_id_col and dep_pid_col:
                    proj_pid = display_row.get(proj_id_col)
                    proj_deps = dep_df[dep_df[dep_pid_col] == proj_pid]
                    if not proj_deps.empty:
                        st.dataframe(proj_deps.reset_index(drop=True), use_container_width=True, hide_index=True)
                    else:
                        st.info("No dependencies recorded for this project.")
                else:
                    st.warning("Cannot match dependencies — missing Project ID.")
            else:
                st.warning("Dependencies sheet not available.")
else:
    st.info("No projects available to detail with current filters.")
