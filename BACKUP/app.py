"""
app.py
Gen AI Champions Initiative — Streamlit Dashboard
Colorful, metrics-first redesign with live Excel sync.
"""

import os
import io
from pathlib import Path
from datetime import datetime, date as date_type

import streamlit as st
import streamlit.components.v1 as components
import plotly.graph_objects as go
import plotly.express as px

from excel_parser import parse_excel, DashboardData

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Gen AI Champions Tracking Dashboard", page_icon="🚀", layout="wide")

BASE_DIR = Path(__file__).parent

DEFAULT_EXCEL_PATH = BASE_DIR / "data" / "AI Champions Initiative Report.xlsx"

EXCEL_SOURCE_PATH = os.environ.get(
    "EXCEL_SOURCE_PATH",
    str(DEFAULT_EXCEL_PATH)
)

STATUS_COLORS = {
    "Not Started": "#94A3B8",
    "Started":     "#F59E0B",
    "In Progress": "#3B82F6",
    "Completed":   "#10B981",
    "Blocked":     "#F43F5E",
    "Overdue":     "#DC2626",
}

import re

def _html(s: str) -> str:
    """Collapse a multi-line, indented HTML f-string into single-line HTML before
    handing it to st.markdown. Streamlit runs content through a Markdown parser
    first, and Markdown treats a blank line followed by 4+ spaces of indentation
    as a literal code block -- which is exactly what indented Python f-strings
    produce. Flattening to one line with no leading whitespace avoids that
    entirely, since HTML doesn't need whitespace between block-level tags."""
    return re.sub(r'\n\s*', '', s.strip())


def status_for(pct):
    if pct is None or pct == 0:
        return "Not Started"
    if pct < 25:
        return "Started"
    if pct < 100:
        return "In Progress"
    return "Completed"


def _parse_display_date(s: str | None):
    """Parse the display-string date (as stored by excel_parser) back into a
    date object for comparison. Returns None if it can't be parsed."""
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d %b %Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None


def resolve_status(s):
    """Progress-based status, upgraded to 'Overdue' if the End Date has passed
    and the activity isn't already Completed."""
    base_status = status_for(s.progress_pct)
    if base_status == "Completed":
        return base_status, False
    end = _parse_display_date(s.end_date)
    if end is not None and end < date_type.today():
        return "Overdue", True
    return base_status, False

# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
h1, h2, h3 { font-family: 'Poppins', sans-serif !important; }

.block-container { padding-top: 1.6rem; }

.st-key-fixed_header {
    position: sticky;
    top: 0;
    z-index: 999;
    background: #FFFFFF;
    padding: 10px 0 8px 0;
    margin-bottom: 8px;
    border-bottom: 1px solid #EEF0F6;
    box-shadow: 0 4px 10px rgba(0,0,0,0.04);
}


.kpi-card {
    border-radius: 18px;
    padding: 20px 22px;
    color: white;
    box-shadow: 0 8px 20px rgba(0,0,0,0.10);
    min-height: 110px;
}
.kpi-label { font-size: 12.5px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; opacity: 0.85; }
.kpi-value { font-family: 'Poppins', sans-serif; font-size: 30px; font-weight: 800; margin-top: 4px; }
.kpi-sub { font-size: 11.5px; opacity: 0.85; margin-top: 4px; }

.grad-violet  { background: linear-gradient(135deg, #7C3AED, #A78BFA); }
.grad-teal    { background: linear-gradient(135deg, #0D9488, #2DD4BF); }
.grad-amber   { background: linear-gradient(135deg, #D97706, #FBBF24); }
.grad-rose    { background: linear-gradient(135deg, #E11D48, #FB7185); }
.grad-blue    { background: linear-gradient(135deg, #2563EB, #60A5FA); }
.grad-red     { background: linear-gradient(135deg, #B91C1C, #EF4444); }

.topic-card {
    background: white;
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
    margin-bottom: 16px;
    border-left: 6px solid #A78BFA;
}
.topic-header-row { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
.topic-title { font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 18px; color: #1E1B2E; }
.topic-overall-badge {
    font-family: 'Poppins', sans-serif; font-weight: 800; font-size: 16px;
    padding: 6px 18px; border-radius: 20px; white-space: nowrap;
}
.topic-overall-track { height: 10px; border-radius: 20px; background: #EEF0F6; overflow: hidden; margin-bottom: 16px; }
.topic-overall-fill { height: 100%; border-radius: 20px; }

/* Fixed-width grid so every row's columns line up vertically, regardless of
   how long any individual piece of text is. Overflow truncates with an
   ellipsis instead of pushing later columns out of alignment. */
.sub-grid-cols { grid-template-columns: 1fr 130px 46px 220px 130px 96px; }

.sub-header {
    display: grid;
    align-items: center;
    gap: 10px;
    padding: 4px 0 8px 0;
    font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px;
    color: #9CA3AF; border-bottom: 1px solid #F1F1F6;
}

.sub-row {
    display: grid;
    align-items: center;
    gap: 10px;
    padding: 12px 0;
    border-top: 1px solid #F1F1F6;
}
.sub-name, .sub-lead, .sub-dates {
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.sub-name { font-weight: 700; font-size: 14.5px; color: #1E1B2E; }
.sub-lead { font-size: 13px; color: #6B7280; }
.sub-dates { font-size: 12px; color: #6B7280; font-weight: 600; }
.pbar-track { width: 100%; height: 9px; border-radius: 20px; background: #EEF0F6; overflow: hidden; }
.pbar-fill  { height: 100%; border-radius: 20px; }
.pbar-pct   { font-size: 12px; font-weight: 700; text-align: right; }

.status-chip {
    display: inline-block; padding: 3px 10px; border-radius: 20px;
    font-size: 10.5px; font-weight: 700; color: white; white-space: nowrap;
    justify-self: start;
}

.member-chip {
    display: inline-flex; align-items: center; gap: 8px;
    background: #F8F7FC; border-radius: 30px; padding: 6px 14px 6px 6px;
    margin: 0 8px 8px 0; font-size: 12.5px; font-weight: 600; color: #1E1B2E;
}
.member-avatar {
    width: 26px; height: 26px; border-radius: 50%; display:flex; align-items:center; justify-content:center;
    color: white; font-size: 11px; font-weight: 700; font-family: 'Poppins', sans-serif;
}

.stButton button {
    white-space: nowrap;
    min-width: 130px;
}

.sidebar-team-group { font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; color: #9CA3AF; font-weight: 700; margin: 16px 0 8px 0; }
.sidebar-member {
    display: flex; align-items: center; gap: 10px;
    padding: 8px 6px; border-radius: 10px; margin-bottom: 4px;
}
.sidebar-member:hover { background: #F8F7FC; }
.sidebar-avatar {
    width: 30px; height: 30px; border-radius: 50%; flex-shrink: 0;
    display:flex; align-items:center; justify-content:center;
    color: white; font-size: 11.5px; font-weight: 700; font-family: 'Poppins', sans-serif;
}
.sidebar-member-name { font-size: 13px; font-weight: 700; color: #1E1B2E; line-height: 1.2; }
.sidebar-member-dept { font-size: 11px; color: #9CA3AF; }

.note-card {
    background: #FFFBEA; border-left: 5px solid #FBBF24; border-radius: 10px;
    padding: 12px 16px; margin-bottom: 10px; font-size: 13px; color: #4B4B2E;
}

.docupedia-link {
    display: inline-flex; align-items: center; justify-content: center;
    width: 28px; height: 28px; border-radius: 6px; background: #EEF2FF;
    color: #4F46E5; font-size: 14px;
    text-decoration: none; transition: all 0.2s;
}
.docupedia-link:hover {
    background: #4F46E5;
}

.expand-btn {
    cursor: pointer; user-select: none; display: inline-flex;
    align-items: center; justify-content: center;
    width: 24px; height: 24px; border-radius: 50%;
    background: #F3F4F6; color: #6B7280;
    font-weight: bold; transition: all 0.2s;
    margin-right: 8px;
}
.expand-btn:hover {
    background: #E5E7EB; transform: scale(1.1);
}
.expand-btn.expanded {
    background: #7C3AED; color: white;
}

.sessions-container {
    margin: 12px 0 0 40px; padding: 12px;
    background: #F9FAFB; border-radius: 8px;
    border-left: 3px solid #A78BFA;
}
.session-row {
    display: grid;
    grid-template-columns: minmax(320px, 2.4fr) minmax(120px, 1fr) minmax(130px, 1fr) minmax(180px, 1.4fr) minmax(160px, 1.2fr);
    width: 100%;
    align-items: center;
    column-gap: 16px;
    padding: 12px 18px;
    margin-bottom: 8px;
    box-sizing: border-box;
    background: white;
    border-radius: 6px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}

.session-name,
.session-date,
.session-status,
.session-moderator,
.session-link {
    width: 100%;
    min-width: 0;
    text-align: left;
    justify-self: stretch;
    align-self: center;
}
.session-name {
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    font-size: 13px;
    color: #1E1B2E;
    white-space: normal;
    overflow: visible;
    text-overflow: clip;
    line-height: 1.4;
}
.session-date, .session-moderator, .session-link {
    font-family: 'Inter', sans-serif;
    font-size: 13px;
    font-weight: 600;
    color: #6B7280;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.session-status {
    font-family: 'Inter', sans-serif;
    font-size: 13px;
}
.session-link a {
    font-family: 'Inter', sans-serif;
    font-size: 13px;
    font-weight: 600;
    color: #4F46E5;
    text-decoration: none;
}
.session-count-link {
    margin-left: 8px;
    font-size: 11px;
    font-weight: 700;
    color: #7C3AED !important;
    text-decoration: underline;
    text-underline-offset: 2px;
    cursor: pointer;
}
.session-count-link:hover {
    color: #5B21B6 !important;
    text-decoration-thickness: 2px;
}
.session-anchor {
    display: block;
    position: relative;
    top: -16px;
    visibility: hidden;
}
.session-status-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-family: 'Inter', sans-serif;
    font-size: 12px;
    font-weight: 700;
    white-space: nowrap;
}
.session-completed { background: #D1FAE5; color: #065F46; }
.session-ongoing { background: #DBEAFE; color: #1E40AF; }
.session-planned { background: #FEF3C7; color: #92400E; }

</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# DATA LOADING
# ---------------------------------------------------------------------------
def _diagnose_path(path: str) -> dict:
    info = {"path": path, "exists": os.path.exists(path)}
    if info["exists"]:
        info["size_bytes"] = os.path.getsize(path)
        info["extension"] = os.path.splitext(path)[1]
    return info


def _friendly_error(e: Exception) -> str:
    if isinstance(e, PermissionError):
        return (
            "**Permission denied** — this almost always means the Excel file is currently "
            "open in Excel (by you or someone else on the shared copy), which locks it. "
            "Close the file, delete any leftover `~$<filename>.xlsx` lock file in the same "
            "folder if present, wait for OneDrive to finish syncing, then try again."
        )
    return str(e)


@st.cache_data(show_spinner=False)
def load_data(path: str) -> DashboardData:
    return parse_excel(path)


def refresh_data(path: str):
    load_data.clear()
    st.session_state["data"] = load_data(path)
    st.session_state["last_synced"] = datetime.now()


if "data" not in st.session_state:
    try:
        st.session_state["data"] = load_data(EXCEL_SOURCE_PATH)
        st.session_state["last_synced"] = datetime.now()
    except PermissionError:
        st.session_state["data"] = None
        st.session_state["load_error_type"] = "permission"
    except Exception as e:
        st.session_state["data"] = None
        st.session_state["load_error_type"] = "other"
        st.session_state["load_error"] = _friendly_error(e)
        st.session_state["path_diagnostic"] = _diagnose_path(EXCEL_SOURCE_PATH)

data: DashboardData | None = st.session_state.get("data")

# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------
with st.container(key="fixed_header"):
    col_title, col_sync = st.columns([4, 1.4], vertical_alignment="center")
    with col_title:
        st.markdown(f"## 🚀 {data.initiative_name if data else 'Gen AI Champions Initiative'}")
        if data:
            st.caption(f"Owner: **{data.owner or '—'}** · Supported by: **{data.supported_by or '—'}**")
    with col_sync:
        if st.button("🔄 Sync Now", width='stretch'):
            try:
                refresh_data(EXCEL_SOURCE_PATH)
                st.success("Synced!")
                st.rerun()
            except PermissionError:
                filename = os.path.basename(EXCEL_SOURCE_PATH)
                st.toast(f"⚠️ {filename} is open — please close it", icon="⚠️")
                st.warning(f"**Please close \"{filename}\" in Excel**, then click Sync Now again. The file can't be read while it's open.")
            except Exception as e:
                st.session_state["path_diagnostic"] = _diagnose_path(EXCEL_SOURCE_PATH)
                st.error(f"Sync failed.\n\n{_friendly_error(e)}")
                st.json(st.session_state["path_diagnostic"])

    last_synced = st.session_state.get("last_synced")
    if last_synced:
        st.caption(f"Last synced: {last_synced.strftime('%d %b %Y, %H:%M')}")

if data is None:
    if st.session_state.get("load_error_type") == "permission":
        filename = os.path.basename(EXCEL_SOURCE_PATH)
        st.warning(f"**Please close \"{filename}\" in Excel**, then click Sync Now above. The file can't be read while it's open.")
    else:
        st.error(
            f"Couldn't read the file at `{EXCEL_SOURCE_PATH}`.\n\n"
            f"{st.session_state.get('load_error', 'Unknown error.')}"
        )
        diag = st.session_state.get("path_diagnostic")
        if diag:
            st.caption("Path diagnostic:")
            st.json(diag)
    st.stop()

# ---------------------------------------------------------------------------
# TEAM (sidebar — persistent left panel, visible on every view)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 👥 Team")
    group_colors = {"AI Champions": "#7C3AED", "AI Champions Network": "#0D9488"}
    for group in ["AI Champions", "AI Champions Network"]:
        members = [c for c in data.champions if c.group == group]
        if not members:
            continue
        st.markdown(f'<div class="sidebar-team-group">{group} ({len(members)})</div>', unsafe_allow_html=True)
        for m in members:
            initials = "".join([p[0] for p in m.name.split()][:2]).upper()
            st.markdown(_html(f"""
            <div class="sidebar-member">
                <div class="sidebar-avatar" style="background:{group_colors[group]};">{initials}</div>
                <div>
                    <div class="sidebar-member-name">{m.name}</div>
                    <div class="sidebar-member-dept">{m.department or ''}</div>
                </div>
            </div>
            """), unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------------------------
all_subs = [s for t in data.topics for s in t.subtopics]
total = len(all_subs)

# Resolve each sub-topic's status ONCE here (progress-based, upgraded to
# "Overdue" if its End Date has passed) so KPIs, charts, and the Topics
# section below are always counting things the exact same way.
resolved = [resolve_status(s)[0] for s in all_subs]

completed = resolved.count("Completed")
in_progress = resolved.count("In Progress")
not_started = resolved.count("Not Started")
started = resolved.count("Started")
overdue = resolved.count("Overdue")

# Use the same rolled-up progress value(s) already shown in the Topics section below,
# rather than recomputing our own average -- avoids showing two different "overall progress"
# numbers for the same thing.
topic_level_pcts = [t.overall_progress_pct for t in data.topics if t.overall_progress_pct is not None]
if topic_level_pcts:
    avg_progress = round(sum(topic_level_pcts) / len(topic_level_pcts), 1)
else:
    avg_progress = round(sum((s.progress_pct or 0) for s in all_subs) / total, 1) if total else 0
champ_count = sum(1 for c in data.champions if c.group == "AI Champions")
net_count = sum(1 for c in data.champions if c.group == "AI Champions Network")

kpi_defs = [
    ("grad-violet", "Overall Progress", f"{avg_progress}%", f"{total} sub-topics tracked"),
    ("grad-blue",   "In Progress",      in_progress,        "sub-topics actively moving"),
    ("grad-amber",  "Not Started",       not_started,        "need an owner assigned"),
    ("grad-teal",   "Completed",         completed,          "fully wrapped up"),
]
if overdue > 0:
    kpi_defs.append(("grad-red", "Overdue", overdue, "past End Date, not completed"))
kpi_defs.append(("grad-rose", "Team Size", len(data.champions), f"{champ_count} champions · {net_count} network"))

kpi_cols = st.columns(len(kpi_defs))
for col, (grad, label, value, sub) in zip(kpi_cols, kpi_defs):
    with col:
        st.markdown(_html(f"""
        <div class="kpi-card {grad}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """), unsafe_allow_html=True)


st.write("")

# ---------------------------------------------------------------------------
# CHARTS ROW
# ---------------------------------------------------------------------------
c1, c2 = st.columns([1, 1.4])

with c1:
    st.markdown("#### Status Mix")
    status_counts = {
        "Not Started": not_started,
        "Started": started,
        "In Progress": in_progress,
        "Overdue": overdue,
        "Completed": completed,
    }
    # Drop zero-count slices so the legend doesn't show empty entries
    status_counts = {k: v for k, v in status_counts.items() if v > 0}
    fig = go.Figure(data=[go.Pie(
        labels=list(status_counts.keys()),
        values=list(status_counts.values()),
        hole=0.55,
        marker=dict(colors=[STATUS_COLORS[k] for k in status_counts.keys()]),
        textinfo="value+percent",
    )])
    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=280, showlegend=True,
                       legend=dict(orientation="h", y=-0.15))
    st.plotly_chart(fig, width='stretch')

with c2:
    st.markdown("#### Progress by Sub-Topic")
    sorted_subs = sorted(all_subs, key=lambda s: s.progress_pct or 0, reverse=True)
    names = [s.name for s in sorted_subs]
    pcts = [s.progress_pct or 0 for s in sorted_subs]
    colors = [STATUS_COLORS[resolve_status(s)[0]] for s in sorted_subs]
    fig2 = go.Figure(go.Bar(
        x=pcts, y=names, orientation="h",
        marker=dict(color=colors),
        text=[f"{p}%" for p in pcts], textposition="outside",
    ))
    fig2.update_layout(margin=dict(t=10, b=10, l=10, r=30), height=280,
                        xaxis=dict(range=[0, 110], showgrid=False, title=None),
                        yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig2, width='stretch')

st.divider()

# ---------------------------------------------------------------------------
# TOPICS
# ---------------------------------------------------------------------------
st.markdown("### 📌 Topics & Sub-Topics")

for topic_idx, topic in enumerate(data.topics):
    overall = topic.overall_progress_pct
    overall_color = STATUS_COLORS[status_for(overall)] if overall is not None else "#94A3B8"
    overall_display = overall if overall is not None else 0
    
    # Build all subtopic rows HTML first
    subtopic_rows_html = ""
    subtopics_with_sessions = []
    
    for idx, s in enumerate(topic.subtopics):
        pct = s.progress_pct or 0
        status, is_overdue = resolve_status(s)
        color = STATUS_COLORS[status]
        lead = s.lead or "Unassigned"
        has_sessions = len(s.sessions) > 0

        # Completion Timeline -- End Date only (Start Date isn't shown here per
        # request). Always render this cell, even when empty, so the fixed
        # grid stays aligned across every row.
        if s.end_date:
            timeline_color = STATUS_COLORS["Overdue"] if is_overdue else "#6B7280"
            timeline_text = f"{s.end_date}" + (" (Overdue)" if is_overdue else "")
        else:
            timeline_color = "#6B7280"
            timeline_text = "—"
        dates_html = f'<div class="sub-dates" style="color:{timeline_color};">{timeline_text}</div>'

        
        # Session indicator
        session_indicator = ""
        if has_sessions:
            session_anchor_id = f"session-{topic_idx}-{idx}"
            session_indicator = (
                f'<a class="session-count-link" '
                f'href="#{session_anchor_id}" '
                f'data-session-target="{session_anchor_id}" '
                f'data-session-name="{s.name}" '
                f'title="View {len(s.sessions)} session(s)">'
                f'({len(s.sessions)} sessions)</a>'
            )
        
        subtopic_rows_html += _html(f"""
        <div class="sub-row sub-grid-cols">
            <div class="sub-name">{s.name}{session_indicator}</div>
            <div class="pbar-track"><div class="pbar-fill" style="width:{pct}%; background:{color};"></div></div>
            <div class="pbar-pct" style="color:{color};">{pct}%</div>
            <div class="sub-lead">👤 {lead}</div>
            {dates_html}
            <div><span class="status-chip" style="background:{color};">{status}</span></div>
        </div>
        """)
        
        if has_sessions:
            subtopics_with_sessions.append((s, idx))
    
    # Render complete topic card with all subtopics
    st.markdown(_html(f"""
    <div class="topic-card" style="border-left-color:{overall_color};">
        <div class="topic-header-row">
            <div class="topic-title">{topic.name}</div>
            <div class="topic-overall-badge" style="background:{overall_color}22; color:{overall_color};">
                Overall: {overall if overall is not None else '—'}%
            </div>
        </div>
        <div class="topic-overall-track">
            <div class="topic-overall-fill" style="width:{overall_display}%; background:{overall_color};"></div>
        </div>
        <div class="sub-header sub-grid-cols">
            <div>Activity</div><div style="grid-column: span 2;">Progress</div><div>Lead</div><div>Completion Timeline</div><div>Status</div>
        </div>
        {subtopic_rows_html}
    </div>
    """), unsafe_allow_html=True)
    
    # Now render sessions using Streamlit expanders (outside the HTML card)
    for s, idx in subtopics_with_sessions:
        session_anchor_id = f"session-{topic_idx}-{idx}"
        st.markdown(
            f'<span id="{session_anchor_id}" class="session-anchor"></span>',
            unsafe_allow_html=True,
        )
        with st.expander(
            f"📋 {s.name} — View {len(s.sessions)} Session(s)",
            expanded=False,
        ):
            for session in s.sessions:
                # Determine session status badge
                session_status = session.status or "Unknown"
                status_lower = session_status.lower()
                if "complete" in status_lower:
                    status_class = "session-completed"
                elif "ongoing" in status_lower or "progress" in status_lower:
                    status_class = "session-ongoing"
                else:
                    status_class = "session-planned"
                
                session_date = session.date or "—"
                session_moderator = session.moderator or "—"
                session_link_html = (
                    f'<a href="{session.docupedia_link}" target="_blank">📄 Open Docupedia</a>'
                    if session.docupedia_link else "—"
                )
                
                st.markdown(_html(f"""
                <div class="session-row">
                    <div class="session-name">▸ {session.session_name}</div>
                    <div class="session-date">{session_date}</div>
                    <div class="session-status">
                        <span class="session-status-badge {status_class}">{session_status}</span>
                    </div>
                    <div class="session-moderator">👤 {session_moderator}</div>
                    <div class="session-link">{session_link_html}</div>
                </div>
                """), unsafe_allow_html=True)


# Browser-side session navigation: intercept the session-count link, open the
# matching Streamlit expander and smooth-scroll to it without a Streamlit rerun.
components.html(
    """
    <script>
    (() => {
        const doc = window.parent.document;
        if (doc.__sessionNavigationInstalled) return;
        doc.__sessionNavigationInstalled = true;

        doc.addEventListener("click", (event) => {
            const link = event.target.closest("a[data-session-target]");
            if (!link) return;

            event.preventDefault();
            const targetId = link.getAttribute("data-session-target");
            const sessionName = link.getAttribute("data-session-name");
            const anchor = doc.getElementById(targetId);
            if (!anchor) return;

            // Match the Streamlit expander by its visible activity title.
            // This is more reliable across Streamlit DOM wrapper changes.
            const summaries = Array.from(doc.querySelectorAll("details > summary"));
            const summary = summaries.find((item) =>
                item.innerText &&
                item.innerText.toLowerCase().includes(sessionName.toLowerCase())
            );
            const expander = summary ? summary.closest("details") : null;

            if (summary && expander && !expander.open) {
                summary.click();
            }

            window.setTimeout(() => {
                (expander || anchor).scrollIntoView({
                    behavior: "smooth",
                    block: "start"
                });
            }, 180);
        });
    })();
    </script>
    """,
    height=0,
    width=0,
)

# ---------------------------------------------------------------------------
# NOTES
# ---------------------------------------------------------------------------
if data.notes:
    st.divider()
    st.markdown("### 📝 Notes & Open Items")
    for n in data.notes:
        st.markdown(f'<div class="note-card">💡 {n}</div>', unsafe_allow_html=True)