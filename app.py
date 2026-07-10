"""
app.py
Gen AI Champions Initiative — Streamlit Dashboard
Colorful, metrics-first redesign with live Excel sync.
"""

import os
import io
from pathlib import Path
from datetime import datetime

import streamlit as st
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
}

def status_for(pct):
    if pct is None or pct == 0:
        return "Not Started"
    if pct < 25:
        return "Started"
    if pct < 100:
        return "In Progress"
    return "Completed"

# ---------------------------------------------------------------------------
# STYLE
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700;800&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"]  { font-family: 'Inter', sans-serif; }
h1, h2, h3 { font-family: 'Poppins', sans-serif !important; }

.block-container { padding-top: 1.6rem; }

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

.sub-row { display:flex; align-items:center; gap:12px; padding: 12px 0; border-top: 1px solid #F1F1F6; }
.sub-name { flex: 2; font-weight: 700; font-size: 14.5px; color: #1E1B2E; }
.sub-lead { flex: 1.2; font-size: 13px; color: #6B7280; }
.sub-key  { flex: 2; font-size: 13px; color: #6B7280; }
.pbar-track { flex: 1.3; height: 9px; border-radius: 20px; background: #EEF0F6; overflow: hidden; }
.pbar-fill  { height: 100%; border-radius: 20px; }
.pbar-pct   { width: 40px; font-size: 12px; font-weight: 700; text-align: right; }

.status-chip {
    display: inline-block; padding: 3px 10px; border-radius: 20px;
    font-size: 10.5px; font-weight: 700; color: white; white-space: nowrap;
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
            st.markdown(f"""
            <div class="sidebar-member">
                <div class="sidebar-avatar" style="background:{group_colors[group]};">{initials}</div>
                <div>
                    <div class="sidebar-member-name">{m.name}</div>
                    <div class="sidebar-member-dept">{m.department or ''}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------------------------
all_subs = [s for t in data.topics for s in t.subtopics]
total = len(all_subs)
completed = sum(1 for s in all_subs if status_for(s.progress_pct) == "Completed")
in_progress = sum(1 for s in all_subs if status_for(s.progress_pct) == "In Progress")
not_started = sum(1 for s in all_subs if status_for(s.progress_pct) == "Not Started")

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

k1, k2, k3, k4, k5 = st.columns(5)
kpi_defs = [
    (k1, "grad-violet", "Overall Progress", f"{avg_progress}%", f"{total} sub-topics tracked"),
    (k2, "grad-blue",   "In Progress",      in_progress,        "sub-topics actively moving"),
    (k3, "grad-amber",  "Not Started",       not_started,        "need an owner assigned"),
    (k4, "grad-teal",   "Completed",         completed,          "fully wrapped up"),
    (k5, "grad-rose",   "Team Size",         len(data.champions),f"{champ_count} champions · {net_count} network"),
]
for col, grad, label, value, sub in kpi_defs:
    with col:
        st.markdown(f"""
        <div class="kpi-card {grad}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

st.write("")

# ---------------------------------------------------------------------------
# CHARTS ROW
# ---------------------------------------------------------------------------
c1, c2 = st.columns([1, 1.4])

with c1:
    st.markdown("#### Status Mix")
    status_counts = {"Not Started": not_started, "Started": sum(1 for s in all_subs if status_for(s.progress_pct)=="Started"),
                      "In Progress": in_progress, "Completed": completed}
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
    names = [s.name for s in all_subs]
    pcts = [s.progress_pct or 0 for s in all_subs]
    colors = [STATUS_COLORS[status_for(p)] for p in pcts]
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

st.divider()

# ---------------------------------------------------------------------------
# TOPICS
# ---------------------------------------------------------------------------
st.markdown("### 📌 Topics & Sub-Topics")
for topic in data.topics:
    overall = topic.overall_progress_pct
    overall_color = STATUS_COLORS[status_for(overall)] if overall is not None else "#94A3B8"
    overall_display = overall if overall is not None else 0
    st.markdown(f"""
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
    """, unsafe_allow_html=True)

    for s in topic.subtopics:
        pct = s.progress_pct or 0
        status = status_for(s.progress_pct)
        color = STATUS_COLORS[status]
        lead = s.lead or "Unassigned"
        key = s.key_points or "—"
        st.markdown(f"""
        <div class="sub-row">
            <div class="sub-name">{s.name}</div>
            <div class="pbar-track"><div class="pbar-fill" style="width:{pct}%; background:{color};"></div></div>
            <div class="pbar-pct" style="color:{color};">{pct}%</div>
            <div class="sub-lead">👤 {lead}</div>
            <div class="sub-key">{key}</div>
            <div><span class="status-chip" style="background:{color};">{status}</span></div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# NOTES
# ---------------------------------------------------------------------------
if data.notes:
    st.divider()
    st.markdown("### 📝 Notes & Open Items")
    for n in data.notes:
        st.markdown(f'<div class="note-card">💡 {n}</div>', unsafe_allow_html=True)