"""
Small HTML builder helpers used by app.py to render the styled dashboard
(stat cards, badges, progress bars, member cards, topics table). Kept
separate from app.py so the Streamlit page logic stays readable.

All functions return plain HTML strings meant to be passed to
st.markdown(..., unsafe_allow_html=True).
"""

from __future__ import annotations

import html

from models import DashboardData, Subtopic, TopicGroup

_STATUS_TO_CLASS = {
    "Completed": "completed",
    "In Progress": "inprogress",
    "Started": "started",
    "Not Started": "notstarted",
}


def _esc(value: str | None) -> str:
    return html.escape(value) if value else ""


def initials(name: str) -> str:
    parts = [p for p in name.split() if p]
    if not parts:
        return "?"
    return "".join(p[0].upper() for p in parts[:2])


def status_badge(status_label: str) -> str:
    css_class = _STATUS_TO_CLASS.get(status_label, "notstarted")
    return (
        f'<span class="badge {css_class}"><span class="dot"></span>{_esc(status_label)}</span>'
    )


def progress_bar(pct: float | None, status_label: str) -> str:
    css_class = _STATUS_TO_CLASS.get(status_label, "notstarted")
    color_var = f"var(--{css_class})"
    p = max(0, min(100, round(pct))) if pct is not None else 0
    return (
        '<div class="progress-wrap">'
        '<div class="progress-track">'
        f'<div class="progress-fill" style="width:{p}%;background:{color_var};"></div>'
        "</div>"
        f'<div class="progress-pct" style="color:{color_var};">{p}%</div>'
        "</div>"
    )


def stat_card(label: str, value: str, sub: str = "") -> str:
    sub_html = f'<div class="sub">{_esc(sub)}</div>' if sub else ""
    return (
        '<div class="stat-card">'
        f'<div class="label">{_esc(label)}</div>'
        f'<div class="value">{_esc(value)}</div>'
        f"{sub_html}"
        "</div>"
    )


def stat_row(cards_html: list[str]) -> str:
    return '<div class="stat-row">' + "".join(cards_html) + "</div>"


def member_card(name: str, department: str | None, role: str) -> str:
    dept_html = f'<div class="dept">{_esc(department)}</div>' if department else ""
    return (
        '<div class="member-card">'
        f'<div class="avatar">{_esc(initials(name))}</div>'
        f'<div class="name">{_esc(name)}</div>'
        f'<div class="role">{_esc(role)}</div>'
        f"{dept_html}"
        "</div>"
    )


def topics_table(data: DashboardData, topics: list[TopicGroup]) -> str:
    rows = []
    for t in topics:
        group_status = data.status_label(t.overall_progress_pct)
        rows.append(
            "<tr>"
            f'<td class="topic-name">{_esc(t.name)} <span class="cell-muted">(group)</span></td>'
            f"<td>{progress_bar(t.overall_progress_pct, group_status)}</td>"
            f"<td>{status_badge(group_status)}</td>"
            '<td class="cell-muted">Rolled-up average</td>'
            "</tr>"
        )
        for s in t.subtopics:
            rows.append(_subtopic_row(data, s))

    header = (
        "<tr><th>Topic</th><th>Progress</th><th>Status</th><th>Lead / Key Points</th></tr>"
    )
    return f'<table class="topics-table">{header}{"".join(rows)}</table>'


def _subtopic_row(data: DashboardData, s: Subtopic) -> str:
    status = data.status_label(s.progress_pct)
    lead_html = (
        f'<span class="cell-muted">{_esc(s.lead)}</span>'
        if s.lead and s.lead != "Unassigned"
        else '<span class="muted-italic">Unassigned</span>'
    )
    key_points_html = (
        f'<span class="cell-muted">{_esc(s.key_points)}</span>'
        if s.key_points and s.key_points not in ("—", None)
        else '<span class="muted-italic">No update yet</span>'
    )
    note_html = f'<span class="note-tag">{_esc(s.note)}</span>' if s.note else ""
    return (
        "<tr>"
        f'<td class="topic-name">&nbsp;&nbsp;↳ {_esc(s.name)}{note_html}</td>'
        f"<td>{progress_bar(s.progress_pct, status)}</td>"
        f"<td>{status_badge(status)}</td>"
        f"<td>{lead_html}<br>{key_points_html}</td>"
        "</tr>"
    )
