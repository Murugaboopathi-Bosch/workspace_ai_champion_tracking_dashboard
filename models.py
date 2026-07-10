"""
Data models for the Gen AI Champions Dashboard.

These mirror the structure described in the build prompt:
Region 1 (header/team block) -> ChampionMember
Region 2 (topics table)      -> TopicGroup / Subtopic
Region 3 (freeform notes)    -> plain list[str]
"""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field


class ChampionMember(BaseModel):
    """A single person listed under 'AI Champions' or 'AI Champions Network'."""

    name: str
    department: Optional[str] = None
    group: str  # "AI Champions" or "AI Champions Network"


class Subtopic(BaseModel):
    """A single row under a topic group (or an ungrouped batch)."""

    name: str
    progress_pct: Optional[float] = None  # 0-100, None if blank / no data yet
    lead: Optional[str] = None
    key_points: Optional[str] = None
    note: Optional[str] = None  # stray 6th-column value, if present


class TopicGroup(BaseModel):
    """A topic header row plus the subtopic rows that belong to it."""

    name: str
    overall_progress_pct: Optional[float] = None  # rolled-up value, already computed in-sheet
    subtopics: list[Subtopic] = Field(default_factory=list)


class DashboardData(BaseModel):
    """Full parsed snapshot of one Excel source file."""

    initiative_name: str
    owner: Optional[str] = None
    supported_by: Optional[str] = None
    champions: list[ChampionMember] = Field(default_factory=list)
    topics: list[TopicGroup] = Field(default_factory=list)
    notes: list[str] = Field(default_factory=list)
    last_synced: datetime
    source_file: str

    # --- convenience helpers used by the UI / chatbot -------------------

    def all_subtopics(self) -> list[Subtopic]:
        return [s for t in self.topics for s in t.subtopics]

    def overall_progress_pct(self) -> Optional[float]:
        """Simple average across topic-group rollups that have a value."""
        vals = [t.overall_progress_pct for t in self.topics if t.overall_progress_pct is not None]
        if not vals:
            return None
        return sum(vals) / len(vals)

    def counts_by_status(self) -> dict[str, int]:
        """Bucket every subtopic into not_started / in_progress / done."""
        buckets = {"not_started": 0, "in_progress": 0, "done": 0}
        for s in self.all_subtopics():
            if s.progress_pct is None:
                buckets["not_started"] += 1
            elif s.progress_pct >= 100:
                buckets["done"] += 1
            else:
                buckets["in_progress"] += 1
        return buckets

    def status_label(self, progress_pct: float | None) -> str:
        """4-tier status label used for badge colors in the UI."""
        if progress_pct is None or progress_pct == 0:
            return "Not Started"
        if progress_pct >= 100:
            return "Completed"
        if progress_pct >= 50:
            return "In Progress"
        return "Started"

    def counts_by_bucket(self) -> dict[str, int]:
        """4-tier bucket counts (Not Started / Started / In Progress / Completed),
        matching the badge categories used in the dashboard UI."""
        buckets = {"Not Started": 0, "Started": 0, "In Progress": 0, "Completed": 0}
        for s in self.all_subtopics():
            buckets[self.status_label(s.progress_pct)] += 1
        return buckets

    def to_summary_text(self) -> str:
        """Compact plain-text summary, used to ground the LLM chatbot / report."""
        lines = [
            f"Initiative: {self.initiative_name}",
            f"Owner: {self.owner or 'Unassigned'}",
            f"Supported by: {self.supported_by or 'Unassigned'}",
            f"Last synced: {self.last_synced.isoformat()}",
            "",
            "Team:",
        ]
        for c in self.champions:
            dept = f" ({c.department})" if c.department else ""
            lines.append(f"  - [{c.group}] {c.name}{dept}")

        lines.append("")
        lines.append("Topics:")
        for t in self.topics:
            pct = f"{t.overall_progress_pct:.0%}" if t.overall_progress_pct is not None else "n/a"
            lines.append(f"  - {t.name} (overall: {pct})")
            for s in t.subtopics:
                spct = f"{s.progress_pct:.0%}" if s.progress_pct is not None else "not started"
                lines.append(
                    f"      * {s.name} | progress: {spct} | lead: {s.lead or 'Unassigned'} "
                    f"| key points: {s.key_points or '-'}"
                )

        if self.notes:
            lines.append("")
            lines.append("Notes / Open Items:")
            for n in self.notes:
                lines.append(f"  - {n}")

        return "\n".join(lines)
