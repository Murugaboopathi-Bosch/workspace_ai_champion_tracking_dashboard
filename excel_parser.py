"""
excel_parser.py
Parses the Gen AI Champions Initiative Excel sheet into structured data.

Handles the real-world messy layout:
  Region 1: header block (Initiative Name / Owner / Supported By / Champions / Network)
  Region 2: Topic / Sub topics / Progress / Lead / Key Points table (with topic-group rollups)
  Region 3: freeform notes after the table
"""

import re
from dataclasses import dataclass, field
from datetime import datetime
from openpyxl import load_workbook


@dataclass
class ChampionMember:
    name: str
    department: str | None = None
    group: str = "AI Champions"  # or "AI Champions Network"


@dataclass
class Subtopic:
    name: str
    progress_pct: float | None = None
    lead: str | None = None
    key_points: str | None = None
    note: str | None = None  # for stray 6th-column values


@dataclass
class TopicGroup:
    name: str
    overall_progress_pct: float | None = None
    subtopics: list[Subtopic] = field(default_factory=list)


@dataclass
class DashboardData:
    initiative_name: str
    owner: str | None
    supported_by: str | None
    champions: list[ChampionMember]
    topics: list[TopicGroup]
    notes: list[str]
    last_synced: datetime
    source_file: str


NAME_DEPT_RE = re.compile(r"^(.*?)\s*\(([^)]+)\)\s*$")


def _split_name_dept(raw: str) -> tuple[str, str | None]:
    if not raw:
        return raw, None
    m = NAME_DEPT_RE.match(raw.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return raw.strip(), None


def _is_blank_row(row) -> bool:
    return all(c is None or str(c).strip() == "" for c in row)


def parse_excel(filepath: str) -> DashboardData:
    wb = load_workbook(filepath, data_only=True)
    ws = wb[wb.sheetnames[0]]  # always take the first sheet, don't hardcode a name
    rows = list(ws.iter_rows(values_only=True))

    # ---- Region 1: header block ----
    initiative_name, owner, supported_by = None, None, None
    champions: list[ChampionMember] = []

    header_end_idx = 0
    for i, row in enumerate(rows[:15]):
        if row and str(row[0]).strip().lower() == "topic" and len(row) > 1 and str(row[1] or "").strip().lower().startswith("sub"):
            header_end_idx = i
            break
        if i == 0:
            continue  # this is the label row ("Initiative Name | Initiative Owner | ..."), skip it
        if row:
            # Initiative Name, Owner, and Supported By are NOT all on the same row in the
            # real sheet -- they're staggered across the header block (e.g. name on row 2,
            # owner/supported-by on row 3). Take the first non-empty value found for each,
            # regardless of which row it lands on.
            if len(row) > 0 and row[0] and initiative_name is None:
                initiative_name = row[0]
            if len(row) > 1 and row[1] and owner is None:
                owner = row[1]
            if len(row) > 2 and row[2] and supported_by is None:
                supported_by = row[2]
            # Champion/network names DO appear one-per-row across many rows, so collect all of them.
            if len(row) > 3 and row[3]:
                n, d = _split_name_dept(str(row[3]))
                champions.append(ChampionMember(name=n, department=d, group="AI Champions"))
            if len(row) > 4 and row[4]:
                n, d = _split_name_dept(str(row[4]))
                champions.append(ChampionMember(name=n, department=d, group="AI Champions Network"))

    # ---- Region 2: topics table ----
    topics: list[TopicGroup] = []
    current_group: TopicGroup | None = None
    notes: list[str] = []
    in_notes_section = False
    blank_run = 0

    for row in rows[header_end_idx + 1:]:
        if _is_blank_row(row):
            blank_run += 1
            if blank_run >= 2 and current_group is not None:
                in_notes_section = True
            continue
        blank_run = 0

        topic_val = row[0] if len(row) > 0 else None
        subtopic_val = row[1] if len(row) > 1 else None
        progress_val = row[2] if len(row) > 2 else None
        lead_val = row[3] if len(row) > 3 else None
        keypoints_val = row[4] if len(row) > 4 else None
        stray_val = row[5] if len(row) > 5 else None

        # Freeform note row: text sits in "sub topic" column, no progress/lead/keypoints AT ALL.
        # IMPORTANT: use "is None" here, not truthiness -- a Progress value of 0 (a real,
        # valid "Not Started" row) is falsy in Python but is NOT the same as blank/missing,
        # and must not be swept into Notes.
        is_note_row = (
            subtopic_val is not None
            and progress_val is None
            and lead_val is None
            and keypoints_val is None
            and topic_val is None
        )
        if in_notes_section or is_note_row:
            if subtopic_val:
                notes.append(str(subtopic_val).strip())
            continue

        if topic_val and not subtopic_val:
            # New topic group header row (progress here is the rolled-up average)
            current_group = TopicGroup(
                name=str(topic_val).strip(),
                overall_progress_pct=_pct(progress_val),
            )
            topics.append(current_group)
            continue

        if subtopic_val:
            if current_group is None:
                current_group = TopicGroup(name="General", overall_progress_pct=None)
                topics.append(current_group)
            current_group.subtopics.append(Subtopic(
                name=str(subtopic_val).strip(),
                progress_pct=_pct(progress_val),
                lead=str(lead_val).strip() if lead_val else None,
                key_points=str(keypoints_val).strip() if keypoints_val else None,
                note=str(stray_val).strip() if stray_val else None,
            ))

    return DashboardData(
        initiative_name=str(initiative_name) if initiative_name else "Untitled Initiative",
        owner=str(owner) if owner else None,
        supported_by=str(supported_by) if supported_by else None,
        champions=champions,
        topics=topics,
        notes=notes,
        last_synced=datetime.now(),
        source_file=filepath,
    )


def _pct(v) -> float | None:
    if v is None or v == "":
        return None
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return round(v * 100, 1) if v <= 1 else round(v, 1)