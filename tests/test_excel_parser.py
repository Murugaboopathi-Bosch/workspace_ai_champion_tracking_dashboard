"""
Unit tests for excel_parser.parse_excel.

A small fixture workbook is generated on the fly (in a temp dir) that mimics
the exact real-file layout: header block, blank separator, topics table with
a rolled-up group + an ungrouped second batch, a stray 6th-column note, and
freeform Region-3 notes.
"""

from pathlib import Path

import openpyxl
import pytest

from excel_parser import parse_excel


@pytest.fixture()
def fixture_path(tmp_path: Path) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active

    rows = [
        ["Initiative Name", "Initiative Owner", "Supported By", "AI Champions", "AI Champions Network"],
        ["Test Initiative@SWD", None, None, "Jane Doe (BD/SWD-FSB1)", "John Roe (BD/SWD-BEA1)"],
        [None, "Owner Person", "Support Person", "Second Champ (BD/SWD-FSB2)", None],
        [None, None, None, None, None],  # blank separator
        ["Topic", "Sub topics", "Progress", "Lead", "Key Points"],
        ["Group A", None, 0.5, None, None],
        [None, "Subtopic 1", 0.5, "Jane Doe (BD/SWD-FSB1)", "Some notes"],
        [None, "Subtopic 2", None, "Second Champ", None, "July"],  # stray 6th col
        [None, None, None, None, None],  # blank separator
        [None, "Ungrouped Subtopic", 0, None, None],  # second batch, no parent Topic row
        [None, None, None, None, None],
        [None, None, None, None, None],
        [None, "This is a freeform note sentence.", None, None, None],
        [None, "Another open item.", None, None, None],
    ]
    for row in rows:
        ws.append(row)

    path = tmp_path / "fixture.xlsx"
    wb.save(path)
    return str(path)


def test_parses_header_block(fixture_path: str):
    data = parse_excel(fixture_path)
    assert data.initiative_name == "Test Initiative@SWD"
    assert data.owner == "Owner Person"
    assert data.supported_by == "Support Person"

    champ_names = {c.name for c in data.champions}
    assert "Jane Doe" in champ_names
    assert "John Roe" in champ_names
    assert "Second Champ" in champ_names

    jane = next(c for c in data.champions if c.name == "Jane Doe")
    assert jane.department == "BD/SWD-FSB1"
    assert jane.group == "AI Champions"


def test_parses_topics_and_ungrouped_batch(fixture_path: str):
    data = parse_excel(fixture_path)
    assert len(data.topics) == 1

    group = data.topics[0]
    assert group.name == "Group A"
    assert group.overall_progress_pct == 50.0

    names = [s.name for s in group.subtopics]
    assert "Subtopic 1" in names
    assert "Subtopic 2" in names
    # ungrouped batch attaches to the same (most recent) topic group
    assert "Ungrouped Subtopic" in names

    sub2 = next(s for s in group.subtopics if s.name == "Subtopic 2")
    assert sub2.lead == "Second Champ"
    assert sub2.key_points == "—"  # blank rendered as em-dash placeholder
    assert sub2.note == "July"  # stray 6th column captured, doesn't error out


def test_parses_freeform_notes(fixture_path: str):
    data = parse_excel(fixture_path)
    assert "This is a freeform note sentence." in data.notes
    assert "Another open item." in data.notes


def test_handles_missing_topics_header_gracefully(tmp_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Initiative Name", "Initiative Owner", "Supported By", "AI Champions", "AI Champions Network"])
    ws.append(["Only Header@Test", None, None, None, None])
    path = tmp_path / "no_table.xlsx"
    wb.save(path)

    data = parse_excel(str(path))
    assert data.initiative_name == "Only Header@Test"
    assert data.topics == []
    assert data.notes == []
