"""Deck scanning tests use a synthetic deck built in memory: no confidential template needed."""
import io
import struct
import zlib
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Pt

from resco.config import Lob, load_deck_rules, load_lobs, load_ranges
from resco.deck_scan import build_map, classify_geometry, parse_title, scan_deck

RULES = load_deck_rules()


def _png() -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b"\x00\xff\xff\xff"
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _add_slide(prs, title, pics):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if title:
        tb = slide.shapes.add_textbox(Pt(10), Pt(10), Pt(500), Pt(30))
        tb.text_frame.text = title
    for name, geo in pics:
        p = slide.shapes.add_picture(io.BytesIO(_png()), *(Pt(v) for v in geo))
        p.name = name
    return slide


@pytest.mark.parametrize(
    "text,expected",
    [
        ("Energy Upstream (ENUPS) - Results", ("ENUPS", "Results")),
        ("Energy Upstream (ENUPS) – Expert Judgement", ("ENUPS", "Expert Judgement")),
        ("Energy Upstream (ENUPS) – Expert Judgement (old methodology)", ("ENUPS", "Expert Judgement")),
        ("Energy Downstream (ENDMW)– Expert Judgement (old methodology)", ("ENDMW", "Expert Judgement")),
        ("Life Medical Expenses (LIMED)  -  Patterns", ("LIMED", "Patterns")),
        ("Total Results", ("ALL", "Results")),
        ("Total - Triangles", ("ALL", "Triangles")),
    ],
)
def test_parse_title_matches(text, expected):
    p = parse_title(text)
    assert (p.code, p.type) == expected


@pytest.mark.parametrize(
    "text",
    ["Energy (ENTOT) – Gross vs Net summary – MBE", "Executive Summary (1/3)", "Contents", "Total Net – Gross vs Net summary – MBE"],
)
def test_parse_title_ignores_narrative(text):
    assert parse_title(text) is None


def test_classify_by_geometry_not_name():
    assert classify_geometry((925, 16, 65, 56), RULES) == "logo"
    assert classify_geometry((95, 100, 720, 288), RULES) == "table"
    assert classify_geometry((95.5, 99, 719, 290), RULES) == "table"  # within tolerance
    assert classify_geometry((95, 150, 720, 173), RULES) == "table_life"
    assert classify_geometry((300, 300, 100, 100), RULES) is None
    assert classify_geometry(None, RULES) is None


def _lobs():
    mk = lambda sel, rc, ej: Lob(sel, rc, True, True, True, True, bool(rc), bool(ej), ej)  # noqa: E731
    return [mk("Summary_ENUPS", "ENUPS", 5), mk("Summary_LIMED", "LIMED", 65), mk("Summary_ALL", "", None)]


def test_scan_and_map_synthetic_deck(tmp_path: Path):
    prs = Presentation()
    logo = ("Picture 4", (925, 16, 65, 56))
    _add_slide(prs, "Executive Summary (1/3)", [logo])  # narrative
    _add_slide(prs, "Energy Upstream (ENUPS) - Results", [logo, ("Picture 4", (95, 100, 720, 288)), ("Picture 7", (693, 400, 122, 50))])
    # title typo: says Results but is a Triangles layout (like the real LIMED slide)
    _add_slide(
        prs,
        "Life Medical Expenses (LIMED) – Results",
        [logo, ("Picture 4", (20, 100, 634, 382)), ("A", (670, 100, 274, 108)), ("B", (670, 220, 274, 108)), ("C", (670, 340, 274, 108))],
    )
    _add_slide(prs, "Total Results", [logo, ("P", (95, 100, 720, 288)), ("Q", (300, 300, 50, 50))])  # one stray picture
    _add_slide(prs, "Mystery (ZZZZ) - Results", [logo, ("P", (95, 100, 720, 288))])  # unknown LoB
    path = tmp_path / "t.pptx"
    prs.save(path)

    scans = scan_deck(path, RULES)
    dm = build_map(scans, _lobs(), load_ranges().life_resclasses, RULES)

    assert [m.lob for m in dm.mapped] == ["ENUPS", "LIMED", "ALL", "ZZZZ"]
    assert len(dm.narrative) == 1 and dm.narrative[0].index == 1
    enups, limed, total, mystery = dm.mapped
    assert [s.kind for s in enups.slots] == ["table", "profit"]
    assert limed.type == "Triangles" and limed.title_type == "Results"
    assert any(f.level == "WARN" and "treated as Triangles" in f.message for f in dm.flags)
    assert any(f.level == "WARN" and "unrecognised picture" in f.message and f.lob == "ALL" for f in dm.flags)
    assert any(f.level == "INFO" and "profit" in f.message and f.lob == "ALL" for f in dm.flags)
    assert any(f.level == "ERROR" and f.lob == "ZZZZ" for f in dm.flags)
    assert mystery.selector is None
    assert dm.lobs_without_slides == []


def test_alias_maps_title_typo_to_lob(tmp_path: Path):
    mk = lambda sel, rc, aliases: Lob(sel, rc, True, True, True, True, bool(rc), False, None, aliases)  # noqa: E731
    lobs = [mk("Summary_ENDWM", "ENDWM", ("ENDMW",))]
    prs = Presentation()
    _add_slide(prs, "Energy Downstream (ENDMW) - Results", [("P", (95, 100, 720, 288))])
    path = tmp_path / "a.pptx"
    prs.save(path)
    dm = build_map(scan_deck(path, RULES), lobs, load_ranges().life_resclasses, RULES)
    assert dm.mapped[0].lob == "ENDWM" and dm.mapped[0].selector == "Summary_ENDWM"
    assert any(f.level == "WARN" and "alias" in f.message for f in dm.flags) and not any(f.level == "ERROR" for f in dm.flags)


def test_moved_and_resized_pictures_are_matched_approximately(tmp_path: Path):
    mk = lambda sel, rc: Lob(sel, rc, True, True, True, True, bool(rc), False, None)  # noqa: E731
    prs = Presentation()
    logo = ("Picture 4", (925, 16, 65, 56))
    # table nudged by ~8 pt and made 6% wider; profit box moved 12 pt: both beyond the strict 4 pt tolerance
    _add_slide(prs, "Energy Upstream (ENUPS) - Results", [logo, ("Picture 4", (103, 108, 763, 288)), ("Picture 7", (705, 410, 122, 50))])
    # a decorative picture nowhere near any slot must NOT be grabbed
    _add_slide(prs, "Energy Downstream (ENDWM) - Results", [logo, ("P", (95, 100, 720, 288)), ("Deco", (10, 450, 60, 60))])
    path = tmp_path / "m.pptx"
    prs.save(path)
    dm = build_map(scan_deck(path, RULES), [mk("Summary_ENUPS", "ENUPS"), mk("Summary_ENDWM", "ENDWM")], load_ranges().life_resclasses, RULES)
    enups, endwm = dm.mapped
    assert sorted(s.kind for s in enups.slots) == ["profit", "table"]
    assert sum(1 for f in dm.flags if f.lob == "ENUPS" and f.level == "WARN" and "approximately" in f.message) == 2
    assert [s.kind for s in endwm.slots] == ["table"]
    assert any(f.lob == "ENDWM" and "unrecognised picture 'Deco'" in f.message for f in dm.flags)
