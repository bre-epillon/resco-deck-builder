"""A LoB slide that has a title but no picture yet (a slide just added to the template) gets its pictures inserted."""
from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation

from resco.config import load_deck_rules, load_lobs, load_ranges
from resco.deck_scan import build_map, scan_deck
from resco.deck_writer import Job, write_deck
from resco.verify import compare, snapshot
from test_build_pipeline import _png  # noqa: E402
from test_deck_scan import _add_slide  # noqa: E402

RULES = load_deck_rules()
LOGO = ("Picture 4", (925, 16, 65, 56))


@pytest.fixture()
def deck(tmp_path: Path) -> Path:
    prs = Presentation()
    _add_slide(prs, "Executive Summary (1/3)", [LOGO])
    _add_slide(prs, "Dubai (DBTOT) - Results", [LOGO])  # empty: just added
    _add_slide(prs, "Life India (LIIND) - Results", [LOGO])  # empty Life slide
    _add_slide(prs, "Dubai Property (DBPRO) - Expert Judgement", [LOGO])  # empty
    _add_slide(prs, "Energy Upstream (ENUPS) - Results", [LOGO, ("Picture 4", (95, 100, 720, 288)), ("Picture 7", (693, 400, 122, 50))])
    path = tmp_path / "template.pptx"
    prs.save(path)
    return path


def _map(deck: Path):
    return build_map(scan_deck(deck, RULES), load_lobs(), load_ranges().life_resclasses, RULES)


def test_empty_slides_get_slots_at_the_standard_positions(deck: Path) -> None:
    dm = _map(deck)
    by_lob = {m.lob: m for m in dm.mapped}
    dbtot, liind, dbpro, enups = by_lob["DBTOT"], by_lob["LIIND"], by_lob["DBPRO"], by_lob["ENUPS"]
    assert [(s.kind, s.new) for s in dbtot.slots] == [("table", True), ("profit", True)]
    assert dbtot.slots[0].geometry == RULES.kinds["table"].geometry and dbtot.type == "Results"
    assert [s.kind for s in liind.slots] == ["table_life", "profit"]  # Life classes use the shorter table
    assert [(s.kind, s.new) for s in dbpro.slots] == [("ej", True)] and dbpro.type == "Expert Judgement"
    assert not any(s.new for s in enups.slots)  # a slide that already has its pictures is untouched
    flagged = {f.lob for f in dm.flags if f.level == "WARN" and "no pictures" in f.message}
    assert flagged == {"DBTOT", "LIIND", "DBPRO"}
    assert not any(f.lob == "DBTOT" and "no picture matches" in f.message for f in dm.flags)


def test_pictures_are_inserted_and_verified(deck: Path, tmp_path: Path) -> None:
    dm = _map(deck)
    m = next(x for x in dm.mapped if x.lob == "DBTOT")
    jobs = [Job(m.slide, m.title, s.shape_id, s.kind, m.lob, _png(tmp_path / f"{s.kind}.png", (10, 20, 30), 8, 3), new=True, geometry=s.geometry) for s in m.slots]
    out = tmp_path / "out.pptx"
    res = write_deck(deck, out, jobs)
    assert all(r.old_size is None and r.old_sha is None and r.shape_id > 0 for r in res)

    slide = Presentation(str(out)).slides[m.slide - 1]
    pics = [s for s in slide.shapes if s.shape_type == 13 and round(s.left / 12700) != 925]  # not the logo
    assert len(pics) == 2
    assert [(round(s.left / 12700), round(s.top / 12700), round(s.width / 12700), round(s.height / 12700)) for s in pics] == [(95, 100, 720, 288), (693, 400, 122, 50)]

    v = compare(snapshot(deck), snapshot(out), {}, {}, {m.slide: 2})
    assert v.ok, v.messages
    assert v.stats["pictures_inserted"] == 2 and v.stats["touched_slides"] == 1 and v.stats["untouched_slides_identical"] == 4
    # the same build checked without telling the verification about the insertion must fail
    assert not compare(snapshot(deck), snapshot(out), {}, {}, {}).ok


def test_the_inserted_deck_is_read_back_as_a_normal_deck(deck: Path, tmp_path: Path) -> None:
    """Next quarter the pictures exist: the scan recognises them as ordinary slots (no 'new' any more)."""
    dm = _map(deck)
    m = next(x for x in dm.mapped if x.lob == "DBTOT")
    jobs = [Job(m.slide, m.title, s.shape_id, s.kind, m.lob, _png(tmp_path / f"{s.kind}.png", (1, 2, 3)), new=True, geometry=s.geometry) for s in m.slots]
    out = tmp_path / "out.pptx"
    write_deck(deck, out, jobs)
    again = build_map(scan_deck(out, RULES), load_lobs(), load_ranges().life_resclasses, RULES)
    dbtot = next(x for x in again.mapped if x.lob == "DBTOT")
    assert [(s.kind, s.new) for s in dbtot.slots] == [("table", False), ("profit", False)]
    assert not any(f.lob == "DBTOT" and f.level == "WARN" for f in again.flags)


def test_a_slide_with_an_unrecognised_picture_is_not_treated_as_empty(tmp_path: Path) -> None:
    prs = Presentation()
    _add_slide(prs, "Dubai (DBTOT) - Results", [LOGO, ("Stray", (300, 300, 50, 50))])
    path = tmp_path / "t.pptx"
    prs.save(path)
    dbtot = next(x for x in _map(path).mapped if x.lob == "DBTOT")
    assert dbtot.slots == []  # reported as unrecognised, never guessed


def test_slide_text_is_unchanged_by_the_insertion(deck: Path, tmp_path: Path) -> None:
    dm = _map(deck)
    m = next(x for x in dm.mapped if x.lob == "DBTOT")
    job = Job(m.slide, m.title, -1, "table", "DBTOT", _png(tmp_path / "t.png", (9, 9, 9)), new=True, geometry=(95.0, 100.0, 720.0, 288.0))
    out = tmp_path / "o.pptx"
    write_deck(deck, out, [job])
    texts = [s.text_frame.text for s in Presentation(str(out)).slides[m.slide - 1].shapes if s.has_text_frame]
    assert texts == ["Dubai (DBTOT) - Results"]
