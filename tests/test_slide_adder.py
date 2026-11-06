"""Adding the slides a run needs but the template lacks (src/resco/slide_adder.py, config/slides_add.yaml)."""
from __future__ import annotations

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Pt

from resco.config import NewSlide, load_deck_rules, load_lobs, load_ranges, load_slides_to_add
from resco.deck_scan import build_map, scan_deck
from resco.deck_writer import Job, write_deck
from resco.slide_adder import add_missing_slides, missing_specs
from resco.verify import compare, compare_prepared, snapshot
from test_build_pipeline import _png  # noqa: E402
from test_deck_scan import _add_slide  # noqa: E402

RULES = load_deck_rules()
LOGO = ("Picture 4", (925, 16, 65, 56))
RESULTS = [("Picture 5", (95, 100, 720, 288))]  # a table-only Results slide, like DBENR


@pytest.fixture()
def deck(tmp_path: Path) -> Path:
    prs = Presentation()
    _add_slide(prs, "Executive Summary (1/3)", [LOGO])
    _add_slide(prs, "Dubai (DBTOT) - Gross vs Net summary - MBE", [LOGO])
    gvn = prs.slides[1]
    table = gvn.shapes.add_table(3, 3, Pt(80), Pt(100), Pt(600), Pt(120)).table
    table.cell(0, 1).text = "Gross"
    table.cell(1, 0).text = "UWY"
    table.cell(2, 0).text = "<=2019"
    table.cell(2, 1).text = "57%"
    _add_slide(prs, "Dubai Energy (DBENR) - Results", [LOGO, *RESULTS])
    _add_slide(prs, "Dubai Life (DBLIF) - Results", [LOGO, *RESULTS])
    _add_slide(prs, "Energy Upstream (ENUPS) - Results", [LOGO, *RESULTS, ("Picture 7", (693, 400, 122, 50))])
    path = tmp_path / "template.pptx"
    prs.save(path)
    return path


SPECS = [
    NewSlide("dbpro_results", "Dubai Property (DBPRO) - Results", ("Dubai Energy (DBENR) - Results",), "Dubai Life (DBLIF) - Results", "frame", "DBPRO", "results"),
    NewSlide("dbpro_net", "Dubai Property (DBPRO) – Gross vs Net summary – MBE", ("Nothing (X)", "Dubai (DBTOT) - Gross vs Net summary - MBE"), "Dubai Property (DBPRO) - Results", "full", "DBPRO", "netting"),
]


def _titles(path: Path) -> list[str]:
    return [next((sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text), "") for s in Presentation(str(path)).slides]


def test_missing_specs_ignore_dashes_and_case(deck: Path) -> None:
    present = ["dubai energy (dbenr) – results", "Dubai Property (DBPRO) - results"]
    assert [s.id for s in missing_specs(present, SPECS)] == ["dbpro_net"]


def test_new_slides_are_copies_placed_after_their_anchor(deck: Path, tmp_path: Path) -> None:
    out = tmp_path / "prepared.pptx"
    res = add_missing_slides(deck, out, SPECS, RULES)
    assert res.saved and res.added == 2
    titles = _titles(out)
    assert titles[3:6] == ["Dubai Life (DBLIF) - Results", "Dubai Property (DBPRO) - Results", "Dubai Property (DBPRO) – Gross vs Net summary – MBE"]
    assert res.added_positions == {5, 6}
    assert [e["status"] for e in res.entries] == ["added", "added"] and res.entries[0]["position"] == 5 and res.entries[1]["like"].startswith("Dubai (DBTOT)")

    prs = Presentation(str(out))
    frame, full = prs.slides[4], prs.slides[5]
    assert [s.shape_type for s in frame.shapes].count(13) == 1  # only the logo: the content picture was not copied
    assert any(getattr(s, "has_table", False) and s.has_table for s in full.shapes)  # a full copy keeps its table
    assert res.kinds_by_title and list(res.kinds_by_title.values()) == [("table",)]  # what the model had, to be inserted by the build


def test_original_slides_are_untouched_and_the_original_file_too(deck: Path, tmp_path: Path) -> None:
    before = deck.read_bytes()
    out = tmp_path / "prepared.pptx"
    res = add_missing_slides(deck, out, SPECS, RULES)
    assert deck.read_bytes() == before
    assert compare_prepared(snapshot(deck), snapshot(out), res.added_positions) == []
    assert compare_prepared(snapshot(deck), snapshot(out), {5}) != []  # claims one slide fewer: must be caught
    with pytest.raises(ValueError, match="overwrite"):
        add_missing_slides(deck, deck, SPECS, RULES)


def test_present_slide_is_not_added_twice_and_missing_model_is_reported(deck: Path, tmp_path: Path) -> None:
    out = tmp_path / "prepared.pptx"
    first = add_missing_slides(deck, out, SPECS[:1], RULES)
    again = add_missing_slides(out, tmp_path / "second.pptx", SPECS[:1], RULES)
    assert first.saved and not again.saved and again.entries[0]["status"] == "present"  # nothing missing: nothing written
    bad = NewSlide("x", "Dubai Treaty (DBTRT) - Results", ("No such model",), None, "frame", "DBTRT", "results")
    res = add_missing_slides(deck, tmp_path / "third.pptx", [bad], RULES)
    assert not res.saved and res.entries[0]["status"].startswith("failed")


def test_unknown_anchor_puts_the_slide_at_the_end(deck: Path, tmp_path: Path) -> None:
    spec = NewSlide("t", "Dubai Treaty (DBTRT) - Results", ("Dubai Energy (DBENR) - Results",), "No such slide", "frame", "DBTRT", "results")
    out = tmp_path / "p.pptx"
    res = add_missing_slides(deck, out, [spec], RULES)
    assert _titles(out)[-1] == "Dubai Treaty (DBTRT) - Results" and "at the end" in res.entries[0]["status"]


def test_added_frame_slide_is_filled_like_the_model_and_verified(deck: Path, tmp_path: Path) -> None:
    out = tmp_path / "prepared.pptx"
    res = add_missing_slides(deck, out, SPECS[:1], RULES)
    dm = build_map(scan_deck(out, RULES), load_lobs(), load_ranges().life_resclasses, RULES, (), res.kinds_by_title)
    m = next(x for x in dm.mapped if x.lob == "DBPRO")
    assert [(s.kind, s.new) for s in m.slots] == [("table", True)]  # table only, like DBENR (not table + profit)
    jobs = [Job(m.slide, m.title, s.shape_id, s.kind, m.lob, _png(tmp_path / "t.png", (5, 6, 7), 8, 3), new=True, geometry=s.geometry) for s in m.slots]
    built = tmp_path / "built.pptx"
    write_deck(out, built, jobs)
    v = compare(snapshot(out), snapshot(built), {}, {}, {m.slide: 1})
    assert v.ok, v.messages
    # without the model's kinds an empty Results slide would get the standard pair (table + profit)
    dm2 = build_map(scan_deck(out, RULES), load_lobs(), load_ranges().life_resclasses, RULES)
    assert [s.kind for s in next(x for x in dm2.mapped if x.lob == "DBPRO").slots] == ["table", "profit"]


def test_selection_decides_whether_a_run_adds_slides() -> None:
    from resco.selection import Selection

    assert Selection.parse().allows_new("DBPRO", "results") and Selection.parse().allows_new("DBPRO", "netting")
    assert not Selection.parse("ENTOT").allows_new("DBPRO", "results")
    assert Selection.parse("DBPRO", "net").allows_new("DBPRO", "netting") and not Selection.parse("DBPRO", "net").allows_new("DBPRO", "results")
    assert not Selection.parse(slides="5").allows_new("DBPRO", "results")  # a slide filter never adds slides
    assert Selection.parse(parts="ibnr").allows_new("DBPRO", "results") is False


def test_shipped_config_is_valid() -> None:
    specs = load_slides_to_add()
    assert [s.id for s in specs] == ["dbpro_results", "dbpro_netting", "dbtrt_results", "dbtrt_netting"]
    assert {s.lob for s in specs} == {"DBPRO", "DBTRT"} and all(s.copy in ("frame", "full") for s in specs)
