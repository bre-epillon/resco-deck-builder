"""Selection of what a run touches: LoBs x content types x slides (src/resco/selection.py)."""
from __future__ import annotations

import pytest

from resco.config import TableTarget
from resco.deck_scan import DeckMap, MappedSlide, SlotEntry, TableMatch
from resco.selection import Selection, SelectionError, parse_parts, parse_slides, preview
from resco.serve import JobRunner


def slot(kind: str, i: int = 1) -> SlotEntry:
    return SlotEntry(kind, i, f"Picture {i}", (0.0, 0.0, 10.0, 10.0))


def slide(n: int, lob: str, typ: str, *kinds: str) -> MappedSlide:
    return MappedSlide(n, f"{lob} - {typ}", lob, f"Summary_{lob}", "", typ, typ, [slot(k, i) for i, k in enumerate(kinds, 1)])


def ladder(n: int, lob: str) -> TableMatch:
    t = TableTarget(f"gvn_{lob}", "ladder", f"{lob} gvn", ("gross",), {"lob": lob})
    return TableMatch(t, n, f"{lob} - Gross vs Net summary", 7, "Table 3", 10, 8)


def overview(n: int, tid: str) -> TableMatch:
    t = TableTarget(tid, "overview", "Earned IBNR allocation", ("uwy",), {}, ({"title": "x"},))
    return TableMatch(t, n, "Earned IBNR allocation", 9, "Table 1", 10, 8)


DECK = DeckMap(
    mapped=[
        slide(16, "ENTOT", "Results", "table", "profit"),
        slide(55, "ENTOT", "Results", "table", "profit"),
        slide(56, "ENTOT", "Triangles", "triangles", "chart_prem", "chart_paid", "chart_incurr"),
        slide(57, "ENUPS", "Results", "table", "profit"),
        slide(60, "ENUPS", "Expert Judgement", "ej"),
    ],
    narrative=[], flags=[], lobs_without_slides=[],
    table_matches=[ladder(17, "ENTOT"), ladder(19, "LITOT"), overview(27, "ibnr_alloc_totals"), overview(43, "ibnr_qoq_by_class")],
)  # fmt: skip


def test_parse_parts_groups_aliases_and_kinds() -> None:
    assert parse_parts("Net, results") == {"netting", "results"}
    assert parse_parts("expert-judgement; IBNR tables") == {"ej", "ibnr"}
    assert parse_parts("profit") == {"profit"}
    assert parse_parts("") == frozenset()
    with pytest.raises(SelectionError, match="unknown part"):
        parse_parts("results, bogus")


def test_parse_slides_ranges() -> None:
    assert parse_slides("16, 17-19 55") == {16, 17, 18, 19, 55}
    assert parse_slides("") == frozenset()
    for bad in ("0", "5-3", "x", "1-2-3"):
        with pytest.raises(SelectionError):
            parse_slides(bad)


def test_no_selection_keeps_everything() -> None:
    sel = Selection.parse()
    assert not sel.active
    assert sel.restrict_plan(DECK.mapped) == DECK.mapped
    assert len(sel.restrict_tables(DECK.table_matches, ["ENTOT", "LITOT"])) == 4  # all ladder tables + both IBNR tables


def test_only_entot_net_results_is_just_the_gross_vs_net_table() -> None:
    sel = Selection.parse("ENTOT", "net")
    assert sel.restrict_plan([m for m in DECK.mapped if m.lob == "ENTOT"]) == []  # no picture is exported
    got = sel.restrict_tables(DECK.table_matches, ["ENTOT"])
    assert [m.target.id for m in got] == ["gvn_ENTOT"]  # not LITOT's table, not the IBNR tables


def test_parts_filter_pictures_by_slide_type_and_kind() -> None:
    res = Selection.parse("ENTOT", "results").restrict_plan(DECK.mapped[:3])
    assert [(m.slide, [s.kind for s in m.slots]) for m in res] == [(16, ["table", "profit"]), (55, ["table", "profit"])]
    prof = Selection.parse("ENTOT", "profit").restrict_plan(DECK.mapped[:3])
    assert [(m.slide, [s.kind for s in m.slots]) for m in prof] == [(16, ["profit"]), (55, ["profit"])]


def test_slides_filter_applies_to_pictures_and_tables() -> None:
    sel = Selection.parse(slides="16-17")
    assert [m.slide for m in sel.restrict_plan(DECK.mapped)] == [16]
    assert [m.slide for m in sel.restrict_tables(DECK.table_matches, ["ENTOT", "LITOT"])] == [17]


def test_ibnr_tables_follow_the_lob_filter_unless_asked_for() -> None:
    assert Selection.parse("ENTOT").restrict_tables(DECK.table_matches, ["ENTOT"]) == [DECK.table_matches[0]]  # LoB-limited run: no IBNR tables
    explicit = Selection.parse("ENTOT", "ibnr").restrict_tables(DECK.table_matches, ["ENTOT"])
    assert [m.target.id for m in explicit] == ["ibnr_alloc_totals", "ibnr_qoq_by_class"]


def test_describe_and_as_dict() -> None:
    assert Selection.parse().describe() == "everything in the template"
    sel = Selection.parse("entot,ENUPS", "net,results", "16-19,55")
    assert sel.lobs == ("ENTOT", "ENUPS")
    text = sel.describe()
    assert "LoBs ENTOT, ENUPS" in text and "Gross vs Net" in text and "slides 16-19, 55" in text
    assert sel.as_dict()["slides"] == [16, 17, 18, 19, 55]


def test_preview_lists_the_slides_that_pass_the_filters() -> None:
    p = preview(DECK, Selection.parse("ENTOT", "results,net", "16,17"))
    rows = {r["slide"]: r for r in p["rows"]}
    assert sorted(rows) == [16, 17, 55]  # triangles slide 56 is filtered out by content type, LITOT by LoB
    assert [rows[s]["selected"] for s in (16, 17, 55)] == [True, True, False]
    assert (p["n_slides"], p["n_pictures"], p["n_tables"]) == (2, 2, 1)
    assert p["needs_excel"] is True
    net_only = preview(DECK, Selection.parse("ENTOT", "net"))
    assert net_only["needs_excel"] is False and net_only["n_tables"] == 1
    assert preview(DECK, Selection.parse("NOPE"))["unknown_lobs"] == ["NOPE"]


def test_server_validates_the_selection() -> None:
    base = {"mode": "dry-run", "deck": "main", "lobs": "ENTOT", "template": __file__, "workbook": __file__}
    assert JobRunner._validate({**base, "parts": "bogus"}) is not None
    assert "not allowed" in JobRunner._validate({**base, "slides": "1; rm -rf"})
    assert "slide range" in JobRunner._validate({**base, "slides": "5-3"})
    assert "template deck not found" in JobRunner._validate({**base, "parts": "net", "slides": "16-17"})  # selection is fine, the fake template is not
