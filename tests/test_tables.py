"""Filling native slide tables from Excel figures: synthetic tables and a synthetic ladder (no real data)."""
from pathlib import Path

import openpyxl
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

from resco.locate import find_block
from resco.tables import (
    BlockTexts, Ladder, LadderError, fill_blocks, fill_ladder, format_ladder_value, norm_uwy, set_cell_text,
)


def make_table(rows: int, cols: int, texts: dict[tuple[int, int], str], size_pt: float = 9):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(rows, cols, Inches(0.2), Inches(0.2), Inches(9), Inches(3)).table
    for r in range(rows):
        for c in range(cols):
            tf = table.cell(r, c).text_frame
            run = tf.paragraphs[0].add_run()
            run.text = texts.get((r, c), "")
            run.font.size = Pt(size_pt)
    return table


def ladder_file(tmp_path: Path, rows: list[tuple]) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ladder"
    ws["A1"] = "Reserve netting ladder"
    ws["F4"], ws["L4"], ws["R4"] = "Gross", "Net before internal QS", "Net after internal QS"
    head = ["LoB", "Level", "Basis", "Branch", "UWY"]
    for ratio in ("GULR", "NULR", "NNULR"):
        head += ["Premium ultimate", "Incurred claims", "SIBNR + IBNR", "Claims ultimate", ratio, "Δ " + ratio]
    for j, h in enumerate(head, 1):
        ws.cell(5, j, h)
    for i, r in enumerate(rows, 6):
        for j, v in enumerate(r, 1):
            ws.cell(i, j, v)
    path = tmp_path / "ladder.xlsx"
    wb.save(path)
    return path


def ladder_row(lob, level, basis, branch, uwy, base=1.0):
    vals = []
    for g in range(3):
        vals += [base * 1e6 * (g + 1), base * 2e6, base * 3e5, base * 2.3e6, 0.76 - 0.1 * g, 0.004 * (g + 1) * (-1 if g == 2 else 1)]
    return (lob, level, basis, branch, uwy, *vals)


def gvn_table():
    texts = {(0, 2): "Gross", (0, 7): "Net before Internal QS", (0, 13): "Net after Internal QS"}
    heads = ["UWY", "Premium Ultimate", "Incurred Claims", "SIBNR + IBNR", "Claims Ultimate", "GULR",
             "Premium Ultimate", "Incurred Claims", "SIBNR + IBNR", "Claims Ultimate", "NULR", "QoQ",
             "Premium Ultimate", "Incurred Claims", "SIBNR + IBNR", "Claims Ultimate", "NNULR", "QoQ"]
    texts.update({(1, 1 + j): h for j, h in enumerate(heads)})
    for r, lab in ((2, "<=2019"), (3, "2020"), (4, "2021"), (5, "Total")):
        texts[(r, 1)] = lab
        for c in range(2, 19):
            texts[(r, c)] = "0.0"
    t = make_table(6, 19, texts)
    t.cell(0, 2).merge(t.cell(0, 6))
    t.cell(0, 7).merge(t.cell(0, 12))
    t.cell(0, 13).merge(t.cell(0, 18))
    return t


def test_format_rules():
    assert format_ladder_value("premium ultimate", 107_712_345.0) == "107.7"
    assert format_ladder_value("premium ultimate", -49_000.0) == "-0.0" or format_ladder_value("premium ultimate", -49_000.0) == "0.0"
    assert format_ladder_value("gulr", 0.7649) == "76%" and format_ladder_value("Δ nulr", -0.004) == "0%"
    assert format_ladder_value("Δ nnulr", -0.017) == "-2%"
    assert format_ladder_value("claims ultimate", None) == ""
    assert norm_uwy("≤2019") == "<=2019" and norm_uwy(2020.0) == "2020" and norm_uwy("Total") == "total"


def test_ladder_fill_matches_rows_and_columns_and_keeps_formatting(tmp_path: Path):
    path = ladder_file(tmp_path, [ladder_row("ENTOT", "Aggregate", "MBE", "Total", u, base=b) for u, b in (("≤2019", 1.0), ("2020", 2.0), ("Total", 9.0))])
    ladder = Ladder(path)
    block = ladder.block({"lob": "ENTOT", "level": "Aggregate", "basis": "MBE", "branch": "Total"})
    table = gvn_table()
    res = fill_ladder(table, block, ladder.groups)
    assert table.cell(2, 2).text == "1.0" and table.cell(2, 3).text == "2.0" and table.cell(2, 6).text == "76%"  # <=2019: gross premium, incurred, GULR
    assert table.cell(5, 2).text == "9.0"  # Total row
    assert table.cell(2, 7).text == "2.0" and table.cell(2, 11).text == "66%"  # net before: premium, NULR
    assert table.cell(2, 12).text == "1%" or table.cell(2, 12).text == "0%"  # QoQ = the ladder's delta NULR column
    assert res.unmatched_rows == ["2021"] and table.cell(4, 2).text == "0.0"  # no ladder row for 2021: left untouched, reported
    assert res.unmatched_cols == []
    assert table.cell(2, 2).text_frame.paragraphs[0].runs[0].font.size == Pt(9)  # formatting survived the text replacement


def test_ladder_selection_is_unambiguous(tmp_path: Path):
    rows = [ladder_row("DCTOT", "Aggregate", "MBE", "Total", "Total"), ladder_row("DCTOT", "LoB", "MBE", "Total", "Total", base=5.0)]
    ladder = Ladder(ladder_file(tmp_path, rows))
    with pytest.raises(LadderError, match="2 different blocks"):
        ladder.block({"lob": "DCTOT", "basis": "MBE", "branch": "Total"})  # the same LoB at two levels: must not pick silently
    assert ladder.block({"lob": "DCTOT", "level": "LoB", "basis": "MBE", "branch": "Total"}).rows["total"][("gross", "premium ultimate")] == 5e6
    with pytest.raises(LadderError, match="no ladder rows"):
        ladder.block({"lob": "XXX"})


def test_overview_blocks_fill_by_header_and_uwy_label():
    texts = {(0, 0): "UWY", (0, 1): "GEP", (0, 2): "Gross Incurred Claims", (0, 3): "", (0, 4): "Movement QoQ Gross Booked IBNR"}
    for r, lab in ((1, "<=2019"), (2, "2020"), (3, "Total")):
        texts[(r, 0)] = lab
        for c in range(1, 5):
            texts[(r, c)] = "+0.6" if c == 4 else "1.0"
    table = make_table(4, 5, texts)
    block = BlockTexts(
        header=["uwy", "gep", "gross incurred claims", "movement qoq gross booked"],
        rows={"<=2019": ["<=2019", "279.0", "171.7", "-9.9"], "2020": ["2020", "126.3", "11.9", "0.6"], "total": ["Total", "1118.4", "519.0", "28.1"]},
    )
    res = fill_blocks(table, [block])
    assert [table.cell(1, c).text for c in range(5)] == ["<=2019", "279.0", "171.7", "1.0", "-9.9"]  # the empty-header column keeps its old text
    assert table.cell(3, 1).text == "1118.4"
    assert res.format_drift == []  # '+0.6' -> '0.6' is a sign change, not another number format
    assert set(res.new_signs) - set(res.old_signs) == {"minus"}  # ... but the template never wrote a '-' sign in this table: reported at table level
    assert res.unmatched_cols == []


def test_two_segments_gross_and_net_blocks():
    texts = {(0, 0): "Gross Earned Statutory IBNR", (0, 4): "Net Earned Statutory IBNR"}
    heads = ["UWY", "Energy", "Life", "", "UWY", "Energy", "Life"]
    texts.update({(1, c): h for c, h in enumerate(heads)})
    for r, lab in ((2, "<=2019"), (3, "Total")):
        texts[(2 if lab == "<=2019" else 3, 0)] = lab
        texts[(2 if lab == "<=2019" else 3, 4)] = lab
    table = make_table(4, 7, texts)
    gross = BlockTexts(["uwy", "energy", "life"], {"<=2019": ["<=2019", "0.1", "0.3"], "total": ["Total", "114.6", "54.8"]})
    net = BlockTexts(["uwy", "energy", "life"], {"<=2019": ["<=2019", "0.0", "0.0"], "total": ["Total", "6.7", "5.5"]})
    res = fill_blocks(table, [gross, net])
    assert table.cell(3, 1).text == "114.6" and table.cell(3, 5).text == "6.7" and table.cell(3, 6).text == "5.5"
    assert res.unmatched_rows == []
    # a missing source block is reported, the other segment is still filled
    table2 = make_table(4, 7, texts)
    res2 = fill_blocks(table2, [gross, None])
    assert table2.cell(3, 1).text == "114.6" and table2.cell(3, 5).text == "" and any("not found" in n for n in res2.notes)


def test_empty_cell_takes_formatting_from_the_paragraph_end_properties():
    from pptx.oxml.ns import qn

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    table = slide.shapes.add_table(1, 1, Inches(1), Inches(1), Inches(2), Inches(1)).table
    para = table.cell(0, 0).text_frame.paragraphs[0]
    end = para._p.makeelement(qn("a:endParaRPr"), {"sz": "900"})
    para._p.append(end)
    set_cell_text(table.cell(0, 0), "12.3")
    run = table.cell(0, 0).text_frame.paragraphs[0].runs[0]
    assert run.text == "12.3" and run._r.find(qn("a:rPr")).get("sz") == "900"


def test_find_block_by_title_and_by_header_with_occurrences():
    g = [[None] * 30 for _ in range(30)]
    for (r, c, v) in ((4, 3, "Gross Earned Statutory IBNR"), (5, 3, "UWY"), (5, 4, "Energy"), (5, 5, "Total"), (5, 6, "checks"), (5, 7, "UWY"), (5, 8, "Energy"),
                      (4, 20, "Gross Earned Statutory IBNR"), (5, 20, "UWY"), (5, 21, "Energy"), (5, 22, "Total")):
        g[r - 1][c - 1] = v
    for r, lab in ((6, "<=2019"), (7, 2020.0), (8, "Total"), (9, "excl 2026")):
        g[r - 1][2] = lab
        g[r - 1][19] = lab
    lux = find_block(g, {"title": "Gross Earned Statutory", "below": "UWY", "occurrence": 1})
    uk = find_block(g, {"title": "Gross Earned Statutory", "below": "UWY", "occurrence": 2})
    assert (lux.header_row, lux.label_col, lux.first_col, lux.last_col, lux.checks_col) == (5, 3, 3, 5, 6)  # stops at 'checks', not at the next block
    assert [lab for _, lab in lux.data_rows] == ["<=2019", "2020", "total"]  # up to and including Total
    assert uk.label_col == 20 and uk.last_col == 22
    assert find_block(g, {"title": "Net Earned Statutory", "below": "UWY"}) is None
    assert find_block(g, {"header": ["UWY", "Energy"], "occurrence": 3}).label_col == 20  # the third match, counted left to right


def test_a_different_number_shape_is_a_format_drift():
    table = make_table(2, 2, {(0, 0): "UWY", (0, 1): "GULR", (1, 0): "Total", (1, 1): "1.0"})
    block = BlockTexts(["uwy", "gulr"], {"total": ["Total", "65%"]})
    res = fill_blocks(table, [block])
    assert table.cell(1, 1).text == "65%" and [(o, n) for _, _, o, n in res.format_drift] == [("1.0", "65%")]


def test_a_figure_that_is_too_wide_for_its_column_is_flagged():
    table = make_table(2, 2, {(0, 0): "UWY", (0, 1): "GEP", (1, 0): "Total", (1, 1): "1.0"}, size_pt=12)
    table.columns[1].width = Pt(20)  # 20 pt column, 12 pt font: '1118.4' cannot fit
    block = BlockTexts(["uwy", "gep"], {"total": ["Total", "1118.4"]})
    res = fill_blocks(table, [block])
    assert [(r, c, t) for r, c, t, _ in res.overflow] == [(1, 1, "1118.4")]
    roomy = make_table(2, 2, {(0, 0): "UWY", (0, 1): "GEP", (1, 0): "Total", (1, 1): "1.0"}, size_pt=9)
    assert fill_blocks(roomy, [block]).overflow == []  # a normal 1.5 inch column holds it
