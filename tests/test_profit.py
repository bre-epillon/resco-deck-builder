"""Profit (loss) movement QoQ boxes under a Gross vs Net table: figures, insertion, update in place, verification (no real data)."""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from resco.config import TableConfig, TableTarget
from resco.deck_scan import TableMatch
from resco.deck_writer import TableJob, write_deck
from resco.profit import BOX_W, GREEN, GREY, RED, find_boxes, format_movement, movement
from resco.selection import Selection
from resco.table_plan import load_sources, make_jobs, report
from resco.tables import Ladder
from resco.verify import compare, snapshot
from test_tables import gvn_table, ladder_file, ladder_row

SEL = {"lob": "ENTOT", "level": "Aggregate", "basis": "MBE", "branch": "UK branch"}
TITLE = "Energy (ENTOT) - Gross vs Net summary - MBE"
BAND = {"gross": "A50021", "net before": "215F9A", "net after": "4E95D9"}


def ladder_for(tmp_path: Path, name: str, total_base: float) -> Ladder:
    rows = [ladder_row("ENTOT", "Aggregate", "MBE", "UK branch", u, base=b) for u, b in (("≤2019", 1.0), ("2020", 2.0), ("Total", total_base))]
    sub = tmp_path / name
    sub.mkdir()
    return Ladder(ladder_file(sub, rows))


def gvn_deck(tmp_path: Path) -> Path:
    """A one-slide deck: title text box plus the Gross vs Net table, its band cells filled like the real deck."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(0.9), Inches(0.4), Inches(11), Inches(0.5)).text_frame.text = TITLE
    model = gvn_table()  # built on its own slide: copy its XML into ours
    frame = slide.shapes.add_table(6, 19, Inches(1.2), Inches(1.4), Inches(11.3), Inches(2.4))
    frame._element.remove(frame._element.graphic)
    frame._element.append(model._graphic_frame._element.graphic)
    for col, key in ((2, "gross"), (7, "net before"), (13, "net after")):
        cell = frame.table.cell(0, col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(BAND[key])
    path = tmp_path / "template.pptx"
    prs.save(path)
    return path


def job_for(deck: Path, cur: Ladder, prior: Ladder | None) -> TableJob:
    slide = Presentation(str(deck)).slides[0]
    anchor = next(s for s in slide.shapes if s.has_table and "gross" in s.table.cell(0, 2).text.lower())
    return TableJob(1, TITLE, anchor.shape_id, "pq_ENTOT", "profit_qoq", block=cur.block(SEL), ladder_groups=cur.groups,
                    prior_block=prior.block(SEL) if prior else None, lob="ENTOT")  # fmt: skip


def test_movement_is_premium_change_minus_claims_change_on_the_total_row(tmp_path: Path):
    cur, prior = ladder_for(tmp_path, "q", 9.0), ladder_for(tmp_path, "q1", 8.0)
    before = movement(cur.block(SEL), prior.block(SEL), "net before internal qs")  # (18 - 16) - (20.7 - 18.4) = -0.3 m
    after = movement(cur.block(SEL), prior.block(SEL), "net after internal qs")  # (27 - 24) - (20.7 - 18.4) = +0.7 m
    assert round(before / 1e6, 6) == -0.3 and round(after / 1e6, 6) == 0.7
    assert movement(cur.block(SEL), None, "net before internal qs") is None
    assert format_movement(-0.3e6) == "-0.3" and format_movement(-0.04e6) == "0.0" and format_movement(None) == ""


def test_boxes_are_inserted_under_their_blocks_then_updated_in_place(tmp_path: Path):
    deck = gvn_deck(tmp_path)
    cur, prior = ladder_for(tmp_path, "q", 9.0), ladder_for(tmp_path, "q1", 8.0)
    out = tmp_path / "built.pptx"
    job = job_for(deck, cur, prior)
    write_deck(deck, out, [], [job])
    assert job.error is None and len(job.inserted_ids) == 2 and job.box_ids == []

    slide = Presentation(str(out)).slides[0]
    anchor = next(s for s in slide.shapes if s.shape_id == job.shape_id)
    boxes = find_boxes(slide)
    assert [b.table.cell(2, 1).text for b in boxes] == ["-0.3", "0.7"]
    assert boxes[0].table.cell(2, 1).text_frame.paragraphs[0].runs[0].font.color.rgb == RED
    assert boxes[1].table.cell(2, 1).text_frame.paragraphs[0].runs[0].font.color.rgb == GREEN
    assert [b.table.cell(r, c).text for b in boxes[:1] for r, c in ((0, 0), (1, 1), (2, 0))] == ["Profit (loss) movement QoQ", "MBE", "A+L"]
    # header colour = the band of its block; right edge = right edge of the block's last column; just below the table
    assert [str(b.table.cell(0, 0).fill.fore_color.rgb) for b in boxes] == [BAND["net before"], BAND["net after"]]
    widths = [c.width for c in anchor.table.columns]
    edge = lambda last: anchor.left + sum(widths[: last + 1])  # noqa: E731
    assert [b.left + b.width for b in boxes] == [edge(12), edge(18)] and all(b.width == BOX_W for b in boxes)
    assert all(b.top > anchor.top + anchor.height for b in boxes)

    v = compare(snapshot(deck), snapshot(out), {}, {}, {}, {1: 2})
    assert v.ok, v.messages
    assert not compare(snapshot(deck), snapshot(out), {}, {}, {}).ok  # not announced: refused

    # next quarter the boxes are part of the template: found by their title and only the figure changes
    out2 = tmp_path / "built2.pptx"
    job2 = job_for(out, ladder_for(tmp_path, "q_next", 10.0), cur)
    write_deck(out, out2, [], [job2])
    assert job2.error is None and job2.inserted_ids == [] and len(job2.box_ids) == 2
    boxes2 = find_boxes(Presentation(str(out2)).slides[0])
    assert [b.table.cell(2, 1).text for b in boxes2] == ["-0.3", "0.7"]  # same step between the synthetic quarters
    assert len(Presentation(str(out2)).slides[0].shapes) == len(Presentation(str(out)).slides[0].shapes)
    v2 = compare(snapshot(out), snapshot(out2), {}, {1: set(job2.box_ids)})
    assert v2.ok, v2.messages


def test_missing_q1_figures_leave_an_empty_grey_box_and_a_note(tmp_path: Path):
    deck = gvn_deck(tmp_path)
    job = job_for(deck, ladder_for(tmp_path, "q", 9.0), None)
    out = tmp_path / "built.pptx"
    write_deck(deck, out, [], [job])
    boxes = find_boxes(Presentation(str(out)).slides[0])
    assert [b.table.cell(2, 1).text for b in boxes] == ["", ""] and boxes[0].table.cell(2, 1).fill.fore_color.rgb == GREY
    assert len(job.result.notes) == 2


def test_sources_jobs_and_report(tmp_path: Path):
    target = TableTarget("pq_ENTOT", "profit_qoq", TITLE, ("gross", "net before"), dict(SEL))
    match = TableMatch(target, 1, TITLE, 4, "Table 3", 6, 19)
    cur_path = ladder_file(tmp_path, [ladder_row("ENTOT", "Aggregate", "MBE", "UK branch", "Total")])
    # without a Q-1 file the boxes are skipped with a warning, never an error
    jobs = make_jobs([match], load_sources([match], TableConfig(), cur_path, None, None, None))
    entries, issues, filled = report(jobs, [match])
    assert jobs[0].warn_only and entries[0]["status"].startswith("skipped") and [i.level for i in issues] == ["WARN"] and filled == {}
    # with it, both quarters are read
    src = load_sources([match], TableConfig(), cur_path, None, None, cur_path)
    assert src.prior is not None and src.info["netting_prior"] == cur_path.name
    assert make_jobs([match], src)[0].prior_block is not None


def test_the_netting_part_selects_the_boxes_too():
    sel = Selection.parse(None, "net", None)
    assert sel.table_kinds() == {"ladder", "profit_qoq"}
    target = TableTarget("pq_ENTOT", "profit_qoq", TITLE, ("gross",), dict(SEL))
    m = TableMatch(target, 14, TITLE, 4, "Table 2", 8, 18)
    assert sel.restrict_tables([m], ["ENTOT"]) == [m] and sel.restrict_tables([m], ["PRTOT"]) == []
    assert Selection.parse(None, "ibnr", None).restrict_tables([m], ["ENTOT"]) == []
