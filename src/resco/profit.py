"""'Profit (loss) movement QoQ' boxes under a Gross vs Net summary table.

Each box is the change in result between Q-1 and Q on the ladder's Total row:
(premium ultimate Q - premium ultimate Q-1) - (claims ultimate Q - claims ultimate Q-1), once for the "Net before internal
QS" block and once for the "Net after internal QS" block (the formula of the netting workbook's Summary_Tot). Q comes from
the netting file, Q-1 from the Q-1 netting file (for the UK branch: the figures as presented last quarter).

The first build inserts the two boxes as small native tables under their blocks, styled like the Excel pictures of the
main deck; from then on they are found by their title and only the figure is updated. Pure python-pptx, no Excel."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

from resco.locate import norm
from resco.tables import FillResult, LadderBlock, cell_text, find_header_row, group_labels, set_cell_text

TITLE = "Profit (loss) movement QoQ"
BASIS_LABEL, ROW_LABEL = "MBE", "A+L"
BLOCKS = (("net before", "Net before Internal QS"), ("net after", "Net after Internal QS"))  # ladder group prefix, label
DEFAULT_HEADER = {"net before": RGBColor(0x21, 0x5C, 0x98), "net after": RGBColor(0x4D, 0x93, 0xD9)}  # colours of the Excel boxes

BOX_W, BOX_H, GAP = Inches(1.75), Inches(0.57), Inches(0.08)  # size of the main deck's boxes, gap below the ladder table
FONT = Pt(8)
RED, GREEN, BLACK, WHITE, GREY = (RGBColor(0xC0, 0x00, 0x00), RGBColor(0x3C, 0x7D, 0x22), RGBColor(0, 0, 0),
                                  RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0xBF, 0xBF, 0xBF))  # fmt: skip


class ProfitError(RuntimeError):
    pass


# ----------------------------------------------------------------------------------------------- figures
def ladder_group(groups: list[str], prefix: str) -> str | None:
    """The ladder's own name of a block ('net before internal qs'), found by its first words."""
    return next((g for g in groups if g.startswith(prefix)), None)


def movement(cur: LadderBlock, prior: LadderBlock | None, group: str | None) -> float | None:
    """Profit (loss) movement of one block on the Total row; None when either quarter lacks a figure."""
    if prior is None or group is None:
        return None
    a, b = cur.rows.get("total"), prior.rows.get("total")
    if not a or not b:
        return None
    vals = [a.get((group, "premium ultimate")), b.get((group, "premium ultimate")), a.get((group, "claims ultimate")), b.get((group, "claims ultimate"))]
    if any(v is None for v in vals):
        return None
    return (vals[0] - vals[1]) - (vals[2] - vals[3])


def format_movement(v: float | None) -> str:
    """Millions with one decimal, like the Excel box; no minus sign on a rounded zero."""
    if v is None:
        return ""
    s = f"{v / 1e6:.1f}"
    return "0.0" if s == "-0.0" else s


def colour_of(text: str) -> RGBColor:
    """A loss is red, a profit green, a rounded zero black (the Excel box's conditional colours)."""
    if not text or text == "0.0":
        return BLACK
    return RED if text.startswith("-") else GREEN


# ----------------------------------------------------------------------------------------------- deck side
@dataclass
class BoxSpec:
    label: str  # 'Net before Internal QS'
    left: int  # EMU
    top: int
    header: RGBColor


@dataclass
class ProfitFill:
    result: FillResult
    filled_ids: list[int] = field(default_factory=list)  # existing boxes whose figure was written
    inserted_ids: list[int] = field(default_factory=list)  # boxes added to the slide


def find_boxes(slide: Any) -> list[Any]:
    """The profit boxes already on the slide (tables whose first cell holds the title), left to right."""
    out = []
    for sh in slide.shapes:
        if getattr(sh, "has_table", False) and sh.has_table and norm(cell_text(sh.table, 0, 0)).startswith(norm(TITLE)):
            out.append(sh)
    return sorted(out, key=lambda s: s.left)


def _cell_fill(cell: Any) -> RGBColor | None:
    tcPr = cell._tc.tcPr
    clr = tcPr.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}") if tcPr is not None else None
    return RGBColor.from_string(clr.get("val")) if clr is not None else None


def box_specs(anchor: Any) -> list[BoxSpec]:
    """Where the boxes go: right-aligned under the last column (QoQ) of each block of the Gross vs Net table, just below
    it; the header takes the colour of the block's band in that table."""
    table = anchor.table
    hr = find_header_row(table)
    if hr is None or hr == 0:
        raise ProfitError("the Gross vs Net table has no band row above a 'UWY' header row")
    groups = group_labels(table, hr)
    edges, x = [], anchor.left
    for col in table.columns:
        x += col.width
        edges.append(x)
    specs = []
    for prefix, label in BLOCKS:
        cols = [c for c, g in enumerate(groups) if g.startswith(prefix)]
        if not cols:
            raise ProfitError(f"the Gross vs Net table has no '{label}' block")
        header = _cell_fill(table.cell(hr - 1, cols[0])) or DEFAULT_HEADER[prefix]
        specs.append(BoxSpec(label, edges[cols[-1]] - BOX_W, anchor.top + anchor.height + GAP, header))
    return specs


def _line(tag: str, visible: bool) -> etree._Element:
    ln = etree.Element(qn(tag), w=str(Pt(1) if visible else 0), cap="flat", cmpd="sng", algn="ctr")
    if visible:
        fill = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(fill, qn("a:srgbClr"), val="000000")
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
    else:
        etree.SubElement(ln, qn("a:noFill"))
    return ln


def _style_cell(cell: Any, fill: RGBColor, lines: tuple[bool, bool, bool, bool]) -> None:
    """Borders (left, right, top, bottom), fill and margins, written in the order the schema wants."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        tcPr.remove(child)
    for tag, visible in zip(("a:lnL", "a:lnR", "a:lnT", "a:lnB"), lines):
        tcPr.append(_line(tag, visible))
    solid = etree.SubElement(tcPr, qn("a:solidFill"))
    etree.SubElement(solid, qn("a:srgbClr"), val=str(fill))
    tcPr.set("marL", str(Inches(0.03)))
    tcPr.set("marR", str(Inches(0.03)))
    tcPr.set("marT", "0")
    tcPr.set("marB", "0")
    tcPr.set("anchor", "ctr")


def _write(cell: Any, text: str, bold: bool, colour: RGBColor) -> None:
    para = cell.text_frame.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.runs[0] if para.runs else para.add_run()
    run.text = text
    run.font.size, run.font.bold = FONT, bold
    run.font.color.rgb = colour


def insert_box(slide: Any, spec: BoxSpec, value: str, lob: str) -> Any:
    """A 3 x 2 table: title over both columns, then 'MBE' over the figure, then 'A+L' and the figure."""
    frame = slide.shapes.add_table(3, 2, Emu(spec.left), Emu(spec.top), BOX_W, BOX_H)
    frame.name = f"ResCo profit QoQ {spec.label}"
    frame._element.nvGraphicFramePr.cNvPr.set("descr", f"{lob} {TITLE} ({spec.label}), inserted by the ResCo deck builder")
    table = frame.table
    tbl_pr = table._tbl.tblPr
    for flag in ("firstRow", "bandRow"):
        tbl_pr.set(flag, "0")
    table.columns[0].width = Emu(int(BOX_W * 0.45))
    table.columns[1].width = Emu(BOX_W - table.columns[0].width)
    for row in table.rows:
        row.height = Emu(BOX_H // 3)
    table.cell(0, 0).merge(table.cell(0, 1))
    for r in range(3):
        for c in range(2):  # outer frame and a line under each row; no line between the two columns
            _style_cell(table.cell(r, c), spec.header if r == 0 else (GREY if (r, c) == (2, 1) and not value else WHITE), (c == 0, c == 1, True, True))
            # an empty cell would keep PowerPoint's default 18 pt end-of-paragraph size and make its row taller
            table.cell(r, c).text_frame.paragraphs[0]._p.get_or_add_endParaRPr().set("sz", str(int(FONT.pt * 100)))
    _write(table.cell(0, 0), TITLE, True, WHITE)
    _write(table.cell(1, 1), BASIS_LABEL, True, BLACK)
    _write(table.cell(2, 0), ROW_LABEL, True, BLACK)
    _write(table.cell(2, 1), value, False, colour_of(value))
    return frame


def write_value(box: Any, value: str) -> str:
    """Update the figure of an existing box (bottom-right cell), keeping its font; returns the old text."""
    table = box.table
    cell = table.cell(len(table.rows) - 1, len(table.columns) - 1)
    old = cell.text.strip()
    if old != value:
        set_cell_text(cell, value)
    for run in cell.text_frame.paragraphs[0].runs:
        run.font.color.rgb = colour_of(value)
    if not value or _cell_fill(cell) == GREY:  # an empty figure is greyed out, as in the Excel box
        cell.fill.solid()
        cell.fill.fore_color.rgb = GREY if not value else WHITE
    return old


def fill_profit(slide: Any, anchor: Any, cur: LadderBlock, prior: LadderBlock | None, groups: list[str], lob: str) -> ProfitFill:
    values = [format_movement(movement(cur, prior, ladder_group(groups, prefix))) for prefix, _ in BLOCKS]
    labels = [label for _, label in BLOCKS]
    res = FillResult(grid=[[[label, label], ["", v]] for label, v in zip(labels, values)])
    out = ProfitFill(res)
    for label, v in zip(labels, values):
        if not v:
            res.notes.append(f"{label}: no {'Q-1 ' if prior is not None else ''}figures for the Total row, the box is left empty")
    boxes = find_boxes(slide)
    if not boxes:
        for spec, v in zip(box_specs(anchor), values):
            frame = insert_box(slide, spec, v, lob)
            out.inserted_ids.append(frame.shape_id)
            res.changed.append((len(res.changed), 1, "", v))
        return out
    if len(boxes) != len(BLOCKS):
        raise ProfitError(f"{len(boxes)} '{TITLE}' table(s) on the slide, expected {len(BLOCKS)} (or none, then they are inserted)")
    for i, (box, v) in enumerate(zip(boxes, values)):
        old = write_value(box, v)
        res.grid[i][1] = [old, v]
        if old == v:
            res.unchanged += 1
        else:
            res.changed.append((i, 1, old, v))
        out.filled_ids.append(box.shape_id)
    return out
