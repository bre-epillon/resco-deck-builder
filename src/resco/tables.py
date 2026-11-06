"""Fill the native tables of narrative slides from Excel figures (netting ladder, IBNR allocation 'Overview').

The tables already exist in the deck with its styling, so only their TEXT is replaced (the run formatting of each cell is
kept). Rows are matched by their UWY label, columns by their header text, so a table with fewer UWY rows (the UK deck)
or an extra column is handled; anything that does not match is reported, never guessed. Pure python (no Excel here):
the Excel side hands over ready-made text blocks."""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl

from resco.locate import BlockLoc, norm, norm_uwy

def norm_title(s: str) -> str:
    return norm(s)


# ----------------------------------------------------------------------------------------------- netting ladder
LADDER_KEYS = ("lob", "level", "basis", "branch")


class LadderError(RuntimeError):
    pass


@dataclass
class LadderBlock:
    rows: dict[str, dict[tuple[str, str], float | None]]  # uwy label (normalised) -> {(group, header): value}
    selection: dict[str, str]


class Ladder:
    """The netting workbook's flat 'Ladder' sheet: LoB, Level, Basis, Branch, UWY, then column groups (Gross, ...)."""

    def __init__(self, path: Path, sheet: str = "Ladder") -> None:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb[sheet] if sheet in wb.sheetnames else wb[wb.sheetnames[0]]
            grid = [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
        hr = next((i for i, row in enumerate(grid[:30]) if {"lob", "uwy"} <= {norm(v) for v in row}), None)
        if hr is None:
            raise LadderError("no header row with 'LoB' and 'UWY' found in the netting sheet")
        head = grid[hr]
        self.cols = {norm(v): j for j, v in enumerate(head) if norm(v) in {"lob", "level", "basis", "branch", "uwy"}}
        missing = {"lob", "level", "basis", "branch", "uwy"} - set(self.cols)
        if missing:
            raise LadderError(f"netting sheet: header column(s) missing: {sorted(missing)}")
        group_row = grid[hr - 1] if hr > 0 else []
        self.keys: dict[tuple[str, str], int] = {}
        group = ""
        for j in range(len(head)):
            g = norm(group_row[j]) if j < len(group_row) else ""
            group = g or group
            h = norm(head[j])
            if h and j > self.cols["uwy"]:
                self.keys[(group, h)] = j
        self.rows = [r for r in grid[hr + 1 :] if r and r[self.cols["lob"]] not in (None, "")]
        self.groups = sorted({g for g, _ in self.keys})

    def block(self, select: dict[str, str]) -> LadderBlock:
        want = {k: norm(v) for k, v in select.items() if k in LADDER_KEYS}
        hits = [r for r in self.rows if all(norm(r[self.cols[k]]) == v for k, v in want.items())]
        if not hits:
            raise LadderError(f"no ladder rows for {select}")
        combos = {tuple(norm(r[self.cols[k]]) for k in LADDER_KEYS) for r in hits}
        if len(combos) > 1:  # e.g. the same LoB at two levels with different numbers: refuse to pick silently
            raise LadderError(f"{select} matches {len(combos)} different blocks {sorted(combos)}: add the missing keys (level/basis/branch)")
        out: dict[str, dict[tuple[str, str], float | None]] = {}
        for r in hits:
            out[norm_uwy(r[self.cols["uwy"]])] = {k: (float(r[j]) if isinstance(r[j], (int, float)) and not isinstance(r[j], bool) else None) for k, j in self.keys.items()}
        return LadderBlock(out, dict(select))


RATIO_HEADERS = {"gulr", "nulr", "nnulr"}
DELTA = norm("Δ")  # lower-cased by norm(): the ladder headers are compared in this form


def format_ladder_value(header: str, v: float | None) -> str:
    """Amounts in millions with one decimal, ratios as whole percent; no minus sign on a rounded zero."""
    if v is None:
        return ""
    header = norm(header)  # norm() lower-cases the delta sign too: compare in that form
    if header in RATIO_HEADERS or header.startswith(DELTA):
        s = f"{v * 100:.0f}%"
        return "0%" if s in ("-0%",) else s
    s = f"{v / 1e6:.1f}"
    return "0.0" if s == "-0.0" else s


# ----------------------------------------------------------------------------------------------- slide tables
@dataclass
class FillResult:
    changed: list[tuple[int, int, str, str]] = field(default_factory=list)  # row, col, old, new
    unchanged: int = 0
    unmatched_rows: list[str] = field(default_factory=list)  # slide row labels without source data
    unmatched_cols: list[str] = field(default_factory=list)  # slide headers without source column
    unused_source_rows: list[str] = field(default_factory=list)
    format_drift: list[tuple[int, int, str, str]] = field(default_factory=list)  # new text has another shape than the old
    notes: list[str] = field(default_factory=list)
    overflow: list[tuple[int, int, str, float]] = field(default_factory=list)  # row, col, text, column width pt: may not fit
    old_signs: list[str] = field(default_factory=list)  # sign styles of the template's figures
    new_signs: list[str] = field(default_factory=list)  # ... and of the new ones
    grid: list[list[list[str]]] = field(default_factory=list)  # [row][col] -> [old, new] for the dashboard


def cell_text(table: Any, r: int, c: int) -> str:
    return table.cell(r, c).text.strip()


def table_signature(table: Any, rows: int = 2) -> str:
    """Lower-case text of the first rows, used to recognise which table on a slide is the target."""
    return " | ".join(norm(table.cell(r, c).text) for r in range(min(rows, len(table.rows))) for c in range(len(table.columns)))


def find_header_row(table: Any) -> int | None:
    for r in range(min(len(table.rows), 4)):
        if any(norm_uwy(cell_text(table, r, c)) == "uwy" for c in range(len(table.columns))):
            return r
    return None


def uwy_columns(table: Any, hr: int) -> list[int]:
    return [c for c in range(len(table.columns)) if norm_uwy(cell_text(table, hr, c)) == "uwy"]


def group_labels(table: Any, hr: int) -> list[str]:
    """Group header (merged cell in the row above the headers) for each column: the last non-empty text to the left."""
    out, cur = [], ""
    for c in range(len(table.columns)):
        t = norm(cell_text(table, hr - 1, c)) if hr > 0 else ""
        cur = t or cur
        out.append(cur)
    return out


def data_row_labels(table: Any, hr: int, label_col: int) -> dict[str, int]:
    rows: dict[str, int] = {}
    for r in range(hr + 1, len(table.rows)):
        lab = norm_uwy(cell_text(table, r, label_col))
        if lab:
            rows[lab] = r
    return rows


def shape_of(text: str) -> str:
    """Visual shape of a figure without its sign: digits become 9, so '9.9' and '9%' differ, '-0.1', '+0.1' and '(0.1)' do not."""
    return re.sub(r"\d+", "9", re.sub(r"[()+\-\u2212]", "", text.strip()))


def sign_style(text: str) -> str | None:
    """How a signed figure is written: '(0.1)', '-0.1' or '+0.1'; None for an unsigned one."""
    t = text.strip()
    if t.startswith("("):
        return "parentheses"
    if t[:1] in ("-", "\u2212"):
        return "minus"
    if t.startswith("+"):
        return "plus"
    return None


def set_cell_text(cell: Any, text: str) -> None:
    """Replace the text of a table cell keeping the first run's formatting (font, size, colour)."""
    from pptx.oxml.ns import qn

    tf = cell.text_frame
    para = tf.paragraphs[0]
    runs = para.runs
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        run = para.add_run()
        run.text = text
        end = para._p.find(qn("a:endParaRPr"))
        if end is not None:  # an empty cell keeps its formatting in endParaRPr: carry it over to the new run
            rpr = run._r.get_or_add_rPr()
            for k, v in end.attrib.items():
                rpr.set(k, v)
            for child in end:
                rpr.append(copy.deepcopy(child))
    for extra_p in tf.paragraphs[1:]:  # a figure never needs more than one paragraph
        extra_p._p.getparent().remove(extra_p._p)


AVG_CHAR_EM = 0.55  # average width of a figure character as a fraction of the font size (digits are about 0.5 em)


def _fit_problem(table: Any, r: int, c: int, text: str) -> float | None:
    """Column width in points when `text` is probably too wide for its cell (estimated from font size and margins)."""
    cell = table.cell(r, c)
    run = next((x for x in cell.text_frame.paragraphs[0].runs), None)
    size = run.font.size.pt if run is not None and run.font.size is not None else None
    if size is None or cell.is_spanned:
        return None
    width = table.columns[c].width / 12700
    if cell.is_merge_origin:
        width = sum(table.columns[k].width for k in range(c, c + cell.span_width)) / 12700
    avail = width - (cell.margin_left + cell.margin_right) / 12700
    return width if len(text) * size * AVG_CHAR_EM > avail else None


def _apply(table: Any, res: FillResult, r: int, c: int, new: str) -> None:
    old = cell_text(table, r, c)
    res.grid[r][c] = [old, new]
    for text, bucket in ((old, res.old_signs), (new, res.new_signs)):
        style = sign_style(text)
        if style:
            bucket.append(style)
    if old == new:
        res.unchanged += 1
        return
    if old and shape_of(old) != shape_of(new):
        res.format_drift.append((r, c, old, new))
    set_cell_text(table.cell(r, c), new)
    res.changed.append((r, c, old, new))
    fit = _fit_problem(table, r, c, new)
    if fit:
        res.overflow.append((r, c, new, fit))


def _blank_grid(table: Any) -> list[list[list[str]]]:
    return [[[cell_text(table, r, c), cell_text(table, r, c)] for c in range(len(table.columns))] for r in range(len(table.rows))]


def fill_ladder(table: Any, block: LadderBlock, ladder_groups: list[str]) -> FillResult:
    """Gross / Net before / Net after columns: header row 'Premium Ultimate ... GULR | ... NULR QoQ | ... NNULR QoQ'."""
    res = FillResult(grid=_blank_grid(table))
    hr = find_header_row(table)
    if hr is None:
        res.notes.append("no 'UWY' header row in the slide table")
        return res
    ucols = uwy_columns(table, hr)
    ucol = ucols[0]
    groups = group_labels(table, hr)
    rows = data_row_labels(table, hr, ucol)
    sample = next(iter(block.rows.values()), {})
    group_map = {}
    for g in set(groups):
        if g:
            group_map[g] = next((lg for lg in ladder_groups if lg == g), next((lg for lg in ladder_groups if lg.startswith(g) or g.startswith(lg)), None))
    col_keys: dict[int, tuple[tuple[str, str], str]] = {}
    last_header: dict[str, str] = {}
    for c in range(ucol + 1, len(table.columns)):
        h = norm(cell_text(table, hr, c))
        if not h:
            continue
        g = groups[c]
        if h == "qoq":  # QoQ under a ratio column = the ladder's 'Δ <ratio>' column of the same group
            h = norm("Δ " + last_header.get(g, ""))
        else:
            last_header[g] = h
        lg = group_map.get(g)
        key = (lg, h) if lg is not None else None
        if key is None or key not in sample:
            res.unmatched_cols.append(f"{g or '?'} / {h}")
            continue
        col_keys[c] = (key, h)
    for lab, r in rows.items():
        src = block.rows.get(lab)
        if src is None:
            res.unmatched_rows.append(lab)
            continue
        for c, (key, h) in col_keys.items():
            _apply(table, res, r, c, format_ladder_value(h, src.get(key)))
    res.unused_source_rows = [k for k in block.rows if k not in rows]
    res.notes += ladder_identity_notes(block, ladder_groups)
    return res


def ladder_identity_notes(block: LadderBlock, groups: list[str], tol: float = 0.005) -> list[str]:
    """Source consistency: ratio = claims ultimate / premium ultimate in every row and group (also catches a mislabelled column)."""
    notes: list[str] = []
    for uwy, vals in block.rows.items():
        for g in groups:
            prem, ult = vals.get((g, "premium ultimate")), vals.get((g, "claims ultimate"))
            ratio_key = next((k for k in vals if k[0] == g and k[1].endswith("ulr") and not k[1].startswith(DELTA)), None)
            ratio = vals.get(ratio_key) if ratio_key else None
            if prem and ult is not None and ratio is not None and abs(prem) > 1 and abs(ult / prem - ratio) > tol:
                notes.append(f"netting ladder, UWY {uwy}, {g}: {ratio_key[1].upper()} {ratio:.1%} differs from claims ultimate / premium ultimate = {ult / prem:.1%}")
    return notes[:6]


@dataclass
class BlockTexts:
    """A block of the IBNR 'Overview' sheet as displayed text: header cells and one text row per UWY label."""

    header: list[str]  # header texts per block column (normalised)
    rows: dict[str, list[str]]  # uwy label (normalised) -> displayed texts per block column
    loc: BlockLoc | None = None
    checks: list[tuple[str, str]] = field(default_factory=list)  # (address, value) for 'checks' cells that are not TRUE


def _match_header(h: str, candidates: list[str], taken: set[int]) -> int | None:
    for i, cand in enumerate(candidates):
        if i not in taken and cand == h:
            return i
    for i, cand in enumerate(candidates):
        if i not in taken and cand and len(h) >= 4 and (cand.startswith(h) or h.startswith(cand)):
            return i
    return None


def fill_blocks(table: Any, blocks: list[BlockTexts | None]) -> FillResult:
    """One block per UWY column segment of the slide table (left to right): rows by UWY label, columns by header text."""
    res = FillResult(grid=_blank_grid(table))
    hr = find_header_row(table)
    if hr is None:
        res.notes.append("no 'UWY' header row in the slide table")
        return res
    ucols = uwy_columns(table, hr)
    if len(ucols) != len(blocks):
        res.notes.append(f"the slide table has {len(ucols)} UWY column group(s) but {len(blocks)} source block(s) are configured")
    bounds = ucols + [len(table.columns)]
    for k, (start, stop) in enumerate(zip(bounds, bounds[1:])):
        block = blocks[k] if k < len(blocks) else None
        if block is None:
            res.notes.append(f"segment {k + 1}: source block not found in the workbook")
            continue
        rows = data_row_labels(table, hr, start)
        taken: set[int] = set()
        col_map: dict[int, int] = {}
        for c in range(start, stop):
            h = norm(cell_text(table, hr, c))
            if not h:
                continue
            i = _match_header(h, block.header, taken)
            if i is None:
                res.unmatched_cols.append(h)
                continue
            taken.add(i)
            col_map[c] = i
        for lab, r in rows.items():
            src = block.rows.get(lab)
            if src is None:
                res.unmatched_rows.append(f"{lab} (segment {k + 1})")
                continue
            for c, i in col_map.items():
                if c == start:  # the UWY label column itself is not a figure
                    continue
                _apply(table, res, r, c, src[i] if i < len(src) else "")
        res.unused_source_rows += [k2 for k2 in block.rows if k2 not in rows]
    return res
