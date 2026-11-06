"""Find things in the workbook by what they ARE (labels, LoB codes), not by fixed addresses.

The reserving workbook changes between versions: blocks get re-ordered or removed (v2.13b dropped ENOTH from the
Expert Judgement sheet and every block below moved up by 10 rows), a UWY row is added every year. These functions
work on a plain 2-D grid of cell values (1-based rows/columns in the results), so they are testable without Excel.
Configured addresses stay as the fallback and as the reference for "this moved" reports.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from resco.values import col_letters


def _text(v: Any) -> str | None:
    return v.strip() if isinstance(v, str) and v.strip() else None


def _empty(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def address(r1: int, c1: int, r2: int, c2: int) -> str:
    return f"{col_letters(c1)}{r1}:{col_letters(c2)}{r2}"


def find_text(
    grid: list[list[Any]], text: str, *, mode: str = "exact", row_from: int = 1, row_to: int | None = None,
    col_from: int = 1, col_to: int | None = None,
) -> tuple[int, int] | None:  # fmt: skip
    """First cell (row-major) whose text matches; returns 1-based (row, col)."""
    want = text.lower()
    for r in range(row_from - 1, min(row_to or len(grid), len(grid))):
        row = grid[r]
        for c in range(col_from - 1, min(col_to or len(row), len(row))):
            t = _text(row[c])
            if t is None:
                continue
            t = t.lower()
            if (mode == "exact" and t == want) or (mode == "startswith" and t.startswith(want)) or (mode == "contains" and want in t):
                return r + 1, c + 1
    return None


# ----------------------------------------------------------------------------------------------- selector sheet
@dataclass
class SheetBlocks:
    table: str | None = None  # whole Results table (all UWY blocks)
    table_short: str | None = None  # first block only (Life classes show just Attritional + Large)
    profit: str | None = None
    triangles: str | None = None
    problems: list[str] = field(default_factory=list)


def locate_selector_blocks(grid: list[list[Any]]) -> SheetBlocks:
    out = SheetBlocks()
    n_rows = len(grid)
    n_cols = len(grid[0]) if grid else 0

    # triangles first: its top row also bounds the table
    tri_top = find_text(grid, "Year 1", row_to=min(n_rows, 120))
    # Results table: the header cell 'UWY'; one row above it holds 'Gross', one column left the vertical block labels
    uwy = find_text(grid, "UWY", row_to=min(n_rows, 40))
    if uwy is None:
        out.problems.append("results table: no 'UWY' header cell found")
    else:
        hr, hc = uwy
        right = hc
        while right < n_cols and not _empty(grid[hr - 1][right]):  # contiguous header cells
            right += 1
        # helper columns (e.g. 'LL alloc weight', 'Last exp date') can sit right next to the table: it ends at the last GULR header
        gulr = [c for c in range(hc, right + 1) if "gulr" in (_text(grid[hr - 1][c - 1]) or "").lower()]
        if gulr:
            right = gulr[-1]
        limit = (tri_top[0] - 1) if tri_top and tri_top[0] > hr else min(n_rows, hr + 60)
        totals = [r for r in range(hr + 1, limit + 1) if (_text(grid[r - 1][hc - 1]) or "").lower() == "total"]
        if not totals:
            out.problems.append("results table: no 'Total' row under the UWY column")
        else:
            top, left = max(hr - 1, 1), max(hc - 1, 1)
            out.table = address(top, left, totals[-1], right)
            out.table_short = address(top, left, totals[0], right)

    # P&L box: its title cell, then title + 'MBE/Stat' header + 'A+L' + 'A' rows, three columns
    pl = find_text(grid, "profit (loss) movement", mode="startswith", row_to=min(n_rows, 100))
    while pl is not None and "qoq" not in (_text(grid[pl[0] - 1][pl[1] - 1]) or "").lower():  # the QoQ box, not YTD
        nxt = find_text(grid, "profit (loss) movement", mode="startswith", row_from=pl[0] + 1, row_to=min(n_rows, 100))
        if nxt is None:
            pl = None
            break
        pl = nxt
    if pl is None:
        out.problems.append("P&L box: no 'Profit (loss) movement QoQ' title found")
    else:
        out.profit = address(pl[0], pl[1], pl[0] + 3, pl[1] + 2)

    # triangles: 'Year 1' header row .. last row of the 'ClmI/GPrmB' block, from the block-label column to '>=60'
    if tri_top is None:
        out.problems.append("triangles: no 'Year 1' header row found")
    else:
        first = find_text(grid, "GPrmB", row_from=tri_top[0], row_to=tri_top[0] + 3)
        if first is None:
            out.problems.append("triangles: no 'GPrmB' block label under the 'Year 1' row")
        else:
            gr, gc = first
            end = find_text(grid, ">=60", row_from=gr, row_to=gr)
            last_block = find_text(grid, "ClmI/GPrmB", row_from=gr + 1, row_to=min(n_rows, gr + 80), col_from=gc, col_to=gc)
            if end is None or last_block is None:
                out.problems.append("triangles: no '>=60' column header or no 'ClmI/GPrmB' block found")
            else:
                right_col, r = end[1], last_block[0]
                while r + 1 <= n_rows and any(not _empty(v) for v in grid[r][max(gc - 2, 0) : right_col]):  # one column left too: hidden rows may be labelled there
                    r += 1
                out.triangles = address(tri_top[0], gc, r, right_col)
    return out


# ----------------------------------------------------------------------------------------------- ExpertJudgement sheet
@dataclass(frozen=True)
class EjBlock:
    first: int  # first data row used (UWY 2017 ...)
    last: int  # last data row used (UWY <= max_year)
    run_first: int  # extent of the whole block of this LoB code in the sheet
    run_last: int

    @property
    def n_rows(self) -> int:
        return self.last - self.first + 1


@dataclass
class EjLayout:
    blocks: dict[str, EjBlock] = field(default_factory=dict)
    problems: list[str] = field(default_factory=list)


def locate_ej_blocks(col_lob: list[Any], col_uwy: list[Any], max_year: int, header: str = "LOB") -> EjLayout:
    """Blocks are runs of identical LoB codes in the 'LOB' column; rows are limited to UWY <= max_year.
    `col_lob` / `col_uwy` are the cell values of that column / the UWY column from row 1 downwards."""
    out = EjLayout()
    runs: list[tuple[str, int, int]] = []
    r = 0
    while r < len(col_lob):
        code = _text(col_lob[r])
        if code is None or code.lower() == header.lower():
            r += 1
            continue
        start = r
        while r + 1 < len(col_lob) and _text(col_lob[r + 1]) == code:
            r += 1
        runs.append((code, start + 1, r + 1))
        r += 1
    for code, first, last in runs:
        if code in out.blocks:
            out.problems.append(f"Expert Judgement: LoB '{code}' appears in more than one block (rows {out.blocks[code].run_first} and {first}); the first is used")
            continue
        used_last = first
        for rr in range(first, last + 1):
            uwy = col_uwy[rr - 1] if rr - 1 < len(col_uwy) else None
            if isinstance(uwy, (int, float)) and not isinstance(uwy, bool) and uwy > max_year:
                break
            used_last = rr
        out.blocks[code] = EjBlock(first, used_last, first, last)
    if not out.blocks:
        out.problems.append(f"Expert Judgement: no LoB blocks found under the '{header}' column")
    return out


# ----------------------------------------------------------------------------------------------- blocks on the IBNR 'Overview' sheet
_DASHES = str.maketrans({"\u2013": "-", "\u2014": "-", "\u2212": "-", "\u00a0": " "})


def norm(s: Any) -> str:
    """Lower case, dashes unified, whitespace collapsed; '≤' -> '<='."""
    if s is None:
        return ""
    t = str(s).translate(_DASHES).replace("\u2264", "<=").replace("\n", " ").replace("\x0b", " ")
    return " ".join(t.split()).lower()


def norm_uwy(s: Any) -> str:
    """UWY label as text: 2020.0 -> '2020', '≤2019' -> '<=2019', 'Total' -> 'total'."""
    if isinstance(s, float) and s == int(s):
        s = int(s)
    t = norm(s)
    return t[:-2] if t.endswith(".0") else t


@dataclass
class BlockLoc:
    header_row: int  # 1-based sheet rows / columns
    label_col: int
    first_col: int
    last_col: int
    data_rows: list[tuple[int, str]] = field(default_factory=list)  # (row, normalised UWY label), up to and including 'Total'
    checks_col: int | None = None


def find_block(grid: list[list[Any]], find: dict[str, Any]) -> BlockLoc | None:
    """Locate a table block by its labels. `find` is either
    {title: 'Gross Earned Statutory', below: 'UWY', occurrence: n}: a title cell with 'UWY' directly under it, or
    {header: ['UWY', 'GEP'], occurrence: n}: a header row starting with these labels.
    Matches are ordered left to right, so occurrence 1 is the Lux block and occurrence 2 the UK block further right."""
    n_rows = len(grid)
    cands: list[tuple[int, int, int]] = []  # (col, row of the title/header, header row), all 0-based
    for r in range(n_rows):
        row = grid[r]
        for c in range(len(row)):
            t = norm(row[c])
            if not t:
                continue
            if "title" in find:
                below = norm_uwy(find.get("below", "UWY"))
                if t.startswith(norm(find["title"])) and r + 1 < n_rows and c < len(grid[r + 1]) and norm_uwy(grid[r + 1][c]) == below:
                    cands.append((c, r, r + 1))
            else:
                toks = [norm_uwy(x) for x in find["header"]]
                if norm_uwy(row[c]) != toks[0]:
                    continue
                j, ok = c + 1, True
                for tok in toks[1:]:
                    while j < len(row) and not norm(row[j]) and j - c < 12:
                        j += 1
                    if j >= len(row) or not norm(row[j]).startswith(tok):
                        ok = False
                        break
                    j += 1
                if ok:
                    cands.append((c, r, r))
    cands.sort()
    occurrence = int(find.get("occurrence", 1))
    if len(cands) < occurrence:
        return None
    lc, _, hr = cands[occurrence - 1]
    row = grid[hr]
    last, empties, checks = lc, 0, None
    for j in range(lc + 1, len(row)):
        t = norm(row[j])
        if t == "checks":
            checks = j
            break
        if t and norm_uwy(row[j]) == "uwy":  # the next block starts here
            break
        if not t:
            empties += 1
            if empties >= 2:
                break
            continue
        empties, last = 0, j
    data: list[tuple[int, str]] = []
    blanks = 0
    for r in range(hr + 1, min(n_rows, hr + 60)):
        label = norm_uwy(grid[r][lc]) if lc < len(grid[r]) else ""
        if not label:
            blanks += 1
            if data and blanks > 2:
                break
            continue
        blanks = 0
        data.append((r + 1, label))
        if label == "total":
            break
    return BlockLoc(hr + 1, lc + 1, lc + 1, last + 1, data, None if checks is None else checks + 1)
