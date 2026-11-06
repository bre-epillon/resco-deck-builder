"""Excel side of the table filling: read blocks of the IBNR allocation 'Overview' sheet as the text Excel DISPLAYS
(so '+0.6', '(0.1)', '65%' arrive exactly as in the workbook). Blocks are found by their titles, not by addresses."""
from __future__ import annotations

import logging
from typing import Any

import pywintypes

from resco.excel_reader import ExcelSession
from resco.locate import find_block, norm, norm_uwy
from resco.tables import BlockTexts
from resco.values import cell_address, col_letters, decode_grid

log = logging.getLogger(__name__)
MIN_READABLE_WIDTH = 16  # a figure shown as '#####' needs a wider column


def _block_texts(ws: Any, loc: Any, grid: list[list[Any]]) -> BlockTexts:
    cols = range(loc.first_col, loc.last_col + 1)
    header = [norm(grid[loc.header_row - 1][c - 1]) for c in cols]

    def read() -> dict[str, list[str]]:
        return {label: [str(ws.api.Cells(r, c).Text) for c in cols] for r, label in loc.data_rows}

    rows = read()
    if any("###" in t for texts in rows.values() for t in texts):
        # too narrow to display the figure: widen the block's columns IN MEMORY (the workbook copy is never saved) and read again
        for c in cols:
            if ws.api.Columns(c).ColumnWidth < MIN_READABLE_WIDTH:
                ws.api.Columns(c).ColumnWidth = MIN_READABLE_WIDTH
        rows = read()
    checks: list[tuple[str, str]] = []
    if loc.checks_col:
        for r, label in loc.data_rows:
            v = ws.api.Cells(r, loc.checks_col).Value2
            if v is False or (isinstance(v, str) and v.strip().upper() == "FALSE"):
                checks.append((f"{col_letters(loc.checks_col)}{r}", "FALSE"))
    return BlockTexts(header, rows, loc, checks)


def read_blocks(book: Any, sheet: str, finds: list[dict[str, Any]], max_rows: int = 300, max_cols: int = 140) -> tuple[list[BlockTexts | None], list[str]]:
    """One BlockTexts (or None when the labels are not found) per find rule, plus human-readable notes."""
    notes: list[str] = []
    try:
        ws = book.sheets[sheet]
    except Exception:  # noqa: BLE001 - xlwings raises its own error for a missing sheet
        return [None] * len(finds), [f"sheet '{sheet}' not found in the workbook"]
    last = ws.used_range.last_cell
    rows, cols = min(last.row, max_rows), min(last.column, max_cols)
    raw = ws.api.Range(ws.api.Cells(1, 1), ws.api.Cells(rows, cols)).Value2
    grid, _ = decode_grid([list(r) for r in raw])
    out: list[BlockTexts | None] = []
    for find in finds:
        loc = find_block(grid, find)
        label = find.get("title") or " ".join(find.get("header", []))
        if loc is None:
            notes.append(f"block '{label}' (occurrence {find.get('occurrence', 1)}) not found on '{sheet}'")
            out.append(None)
            continue
        out.append(_block_texts(ws, loc, grid))
        a = f"{col_letters(loc.first_col)}{loc.header_row}:{col_letters(loc.last_col)}{loc.data_rows[-1][0] if loc.data_rows else loc.header_row}"
        log.info("overview block '%s' #%s found at %s (%d UWY rows)", label, find.get("occurrence", 1), a, len(loc.data_rows))
    return out, notes
