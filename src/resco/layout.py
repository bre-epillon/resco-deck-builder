"""Where things are in THIS version of the workbook (found by labels, see locate.py) versus where the config says they
should be. Everything the build needs from the sheets comes from here, so a moved block is followed instead of failing."""
from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

from resco.config import Ranges
from resco.excel_reader import ExcelSession
from resco.locate import EjBlock, locate_ej_blocks, locate_selector_blocks
from resco.values import col_letters, decode_grid, letters_to_col

log = logging.getLogger(__name__)


@dataclass
class Layout:
    table: str
    table_short: str
    profit: str
    triangles: str
    ej: dict[str, EjBlock] = field(default_factory=dict)
    ej_block_cols: tuple[tuple[str, str], ...] = ()  # (name, "E:O") shifted to where the LOB column actually is
    notes: list[tuple[str, str]] = field(default_factory=list)  # (level, message)
    configured: dict[str, str] = field(default_factory=dict)

    def as_report(self) -> dict[str, Any]:
        items = {k: {"address": getattr(self, k), "configured": self.configured.get(k, ""), "moved": getattr(self, k) != self.configured.get(k, "")}
                 for k in ("table", "table_short", "profit", "triangles")}  # fmt: skip
        return {
            **items,
            "expert_judgement": {code: {"rows": [b.first, b.last], "block": [b.run_first, b.run_last]} for code, b in self.ej.items()},
            "ej_columns": [c for _, c in self.ej_block_cols],
            "notes": [list(n) for n in self.notes],
        }


def from_config(ranges: Ranges) -> Layout:
    """The configured addresses as they are (tests, and the fallback when nothing can be located)."""
    cfg = {"table": ranges.table, "table_short": ranges.table_life, "profit": ranges.profit, "triangles": ranges.triangles}
    return Layout(cfg["table"], cfg["table_short"], cfg["profit"], cfg["triangles"], {}, tuple(ranges.ej_blocks), [], cfg)


def discover_layout(xl: ExcelSession, ranges: Ranges, max_year: int) -> Layout:
    """The selector must already show a valid LoB (labels are the same for every LoB)."""
    lay = from_config(ranges)
    cfg = dict(lay.configured)

    grid, _ = decode_grid(xl.read_sheet_values(ranges.selector_sheet))
    found = locate_selector_blocks(grid)
    for problem in found.problems:
        lay.notes.append(("WARN", f"{ranges.selector_sheet}: {problem}; the configured address is used"))
    uses_short = bool(ranges.life_resclasses)  # only decks with Life classes show the short table
    for name, label in (("table", "Results table"), ("table_short", "Results table (Life: first block only)"), ("profit", "P&L box"), ("triangles", "Triangles")):
        if name == "table_short" and not uses_short:
            continue
        got = getattr(found, name)
        if got is None:
            lay.notes.append(("WARN", f"{label}: could not be located by its labels; using the configured {cfg[name]}"))
        elif got == cfg[name]:
            lay.notes.append(("INFO", f"{label}: {got} (as configured)"))
        else:
            lay.notes.append(("WARN", f"{label}: moved. Configured {cfg[name]}, found {got}: the found range is used"))
            setattr(lay, name, got)

    # Expert Judgement: blocks by LoB code, columns relative to the 'LOB' header
    ej_grid, _ = decode_grid(xl.read_sheet_values(ranges.ej_sheet, max_cols=60))
    header = next(((r, c) for r, row in enumerate(ej_grid[:12], 1) for c, v in enumerate(row, 1) if isinstance(v, str) and v.strip().upper() == "LOB"), None)
    if header is None:
        lay.notes.append(("ERROR", f"{ranges.ej_sheet}: no 'LOB' header cell in the first rows: Expert Judgement blocks cannot be located"))
        return lay
    lob_col = header[1]
    delta = lob_col - letters_to_col(ranges.ej_blocks[0][1].split(":")[0])
    if delta:
        lay.notes.append(("WARN", f"{ranges.ej_sheet}: the 'LOB' column is {col_letters(lob_col)} (configured {ranges.ej_blocks[0][1].split(':')[0]}): column ranges shifted by {delta}"))
        lay.ej_block_cols = tuple((n, f"{col_letters(letters_to_col(a) + delta)}:{col_letters(letters_to_col(b) + delta)}") for n, cols in ranges.ej_blocks for a, b in [cols.split(":")])
    uwy_header = ej_grid[header[0] - 1][lob_col] if lob_col < len(ej_grid[0]) else None
    if not (isinstance(uwy_header, str) and uwy_header.strip().upper() == "UWY"):
        lay.notes.append(("WARN", f"{ranges.ej_sheet}: the column right of 'LOB' is not 'UWY' (found {uwy_header!r}); check the sheet"))
    col_lob = [row[lob_col - 1] for row in ej_grid]
    col_uwy = [row[lob_col] if lob_col < len(row) else None for row in ej_grid]
    ej = locate_ej_blocks(col_lob, col_uwy, max_year)
    lay.ej = ej.blocks
    lay.notes += [("WARN", p) for p in ej.problems]
    lay.notes.append(("INFO", f"Expert Judgement: {len(ej.blocks)} LoB blocks found by code ({', '.join(ej.blocks)})"))
    for note in lay.notes:
        log.log(logging.WARNING if note[0] != "INFO" else logging.INFO, "layout: %s", note[1])
    return lay
