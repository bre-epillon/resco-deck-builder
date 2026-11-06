"""Step 2: loop over LoBs, dump every block to CSV (audit trail) and a manifest. No PowerPoint here."""
from __future__ import annotations

import csv
import logging
import time
from dataclasses import dataclass, field
from pathlib import Path

from resco.config import Lob, Ranges
from resco.excel_reader import ExcelSession, SelectorError
from resco.layout import Layout, from_config
from resco.values import cell_address, decode_grid, letters_to_col, split_address as _split

log = logging.getLogger(__name__)


@dataclass
class BlockResult:
    lob: str
    block: str
    sheet: str
    address: str
    rows: int
    cols: int
    non_empty: int
    n_errors: int


@dataclass
class DumpResult:
    out_dir: Path
    blocks: list[BlockResult] = field(default_factory=list)
    errors: list[tuple[str, str, str, str, bool]] = field(default_factory=list)  # lob, block, cell, text, hidden
    failures: list[tuple[str, str]] = field(default_factory=list)  # lob, message
    ej_checks: list[tuple[str, str]] = field(default_factory=list)  # lob, message
    seconds_per_lob: dict[str, float] = field(default_factory=dict)


def _write_grid(path: Path, grid: list[list]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(grid)


def _record(
    res: DumpResult,
    lob: str,
    block: str,
    sheet: str,
    address: str,
    grid: list[list],
    path: Path,
    hidden: tuple[set[int], set[int]] = (set(), set()),
) -> list[list]:
    first_row, first_col = _split(address)
    clean, errs = decode_grid(grid)
    _write_grid(path, clean)
    non_empty = sum(1 for row in clean for v in row if v not in (None, ""))
    res.blocks.append(BlockResult(lob, block, sheet, address, len(clean), len(clean[0]) if clean else 0, non_empty, len(errs)))
    for r, c, text in errs:
        is_hidden = (first_row + r) in hidden[0] or (first_col + c) in hidden[1]
        res.errors.append((lob, block, cell_address(first_row, first_col, r, c), text, is_hidden))
    return clean


def dump_all(session: ExcelSession, lobs: list[Lob], ranges: Ranges, out_dir: Path, max_year: int, layout: Layout | None = None) -> DumpResult:
    layout = layout or from_config(ranges)
    res = DumpResult(out_dir=out_dir)
    data = out_dir / "data"

    for lob in lobs:
        t0 = time.time()
        code = lob.code
        try:
            session.select(ranges.selector_sheet, ranges.selector_cell, ranges.check_cell, lob.selector, lob.expected_check)
        except SelectorError as exc:
            log.error("%s: %s", code, exc)
            res.failures.append((code, str(exc)))
            continue

        # Dump the superset of what the legacy macro exports (table C4:S30 covers the Life variant C4:S17).
        for block, address in (("table", layout.table), ("profit", layout.profit), ("triangles", layout.triangles)):
            grid = session.read_grid(ranges.selector_sheet, address)
            hidden = session.hidden_rows_cols(ranges.selector_sheet, address)
            _record(res, code, block, ranges.selector_sheet, address, grid, data / code / f"{block}.csv", hidden)

        ej_block = layout.ej.get(lob.resclass) if lob.resclass else None
        if ej_block is not None:
            r0, r1 = ej_block.first, ej_block.last
            if r1 > session.used_last_row(ranges.ej_sheet):
                msg = f"{code}: ExpertJudgement rows {r0}-{r1} lie beyond the sheet used range"
                log.warning(msg)
                res.ej_checks.append((code, msg))
            else:
                for i, (name, cols) in enumerate(layout.ej_block_cols):
                    c0, c1 = cols.split(":")
                    address = f"{c0}{r0}:{c1}{r1}"
                    grid = session.read_grid(ranges.ej_sheet, address)
                    clean = _record(res, code, f"ej_{name}", ranges.ej_sheet, address, grid, data / code / f"ej_{name.replace(' ', '_')}.csv")
                    if i == 0:  # first block starts in column E = LoB code (legacy column layout)
                        codes = {row[0] for row in clean}
                        if codes != {lob.resclass}:
                            msg = f"{code}: ExpertJudgement rows {r0}-{r1} column {c0} holds {sorted(map(str, codes))}, expected ['{lob.resclass}']"
                            log.warning(msg)
                            res.ej_checks.append((code, msg))
        res.seconds_per_lob[code] = round(time.time() - t0, 1)
        log.info("%s done in %.1fs", code, res.seconds_per_lob[code])
    return res


def write_manifest(res: DumpResult) -> None:
    out = res.out_dir
    with (out / "manifest.csv").open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["lob", "block", "sheet", "range", "rows", "cols", "non_empty_cells", "error_cells"])
        for b in res.blocks:
            w.writerow([b.lob, b.block, b.sheet, b.address, b.rows, b.cols, b.non_empty, b.n_errors])
    with (out / "error_cells.csv").open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["lob", "block", "cell", "error", "in_hidden_row_or_col"])
        w.writerows(res.errors)
