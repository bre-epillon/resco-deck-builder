"""Data checks on the grids behind the images (CLAUDE.md validations 3 and 7). Pure: no Excel, no PowerPoint."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from resco.values import cell_address, decode_grid, split_address

REL_TOL = 1e-6


@dataclass
class TotalsResult:
    status: str  # pass | fail | n/a
    comparisons: int = 0
    max_rel_diff: float = 0.0
    failures: list[str] = field(default_factory=list)


def _num(v: Any) -> float | None:
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0  # blank behaves like 0 in a sum
    return None  # text / error value: cannot take part in the arithmetic


def _rel(a: float, b: float) -> float:
    return abs(a - b) / max(abs(a), abs(b), 1.0)


def check_table_totals(grid: list[list[Any]], label_col: int | None = None, rel_tol: float = REL_TOL) -> TotalsResult:
    """Results-table check. Rows are labelled in `label_col` (UWY): detail rows roll into a `<=YYYY` row and everything
    ends in a `Total` row. Every additive column (header without 'GULR') must satisfy
    `<=YYYY row == sum of the detail rows above it` and `Total == <=YYYY row + the later years`."""
    found = next(((i, j) for i, r in enumerate(grid) for j, v in enumerate(r) if v == "UWY" and (label_col in (None, j))), None)
    if found is None:
        return TotalsResult("n/a", failures=["no 'UWY' header cell found"])
    header_idx, label_col = found
    header = grid[header_idx]
    cols = [
        c for c in range(label_col + 1, len(header))
        if isinstance(header[c], str) and header[c].strip() and "GULR" not in header[c].upper()
    ]  # fmt: skip

    res = TotalsResult("pass")

    def compare(what: str, a: float | None, b: float | None, row_label: str, col: int) -> None:
        if a is None or b is None:
            return
        res.comparisons += 1
        d = _rel(a, b)
        res.max_rel_diff = max(res.max_rel_diff, d)
        if d > rel_tol:
            res.failures.append(f"{what} ({row_label}, column '{header[col]}' #{col}): {a:,.2f} vs {b:,.2f}")

    segment: list[int] = []
    for i in range(header_idx + 1, len(grid)):
        label = grid[i][label_col]
        if label == "Total":
            body = [j for j in segment if isinstance(grid[j][label_col], (str, int, float)) and grid[j][label_col] != ""]
            aggs = [j for j in body if isinstance(grid[j][label_col], str) and grid[j][label_col].startswith("<=")]
            for c in cols:
                total = _num(grid[i][c])
                if len(aggs) > 1:
                    agg = aggs[-1]
                    before = [j for j in body if j < agg]
                    after = [j for j in body if j > agg]
                    parts = [_num(grid[j][c]) for j in before]
                    if None not in parts:
                        compare(f"'{grid[agg][label_col]}' vs sum of earlier rows", _num(grid[agg][c]), sum(parts), f"row {i + 1}", c)  # type: ignore[arg-type]
                    parts = [_num(grid[agg][c])] + [_num(grid[j][c]) for j in after]
                    if None not in parts:
                        compare("Total vs rolled-up rows", total, sum(parts), f"row {i + 1}", c)  # type: ignore[arg-type]
                else:
                    parts = [_num(grid[j][c]) for j in body]
                    if None not in parts:
                        compare("Total vs sum of rows", total, sum(parts), f"row {i + 1}", c)  # type: ignore[arg-type]
            segment = []
        elif label in (None, ""):
            continue  # spare / separator row (UK sheets keep blank rows inside a block): not part of any sum
        else:
            segment.append(i)

    if res.comparisons == 0:
        res.status = "n/a"
        res.failures.append("no comparable columns")
    elif res.failures:
        res.status = "fail"
    return res


def visible_error_cells(
    grid: list[list[Any]], address: str, hidden_rows: set[int], hidden_cols: set[int]
) -> list[tuple[str, str]]:
    """Error values (cell, text) that sit in rows/columns that are NOT hidden, i.e. that would show in a picture."""
    first_row, first_col = split_address(address)
    _, errs = decode_grid(grid)
    return [
        (cell_address(first_row, first_col, r, c), text)
        for r, c, text in errs
        if (first_row + r) not in hidden_rows and (first_col + c) not in hidden_cols
    ]
