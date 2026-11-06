"""Pure helpers for cell values coming back from Excel."""
from __future__ import annotations

import datetime as dt
from typing import Any

# Excel error codes as returned by COM (VT_ERROR surfaced as a signed int).
EXCEL_ERRORS = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
    -2146826245: "#GETTING_DATA",
}


def decode_cell(value: Any) -> tuple[Any, str | None]:
    """Return (clean value, error text or None). Error cells keep their text so nothing is silently dropped."""
    if isinstance(value, int) and not isinstance(value, bool) and value in EXCEL_ERRORS:
        err = EXCEL_ERRORS[value]
        return err, err
    if isinstance(value, str) and value in EXCEL_ERRORS.values():
        return value, value
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat(), None
    return value, None


def decode_grid(grid: list[list[Any]]) -> tuple[list[list[Any]], list[tuple[int, int, str]]]:
    """Decode a 2-D grid. Errors are returned as (row_idx, col_idx, text), 0-based within the grid."""
    out: list[list[Any]] = []
    errors: list[tuple[int, int, str]] = []
    for r, row in enumerate(grid):
        new_row = []
        for c, v in enumerate(row):
            val, err = decode_cell(v)
            if err:
                errors.append((r, c, err))
            new_row.append(val)
        out.append(new_row)
    return out, errors


def ej_row_span(first_row: int, max_year: int, first_uwy: int = 2017) -> tuple[int, int]:
    """Legacy rule: rows first_row .. first_row + max_year - first_uwy (inclusive)."""
    return first_row, first_row + max_year - first_uwy


def col_letters(idx: int) -> str:
    """1-based column index -> Excel letters (1 -> A, 27 -> AA)."""
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def letters_to_col(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def cell_address(first_row: int, first_col: int, r: int, c: int) -> str:
    """A1 address of grid cell (r, c), 0-based, for a block whose top-left is (first_row, first_col), 1-based."""
    return f"{col_letters(first_col + c)}{first_row + r}"


def split_address(address: str) -> tuple[int, int]:
    """'C4:S30' -> (first_row, first_col), 1-based."""
    first = address.split(":")[0]
    letters = "".join(ch for ch in first if ch.isalpha())
    return int("".join(ch for ch in first if ch.isdigit())), letters_to_col(letters)
