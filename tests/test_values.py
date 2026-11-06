import datetime as dt

from resco.values import decode_cell, decode_grid, ej_row_span


def test_error_codes_decoded():
    assert decode_cell(-2146826281) == ("#DIV/0!", "#DIV/0!")
    assert decode_cell(-2146826246) == ("#N/A", "#N/A")


def test_plain_values_untouched():
    assert decode_cell(1.5) == (1.5, None)
    assert decode_cell(None) == (None, None)
    assert decode_cell("abc") == ("abc", None)
    assert decode_cell(True) == (True, None)


def test_dates_iso():
    assert decode_cell(dt.datetime(2026, 9, 30)) == ("2026-09-30T00:00:00", None)


def test_grid_error_positions():
    grid = [[1, -2146826281], [None, "x"]]
    clean, errs = decode_grid(grid)
    assert clean[0][1] == "#DIV/0!"
    assert errs == [(0, 1, "#DIV/0!")]


def test_column_letters_roundtrip():
    from resco.values import cell_address, col_letters, letters_to_col

    assert col_letters(1) == "A" and col_letters(27) == "AA" and col_letters(38) == "AL"
    assert letters_to_col("AL") == 38 and letters_to_col("C") == 3
    assert cell_address(4, 3, 0, 0) == "C4"
    assert cell_address(4, 3, 26, 16) == "S30"


def test_ej_row_span_matches_legacy_ranges():
    # legacy: ENUPS first row 5, max year 2026 -> rows 5..14 (10 rows)
    assert ej_row_span(5, 2026) == (5, 14)
    assert ej_row_span(175, 2026) == (175, 184)
