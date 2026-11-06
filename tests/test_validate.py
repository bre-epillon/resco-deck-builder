from resco.validate import check_table_totals, visible_error_cells


def _grid(total_override=None, agg_override=None):
    # columns: C (label C), D (UWY), E amount, F GULR (ignored)
    H = [["", "UWY", "Premium", "GULR"]]
    rows = [
        ["X", "<=2017", 10.0, 0.5],
        ["", 2018.0, 20.0, 0.5],
        ["", 2019.0, 30.0, 0.5],
        ["", "<=2019", agg_override if agg_override is not None else 60.0, 0.5],
        ["", 2020.0, 5.0, 0.5],
        ["", 2026.0, 7.0, 0.5],
        ["", "Total", total_override if total_override is not None else 72.0, 99.0],
    ]
    return [["", "Gross", "", ""]] + H + rows


def test_totals_pass():
    r = check_table_totals(_grid())
    assert r.status == "pass" and r.comparisons == 2 and r.max_rel_diff == 0


def test_totals_fail_on_total_and_on_rollup():
    assert check_table_totals(_grid(total_override=73.0)).status == "fail"
    r = check_table_totals(_grid(agg_override=61.0))
    assert r.status == "fail" and any("<=2019" in f for f in r.failures)


def test_gulr_column_is_not_summed():
    # GULR total (99.0) is nonsense as a sum but must not be checked
    assert check_table_totals(_grid()).status == "pass"


def test_visible_vs_hidden_errors():
    grid = [[1.0, -2146826281], [-2146826246, 2.0]]
    # block at C4:D5; row 5 hidden
    assert visible_error_cells(grid, "C4:D5", {5}, set()) == [("D4", "#DIV/0!")]
    assert visible_error_cells(grid, "C4:D5", set(), {4}) == [("C5", "#N/A")]


def test_blank_spare_row_before_total_is_ignored():
    # UK layout: a blank spare row sits between the last UWY row and Total
    H = [["", "Gross", ""], ["", "UWY", "Premium"]]
    rows = [["", "<=2019", 10.0], ["", 2020.0, 5.0], ["", None, None], ["", "Total", 15.0]]
    r = check_table_totals(H + rows)
    assert r.status == "pass" and r.comparisons == 1
