"""The workbook is located by labels, not by fixed addresses: synthetic grids shaped like the real sheets."""
from resco.locate import locate_ej_blocks, locate_selector_blocks


def grid(rows: int = 140, cols: int = 45) -> list[list]:
    return [[None] * cols for _ in range(rows)]


def put(g: list[list], r: int, c: int, v) -> None:
    g[r - 1][c - 1] = v


COL = {ch: i + 1 for i, ch in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ")}


def main_like(extra_uwy_rows: int = 0, helper_headers: bool = True) -> list[list]:
    """Selector sheet like 'ResCo tables_new': results table, P&L boxes, four triangle blocks."""
    e = extra_uwy_rows
    g = grid()
    put(g, 4, COL["D"], "Gross")
    for j, h in enumerate(["UWY", "Premium Ultimate", "Case Reserve", "Incurred Claims", "SIBNR", "IBNR", "Claims Ultimate", "GULR",
                           "Claims Ultimate", "GULR", "Premium Ultimate", "Claims Ultimate", "GULR", "Claims Ultimate", "GULR", "Δ GULR (QoQ)"]):
        put(g, 5, COL["D"] + j, h)  # D..S
    if helper_headers:  # helper columns right next to the table, not shown on the slides
        for j, h in enumerate(["LL alloc weight", "LL w/o reserves", "Last exp date", "Over LR"]):
            put(g, 5, COL["T"] + j, h)
    put(g, 17 + e, COL["D"], "Total")
    put(g, 30 + 2 * e, COL["D"], "Total")
    put(g, 24, COL["Y"], "Profit (loss) movement QoQ")
    put(g, 29, COL["Y"], "Profit (loss) movement YTD")
    top = 36 + 2 * e
    put(g, top, COL["D"], "Year 1")
    put(g, top, COL["AS"] if False else 40, "Year 1")  # a second set far to the right
    labels = ["GPrmB", "GClmP", "GClmI", "ClmI/GPrmB"]
    for n, lab in enumerate(labels):
        r0 = top + 1 + 14 * n
        put(g, r0, COL["C"], lab)
        put(g, r0, COL["W"], ">=60")
        for r in range(r0 + 1, r0 + 12):  # 11 data rows, some hidden in the real sheet
            put(g, r, COL["C"], "<=2017" if r == r0 + 1 else 2018.0)
            put(g, r, COL["E"], 1.0)
    return g


def test_main_like_layout():
    found = locate_selector_blocks(main_like())
    assert found.problems == []
    assert found.table == "C4:S30"  # right edge = last GULR header, NOT the helper columns T..W
    assert found.table_short == "C4:S17"
    assert found.profit == "Y24:AA27"  # the QoQ box, not the YTD one below it
    assert found.triangles == "C36:W90"


def test_a_new_uwy_row_is_followed():
    found = locate_selector_blocks(main_like(extra_uwy_rows=1))
    assert found.table == "C4:S32" and found.table_short == "C4:S18"
    assert found.triangles == "C38:W92"


def test_missing_labels_are_reported_not_guessed():
    g = main_like()
    put(g, 5, COL["D"], "Underwriting year")  # header renamed
    found = locate_selector_blocks(g)
    assert found.table is None and any("UWY" in p for p in found.problems)


def test_triangles_include_hidden_rows_labelled_one_column_left():
    # UK style: the hidden detail rows carry their label in column B only
    g = grid(80)
    put(g, 32, COL["D"], "Year 1")
    put(g, 33, COL["C"], "GPrmB")
    put(g, 33, COL["W"], ">=60")
    put(g, 60, COL["C"], "ClmI/GPrmB")
    put(g, 61, COL["B"], "<=2017")
    put(g, 62, COL["B"], 2018.0)
    for r in range(63, 68):
        put(g, r, COL["E"], 1.0)
    assert locate_selector_blocks(g).triangles == "C32:W67"


def ej_columns(blocks: list[tuple[str, int]], first_row: int = 5, header_row: int = 4, uwy_start: int = 2017):
    lob, uwy = [None] * 300, [None] * 300
    lob[header_row - 1], uwy[header_row - 1] = "LOB", "UWY"
    r = first_row
    for code, n in blocks:
        for k in range(n):
            lob[r - 1], uwy[r - 1] = code, float(uwy_start + k)
            r += 1
    return lob, uwy


def test_ej_blocks_follow_the_lob_code_when_a_block_is_removed():
    # v2.13: ENOTH sits between ENCON and PRTOT; v2.13b: it is gone and everything below moved up by 10 rows
    old = locate_ej_blocks(*ej_columns([("ENUPS", 10), ("ENDWM", 10), ("ENCON", 10), ("ENOTH", 10), ("PRTOT", 10), ("LIIND", 10)]), max_year=2026)
    new = locate_ej_blocks(*ej_columns([("ENUPS", 10), ("ENDWM", 10), ("ENCON", 10), ("PRTOT", 10), ("LIIND", 10)]), max_year=2026)
    assert (old.blocks["PRTOT"].first, old.blocks["PRTOT"].last) == (45, 54)
    assert (new.blocks["PRTOT"].first, new.blocks["PRTOT"].last) == (35, 44)
    assert "ENOTH" not in new.blocks and new.problems == []


def test_ej_rows_are_limited_to_max_year():
    lay = locate_ej_blocks(*ej_columns([("ENUPS", 10)]), max_year=2023)  # the UK branch only shows UWY up to 2023
    b = lay.blocks["ENUPS"]
    assert (b.first, b.last, b.n_rows) == (5, 11, 7) and b.run_last == 14


def test_ej_a_longer_block_is_used_in_full_up_to_max_year():
    lay = locate_ej_blocks(*ej_columns([("ENUPS", 11)]), max_year=2027)  # next year: a 2027 row appears
    assert lay.blocks["ENUPS"].n_rows == 11


def test_ej_duplicate_code_is_reported():
    lay = locate_ej_blocks(*ej_columns([("ENUPS", 10), ("ENDWM", 10), ("ENUPS", 10)]), max_year=2026)
    assert lay.blocks["ENUPS"].first == 5 and any("more than one block" in p for p in lay.problems)


def test_ej_template_ranges_after_row_growth():
    from resco.config import load_ranges
    from resco.exporter import shift_ej_ranges

    R = load_ranges()
    assert shift_ej_ranges(R, 0) == (tuple(R.ej_clear), tuple(R.ej_blocks_dest), R.ej_print, R.ej_print_life)  # no growth: unchanged
    clear, dests, full, life = shift_ej_ranges(R, 1)  # one more UWY row (2027): every table region gets one row
    assert clear == ("B4:L14", "B18:L28", "B32:L42")  # region i moves down by i rows and grows by one
    assert dests == ("B4", "B18", "K18", "B32")
    assert full == "A1:M43" and life == "A1:M28"  # print areas follow (3 regions / the first two)
