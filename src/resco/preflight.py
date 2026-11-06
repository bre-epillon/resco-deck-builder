"""Structure checks on the workbook, fast (no recalculation per LoB, no pictures): is everything the build needs there?
Run it first whenever the workbook or the deck changed; the build runs the same checks at its start."""
from __future__ import annotations

from resco.config import Lob, Ranges
from resco.excel_reader import ExcelSession
from resco.exporter import ExportIssue, ej_capacity
from resco.layout import Layout


def preflight(xl: ExcelSession, ranges: Ranges, layout: Layout, lob_by_code: dict[str, Lob], kinds_by_lob: dict[str, set[str]]) -> list[ExportIssue]:
    issues: list[ExportIssue] = []
    sheet = ranges.selector_sheet
    capacity = ej_capacity(ranges)
    title_cache: dict[tuple[str, str], int] = {}

    def n_with_title(sheet_name: str, title: str) -> int:
        key = (sheet_name, title)
        if key not in title_cache:
            title_cache[key] = len(xl.charts_by_title(sheet_name, title))
        return title_cache[key]

    tri_ok: dict[str, bool] = {}
    for kind in {k for ks in kinds_by_lob.values() for k in ks} & set(ranges.chart_by_kind):
        name, title = ranges.chart_by_kind[kind], ranges.chart_title_by_kind.get(kind, "")
        by_name = xl.chart_by_name(sheet, name)
        tri_ok[kind] = by_name is not None or (bool(title) and n_with_title(sheet, title) > 0)

    for code, kinds in kinds_by_lob.items():
        lob = lob_by_code[code]
        if not xl.has_sheet(lob.selector):
            issues.append(ExportIssue("ERROR", code, "*", f"summary sheet '{lob.selector}' does not exist in the workbook"))
        for kind in sorted(kinds & set(ranges.chart_by_kind)):
            if not tri_ok.get(kind, False):
                issues.append(ExportIssue("ERROR", code, kind, f"chart '{ranges.chart_by_kind[kind]}' (title '{ranges.chart_title_by_kind.get(kind, '')}') not found on '{sheet}'"))
        if kinds & {"ej", "ej_life"}:
            block = layout.ej.get(lob.resclass)
            kind = "ej_life" if "ej_life" in kinds else "ej"
            if block is None:
                issues.append(ExportIssue("ERROR", code, kind, f"no Expert Judgement block for '{lob.resclass}' in '{ranges.ej_sheet}' (blocks found for: {', '.join(layout.ej) or 'none'})"))
            elif block.n_rows > capacity:
                issues.append(ExportIssue("ERROR", code, kind, f"the Expert Judgement block of '{lob.resclass}' has {block.n_rows} UWY rows but the render template holds {capacity}"))
        for kind in sorted(kinds & set(ranges.pattern_title_by_kind)):
            title = ranges.pattern_title_by_kind[kind]
            if not lob.resclass or not xl.has_sheet(lob.resclass):
                issues.append(ExportIssue("ERROR", code, kind, f"sheet '{lob.resclass}' for the pattern charts does not exist"))
            elif n_with_title(lob.resclass, title) == 0:
                issues.append(ExportIssue("ERROR", code, kind, f"no chart titled '{title}' on '{lob.resclass}'"))
    return issues


def slot_problem(issues: list[ExportIssue], lob: str, kind: str) -> str | None:
    """First ERROR that stops this slot from being produced, for the per-slot status of a workbook check."""
    for i in issues:
        if i.level == "ERROR" and i.lob == lob and i.kind in (kind, "*"):
            return i.message
    return None
