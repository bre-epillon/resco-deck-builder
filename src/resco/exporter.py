"""Turn one LoB into the PNG images its slides need (Excel side of Step 4). All Excel access goes through ExcelSession."""
from __future__ import annotations

import hashlib
import logging
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageChops

from resco.config import Lob, Ranges
from resco.excel_reader import ExcelSession, ExportError
from resco.layout import Layout
from resco.values import split_address

log = logging.getLogger(__name__)

KIND_ORDER = [
    "table", "table_life", "profit", "triangles", "chart_prem", "chart_paid", "chart_incurr",
    "pattern_prem", "pattern_inc", "ej", "ej_life",
]  # fmt: skip
DUPLICATE_DIFF_WARN = 0.001  # fraction of pixels
RANGE_KINDS = {"table", "table_life", "profit", "triangles", "ej", "ej_life"}  # pictures of cell ranges


@dataclass
class ExportedImage:
    kind: str
    path: Path
    width: int
    height: int
    sha256: str
    method: str = "chart"  # clipboard | pdf (range pictures) or chart (chart exports)


@dataclass
class ExportIssue:
    level: str  # ERROR | WARN | INFO
    lob: str
    kind: str
    message: str


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def pixel_diff_fraction(a: Path, b: Path, threshold: int = 8) -> float:
    """Share of pixels whose channels differ by more than `threshold` (1.0 when sizes differ)."""
    with Image.open(a) as ia, Image.open(b) as ib:
        if ia.size != ib.size:
            return 1.0
        hist = ImageChops.difference(ia.convert("RGB"), ib.convert("RGB")).convert("L").histogram()
        return sum(hist[threshold + 1 :]) / (ia.width * ia.height)


def export_lob(
    session: ExcelSession,
    lob: Lob,
    kinds: set[str],
    ranges: Ranges,
    layout: Layout,
    img_dir: Path,
    max_year: int,
    ej_book: object | None,
) -> tuple[dict[str, ExportedImage], list[ExportIssue]]:
    """Select the LoB (raises SelectorError when the recalc did not follow) and export every requested slot kind."""
    session.select(ranges.selector_sheet, ranges.selector_cell, ranges.check_cell, lob.selector, lob.expected_check)
    images: dict[str, ExportedImage] = {}
    issues: list[ExportIssue] = []
    for kind in sorted(kinds, key=lambda k: KIND_ORDER.index(k) if k in KIND_ORDER else 99):
        path = img_dir / lob.code / f"{kind}.png"
        try:
            size = _export_kind(session, lob, kind, ranges, layout, path, ej_book, issues)
        except Exception as exc:  # noqa: BLE001 - one failing slot must not lose the rest of the run
            log.error("%s/%s: %s", lob.code, kind, exc, exc_info=True)
            issues.append(ExportIssue("ERROR", lob.code, kind, str(exc)))
            continue
        method = session.last_method if kind in RANGE_KINDS else "chart"
        images[kind] = ExportedImage(kind, path, size[0], size[1], sha256_file(path), method)
    return images, issues


def _export_kind(
    session: ExcelSession,
    lob: Lob,
    kind: str,
    ranges: Ranges,
    layout: Layout,
    path: Path,
    ej_book: object | None,
    issues: list[ExportIssue],
) -> tuple[int, int]:
    sheet = ranges.selector_sheet
    range_by_kind = {"table": layout.table, "table_life": layout.table_short, "profit": layout.profit, "triangles": layout.triangles}
    if kind in range_by_kind:
        return session.export_range_png(sheet, range_by_kind[kind], path)

    if kind in ranges.chart_by_kind:
        chart = pick_chart(session, sheet, ranges.chart_by_kind[kind], ranges.chart_title_by_kind.get(kind, ""), lob.code, kind, issues)
        return session.export_chart_png(sheet, chart, path)

    if kind in ranges.pattern_title_by_kind:
        title = ranges.pattern_title_by_kind[kind]
        if not lob.resclass:
            raise ExportError("LoB has no ResClass sheet")
        if not session.has_sheet(lob.resclass):
            raise ExportError(f"sheet '{lob.resclass}' not found")
        candidates = session.charts_by_title(lob.resclass, title)
        if not candidates:
            raise ExportError(f"no chart titled '{title}' on '{lob.resclass}'")
        size = session.export_chart_png(lob.resclass, candidates[-1], path)  # topmost copy = what Excel shows
        notes: list[str] = []
        worst = 0.0
        for i, other in enumerate(candidates[:-1]):
            dup = path.with_name(f"{path.stem}_dup{i}.png")
            dsize = session.export_chart_png(lob.resclass, other, dup)
            if dsize != size:
                notes.append(f"'{other.Name}' has a different size ({dsize[0]}x{dsize[1]} vs {size[0]}x{size[1]})")
                worst = 1.0
            else:
                frac = pixel_diff_fraction(path, dup)
                worst = max(worst, frac)
                notes.append(f"'{other.Name}' differs in {frac:.2%} of pixels")
        if candidates[:-1]:
            issues.append(
                ExportIssue(
                    "WARN" if worst > DUPLICATE_DIFF_WARN else "INFO",
                    lob.code,
                    kind,
                    f"{len(candidates)} charts titled '{title}' on '{lob.resclass}'; topmost '{candidates[-1].Name}' used; " + "; ".join(notes),
                )
            )
        return size

    if kind in ("ej", "ej_life"):
        if ej_book is None:
            raise ExportError("the Expert Judgement render template is not open")
        ej: EjTemplate = ej_book  # type: ignore[assignment]
        block = layout.ej.get(lob.resclass)
        if block is None:
            raise ExportError(f"no Expert Judgement block for '{lob.resclass}' in the '{ranges.ej_sheet}' sheet (blocks found for: {', '.join(layout.ej) or 'none'})")
        if block.n_rows > ej.capacity:
            raise ExportError(f"the Expert Judgement block of '{lob.resclass}' has {block.n_rows} UWY rows but the render template holds {ej.capacity}")
        blocks = [(cols, dest) for (_, cols), dest in zip(layout.ej_block_cols, ej.dests)]
        return session.render_expert_judgement(
            ej.book,  # type: ignore[arg-type]
            ranges.ej_scratch_sheet,
            ej.clear,
            blocks,
            ranges.ej_sheet,
            block.first,
            block.last,
            ej.print_life if kind == "ej_life" else ej.print_full,
            path,
        )
    raise ValueError(f"unknown slot kind '{kind}'")


@dataclass
class EjTemplate:
    """The Expert Judgement render workbook, open, with its ranges as they are after any row growth."""

    book: object
    clear: tuple[str, ...]
    dests: tuple[str, ...]
    print_full: str
    print_life: str
    capacity: int  # UWY rows each table region holds


def shift_ej_ranges(ranges: Ranges, extra: int) -> tuple[tuple[str, ...], tuple[str, ...], str, str]:
    """Ranges of the render template after `extra` rows were inserted into every table region (region i moves down by i*extra)."""
    spans = []
    for a in ranges.ej_clear:
        first, last = a.split(":")
        spans.append((split_address(first)[0], split_address(last)[0], first.rstrip("0123456789"), last.rstrip("0123456789")))

    def region_of(row: int) -> int:
        return next((i for i, (r1, r2, *_) in enumerate(spans) if r1 <= row <= r2), len(spans) - 1)

    clear = tuple(f"{c1}{r1 + i * extra}:{c2}{r2 + (i + 1) * extra}" for i, (r1, r2, c1, c2) in enumerate(spans))
    dests = tuple(f"{d.rstrip('0123456789')}{split_address(d)[0] + region_of(split_address(d)[0]) * extra}" for d in ranges.ej_blocks_dest)

    def grow(print_range: str) -> str:
        first, last = print_range.split(":")
        row = split_address(last)[0]
        inside = sum(1 for (_, r2, *_) in spans if r2 <= row)
        return f"{first}:{last.rstrip('0123456789')}{row + inside * extra}"

    return clear, dests, grow(ranges.ej_print), grow(ranges.ej_print_life)


def prepare_ej_template(session: ExcelSession, path: Path, ranges: Ranges, needed_rows: int) -> EjTemplate:
    """Open the render template and, when a workbook block has more UWY rows than it holds (a new UWY row every January),
    insert the missing rows into every table region instead of failing."""
    book = session.open_aux(path)
    capacity = ej_capacity(ranges)
    extra = max(0, needed_rows - capacity)
    if extra:
        log.warning("Expert Judgement blocks have %d UWY rows, the render template holds %d: inserting %d row(s) per table (in memory)", needed_rows, capacity, extra)
        session.expand_table_regions(book, ranges.ej_scratch_sheet, list(ranges.ej_clear), extra)
    clear, dests, print_full, print_life = shift_ej_ranges(ranges, extra)
    return EjTemplate(book, clear, dests, print_full, print_life, capacity + extra)


def ej_capacity(ranges: Ranges) -> int:
    """UWY rows the Expert Judgement render template holds per table (rows of the first 'clear' range, e.g. B4:L13 -> 10)."""
    first, last = ranges.ej_clear[0].split(":")
    return split_address(last)[0] - split_address(first)[0] + 1


def pick_chart(session: ExcelSession, sheet: str, name: str, title: str, lob_code: str, kind: str, issues: list[ExportIssue]) -> object:
    """The configured chart name if its title is right; otherwise the chart with the right title (topmost copy wins);
    otherwise the named chart with a warning. Charts get renamed and moved between workbook versions, titles rarely."""
    by_name = session.chart_by_name(sheet, name) if name else None
    if by_name is not None and (not title or session.chart_title(by_name) == title):
        return by_name
    candidates = session.charts_by_title(sheet, title) if title else []
    if candidates:
        why = f"no chart named '{name}'" if by_name is None else f"chart '{name}' has another title"
        issues.append(ExportIssue("INFO", lob_code, kind, f"chart found by its title '{title}' ('{candidates[-1].Name}'): {why} on '{sheet}'"))
        return candidates[-1]
    if by_name is not None:
        issues.append(ExportIssue("WARN", lob_code, kind, f"chart '{name}' on '{sheet}' has the title '{session.chart_title(by_name)}', expected '{title}'; used anyway"))
        return by_name
    raise ExportError(f"chart '{name}' (title '{title}') not found on '{sheet}'")
