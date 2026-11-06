"""All xlwings / Excel COM access lives here so the rest of the package is testable without Excel.

The workbook is opened read-only in its own hidden Excel instance and is never saved.
"""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any

import pywintypes
import xlwings as xw
from PIL import Image

from resco.values import decode_cell

log = logging.getLogger(__name__)

XL_SCREEN = 1
XL_PICTURE = -4147
XL_NONE = -4142


class SelectorError(RuntimeError):
    """Selector sheet missing, or the check cell did not follow the selector (stale recalc)."""


class ExportError(RuntimeError):
    """An image could not be produced (clipboard contention, missing chart, empty file...)."""


def clipboard_blocker() -> str | None:
    """Name of the process that currently has the Windows clipboard open, if any (diagnostics for failed exports)."""
    try:
        import ctypes
        import os
        from ctypes import wintypes

        user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
        user32.GetOpenClipboardWindow.restype = wintypes.HWND
        kernel32.OpenProcess.restype = wintypes.HANDLE
        hwnd = user32.GetOpenClipboardWindow()
        if not hwnd:
            return None
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        handle = kernel32.OpenProcess(0x1000, False, pid.value)
        buf, size = ctypes.create_unicode_buffer(260), wintypes.DWORD(260)
        if handle and kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value)
        return f"pid {pid.value}"
    except Exception:  # noqa: BLE001 - purely diagnostic
        return None


class _Retry(Exception):
    """Raised inside an export attempt to ask _retry for another go."""


class ExcelSession:
    def __init__(self, workbook: Path, visible: bool = False, picture_mode: str = "auto") -> None:
        """picture_mode: 'clipboard' (legacy CopyPicture only), 'pdf' (Excel -> PDF -> PNG, needs no display or
        clipboard) or 'auto' (CopyPicture first; after repeated failures the rest of the session uses PDF)."""
        self.workbook_path = Path(workbook).resolve()
        self.visible = visible
        self.picture_mode = picture_mode
        self.clipboard_broken = False  # set once CopyPicture has failed in 'auto' mode
        self.last_method = ""  # how the most recent range picture was made: 'clipboard' or 'pdf'
        self._pdf_sheets: set[str] = set()
        self.app: xw.App | None = None
        self.wb: xw.Book | None = None

    def __enter__(self) -> "ExcelSession":
        self.app = xw.App(visible=self.visible, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        log.info("Excel started (pid %s)", self.app.pid)
        self.wb = self.app.books.open(str(self.workbook_path), read_only=True, update_links=False)
        return self

    def __exit__(self, *exc: object) -> None:
        try:
            if self.wb is not None:
                self.wb.close()  # read-only: there is no Save call anywhere in this module
        finally:
            if self.app is not None:
                try:
                    self.app.quit()
                except Exception:  # noqa: BLE001 - last resort so we never leave a hidden Excel behind
                    log.warning("quit failed, killing Excel pid %s", self.app.pid)
                    self.app.kill()

    # -- workbook queries -------------------------------------------------
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.wb.sheets]

    def has_sheet(self, name: str) -> bool:
        return name in self.sheet_names()

    def asof_year(self, name: str = "asof") -> int:
        value = self.wb.names[name].refers_to_range.value
        return value.year

    # -- selector ---------------------------------------------------------
    def select(self, sheet: str, cell: str, check_cell: str, selector: str, expected_code: str) -> None:
        if not self.has_sheet(selector):
            raise SelectorError(f"summary sheet '{selector}' does not exist in the workbook")
        ws = self.wb.sheets[sheet]
        ws[cell].value = selector
        shown = None
        for attempt in (1, 2):
            self.app.api.Calculate()
            shown = ws[check_cell].value
            if shown == expected_code:
                return
            log.warning("check cell shows %r, expected %r (attempt %d)", shown, expected_code, attempt)
        raise SelectorError(f"after selecting {selector}, {check_cell} shows {shown!r}, expected {expected_code!r}")

    # -- reading ----------------------------------------------------------
    def read_grid(self, sheet: str, address: str) -> list[list[Any]]:
        """Raw Value2 grid. xlwings' own .value converts Excel errors to None, which would hide them,
        so the COM value is read directly: errors arrive as negative ints (see values.EXCEL_ERRORS)."""
        raw = self.wb.sheets[sheet].range(address).api.Value2
        if not isinstance(raw, tuple):  # single cell
            raw = ((raw,),)
        return [list(row) for row in raw]

    def hidden_rows_cols(self, sheet: str, address: str) -> tuple[set[int], set[int]]:
        """Absolute 1-based row and column numbers that are hidden inside a range (not rendered in a picture)."""
        ws = self.wb.sheets[sheet]
        rng = ws.range(address)
        r0, c0 = rng.row, rng.column
        rows = {r for r in range(r0, r0 + rng.rows.count) if ws.api.Rows(r).Hidden}
        cols = {c for c in range(c0, c0 + rng.columns.count) if ws.api.Columns(c).Hidden}
        return rows, cols

    def hide_rows(self, sheet: str, spans: tuple[str, ...]) -> None:
        """Hide row spans such as '6:8' in memory (the workbook is read-only and never saved), so they are not drawn."""
        ws = self.wb.sheets[sheet]
        for span in spans:
            ws.api.Rows(span).Hidden = True

    def read_sheet_values(self, sheet: str, max_rows: int = 250, max_cols: int = 80) -> list[list[Any]]:
        """Top-left part of a sheet as one rectangular grid (Value2; errors stay recognisable, see values.decode_grid)."""
        ws = self.wb.sheets[sheet]
        last = ws.used_range.last_cell
        rows, cols = min(last.row, max_rows), min(last.column, max_cols)
        raw = ws.api.Range(ws.api.Cells(1, 1), ws.api.Cells(rows, cols)).Value2
        if not isinstance(raw, tuple):
            raw = ((raw,),)
        return [list(r) for r in raw]

    def used_last_row(self, sheet: str) -> int:
        return self.wb.sheets[sheet].used_range.last_cell.row

    # -- image export -----------------------------------------------------
    @staticmethod
    def _retry(what: str, fn: Any, retries: int = 10) -> None:
        """CopyPicture / Paste fail intermittently when another program holds the clipboard (Remote Desktop clipboard
        redirection, clipboard managers, a user copying something): retry with a growing pause."""
        last: Exception | None = None
        for attempt in range(retries):
            try:
                fn()
                return
            except (pywintypes.com_error, _Retry) as exc:
                last = exc
                time.sleep(min(0.5 * (attempt + 1), 3.0))
        blocker = clipboard_blocker()
        hint = f" (clipboard currently held by {blocker})" if blocker else ""
        raise ExportError(f"{what} failed after {retries} attempts{hint}: {last}")

    @staticmethod
    def _activate(ws: Any) -> None:
        """Worksheet.Activate through COM. xlwings' Sheet.activate() also tries to bring the (hidden) Excel window to
        the foreground and intermittently fails with 'Could not activate App!'."""
        try:
            ws.api.Parent.Activate()  # the workbook first: a sheet of a background workbook cannot be activated
            ws.api.Activate()
        except pywintypes.com_error:
            pass  # sheet could not be activated (e.g. hidden); exports still work for most objects

    @staticmethod
    def _check_png(path: Path) -> tuple[int, int]:
        if not path.exists() or path.stat().st_size == 0:
            raise ExportError(f"{path.name}: no image was written")
        with Image.open(path) as im:
            return im.size

    def export_range_png(self, sheet: str, address: str, path: Path, book: xw.Book | None = None) -> tuple[int, int]:
        """Picture of a range (hidden rows/columns are not drawn). Legacy method first: CopyPicture(xlScreen) -> temporary
        chart object -> Chart.Export. It needs a live display and a free clipboard; when that fails (auto mode) the
        PDF route takes over for the rest of the session."""
        path = path.resolve()  # Excel resolves relative paths against ITS working directory, not ours
        path.parent.mkdir(parents=True, exist_ok=True)
        ws = (book or self.wb).sheets[sheet]
        if self.picture_mode == "pdf" or (self.picture_mode == "auto" and self.clipboard_broken):
            return self._export_range_pdf(ws, address, path)
        try:
            size = self._export_range_clipboard(ws, address, path, retries=3 if self.picture_mode == "auto" else 10)
            self.last_method = "clipboard"
            return size
        except ExportError:
            if self.picture_mode != "auto":
                raise
            log.warning("CopyPicture keeps failing (no display or clipboard busy): switching to the PDF route for this session")
            self.clipboard_broken = True
            return self._export_range_pdf(ws, address, path)

    def _export_range_clipboard(self, ws: Any, address: str, path: Path, retries: int) -> tuple[int, int]:
        self._activate(ws)
        rng = ws.range(address).api

        def once() -> None:
            co = None
            try:
                rng.CopyPicture(Appearance=XL_SCREEN, Format=XL_PICTURE)
                co = ws.api.ChartObjects().Add(rng.Left, rng.Top, rng.Width, rng.Height)
                co.Border.LineStyle = XL_NONE
                co.Activate()  # pasting into a chart that is not active fails for some ranges (legacy used .Select)
                co.Chart.Paste()
                time.sleep(0.3)
                co.Chart.Export(str(path), "PNG")
            finally:
                if co is not None:
                    try:
                        co.Delete()
                    except pywintypes.com_error:
                        pass

        self._retry(f"export {ws.name}!{address}", once, retries=retries)
        self.app.api.CutCopyMode = False
        return self._check_png(path)

    def _export_range_pdf(self, ws: Any, address: str, path: Path) -> tuple[int, int]:
        """Display- and clipboard-independent: Excel prints the range to one PDF page, pdfium rasterises it, the white
        margin is trimmed. Looks slightly different from the screen picture (print metrics, crisper)."""
        import pypdfium2 as pdfium
        from PIL import ImageChops

        app = self.app.api
        pdf = path.with_suffix(".pdf")
        key = f"{ws.book.name}!{ws.name}"
        self._activate(ws)  # PageSetup only works on the active workbook/sheet (the helper workbook is opened last)
        try:
            app.PrintCommunication = False  # batch the page-setup changes (otherwise each one takes seconds)
            ps = ws.api.PageSetup
            ps.PrintArea = ws.range(address).api.Address
            ps.Orientation = 2  # landscape
            ps.PaperSize = 8  # A3: leaves headroom so the one-page fit stays crisp
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 0
            ps.HeaderMargin = ps.FooterMargin = 0
            ps.CenterHorizontally = ps.CenterVertically = False
            ps.LeftHeader = ps.CenterHeader = ps.RightHeader = ps.LeftFooter = ps.CenterFooter = ps.RightFooter = ""
        finally:
            try:
                app.PrintCommunication = True
            except pywintypes.com_error:
                pass  # do not mask the original error
        self._pdf_sheets.add(key)
        if pdf.exists():
            pdf.unlink()
        self._retry(f"PDF export {ws.name}!{address}", lambda: ws.range(address).api.ExportAsFixedFormat(0, str(pdf), 0, False, False), retries=3)
        doc = pdfium.PdfDocument(str(pdf))
        try:
            page = doc[0]
            w_pt, h_pt = page.get_size()
            img = page.render(scale=max(3.0, 3600 / max(w_pt, h_pt))).to_pil().convert("RGB")
        finally:
            doc.close()
        bg = ImageChops.difference(img, Image.new("RGB", img.size, (255, 255, 255)))
        box = bg.point(lambda v: 255 if v > 12 else 0).getbbox()
        if box is None:
            raise ExportError(f"{path.name}: the PDF page is empty")
        pad = 4
        img = img.crop((max(box[0] - pad, 0), max(box[1] - pad, 0), min(box[2] + pad, img.width), min(box[3] + pad, img.height)))
        img.save(path)
        pdf.unlink(missing_ok=True)
        self.last_method = "pdf"
        return img.size

    @staticmethod
    def chart_title(chart_object: Any) -> str:
        try:
            return chart_object.Chart.ChartTitle.Text if chart_object.Chart.HasTitle else ""
        except pywintypes.com_error:
            return ""

    def chart_by_name(self, sheet: str, name: str) -> Any | None:
        ws = self.wb.sheets[sheet]
        for i in range(1, ws.api.ChartObjects().Count + 1):
            co = ws.api.ChartObjects(i)
            if co.Name == name:
                return co
        return None

    def charts_by_title(self, sheet: str, title: str) -> list[Any]:
        """Chart objects whose title matches, bottom-most first (so the last one is what the viewer sees on top)."""
        ws = self.wb.sheets[sheet]
        found: list[tuple[int, int, Any]] = []
        for i in range(1, ws.api.ChartObjects().Count + 1):
            co = ws.api.ChartObjects(i)
            try:
                t = co.Chart.ChartTitle.Text if co.Chart.HasTitle else ""
            except pywintypes.com_error:
                t = ""
            if t == title:
                try:
                    z = int(co.ShapeRange.ZOrderPosition)
                except pywintypes.com_error:
                    z = i
                found.append((z, i, co))
        found.sort(key=lambda x: (x[0], x[1]))
        return [co for _, _, co in found]

    def export_chart_png(self, sheet: str, chart_object: Any, path: Path) -> tuple[int, int]:
        path = path.resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        self._activate(self.wb.sheets[sheet])

        def once() -> None:
            chart_object.Activate()  # legacy activated the chart before exporting; unrendered charts export nothing
            if not chart_object.Chart.Export(str(path), "PNG"):
                raise _Retry("Chart.Export returned False")

        self._retry(f"export chart on {sheet}", once)
        return self._check_png(path)

    def open_aux(self, path: Path) -> xw.Book:
        """Open a helper workbook (read-only; edited in memory only, never saved)."""
        return self.app.books.open(str(path), read_only=True, update_links=False)

    def expand_table_regions(self, aux: xw.Book, sheet: str, regions: list[str], extra: int) -> None:
        """Insert `extra` rows into each table region (bottom-up, above the region's last row so the new rows take the
        data-row formatting). In memory only: the helper workbook is never saved."""
        ws = aux.sheets[sheet]
        for region in reversed(regions):
            last = ws.range(region).row + ws.range(region).rows.count - 1
            ws.api.Rows(f"{last}:{last + extra - 1}").Insert()

    def render_expert_judgement(
        self,
        aux: xw.Book,
        scratch_sheet: str,
        clear: tuple[str, ...],
        blocks: list[tuple[str, str]],  # (source columns "E:O", destination top-left "B4")
        ej_sheet: str,
        first_row: int,
        last_row: int,
        print_range: str,
        path: Path,
    ) -> tuple[int, int]:
        """Legacy method without the clipboard: copy VALUES into the formatted scratch sheet, then export it."""
        ws = aux.sheets[scratch_sheet]
        n_rows = last_row - first_row + 1
        for address in clear:
            rng = ws.range(address)
            rng.api.ClearContents()
            rng.api.EntireRow.Hidden = False  # reset what a previous LoB hid
            if n_rows < rng.rows.count:  # fewer UWY rows than the block holds (e.g. UK, 7 rows): hide the unused ones
                ws.api.Rows(f"{rng.row + n_rows}:{rng.row + rng.rows.count - 1}").Hidden = True
        src_ws = self.wb.sheets[ej_sheet]
        for cols, dest in blocks:
            c0, c1 = cols.split(":")
            raw = src_ws.range(f"{c0}{first_row}:{c1}{last_row}").api.Value2
            clean = tuple(tuple(decode_cell(v)[0] for v in row) for row in raw)
            ws.range(dest).resize(len(clean), len(clean[0])).api.Value2 = clean
        return self.export_range_png(scratch_sheet, print_range, path, book=aux)

    # -- charts -----------------------------------------------------------
    def chart_inventory(self, sheet: str) -> list[dict[str, Any]]:
        """Name, position, size and title of every chart object on a sheet (names can be duplicated)."""
        rows: list[dict[str, Any]] = []
        ws = self.wb.sheets[sheet]
        count = ws.api.ChartObjects().Count
        for i in range(1, count + 1):
            co = ws.api.ChartObjects(i)
            try:
                title = co.Chart.ChartTitle.Text if co.Chart.HasTitle else ""
            except Exception:  # noqa: BLE001 - some chart types raise on title access
                title = ""
            rows.append(
                {
                    "sheet": sheet,
                    "index": i,
                    "name": co.Name,
                    "top_left": co.TopLeftCell.Address.replace("$", ""),
                    "width": round(co.Width),
                    "height": round(co.Height),
                    "visible": bool(co.Visible),
                    "title": title,
                }
            )
        return rows
