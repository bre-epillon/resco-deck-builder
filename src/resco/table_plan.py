"""Glue for the table filling in a build: choose the targets, load their sources (netting ladder, IBNR 'Overview'),
make the table jobs and summarise what happened for the report and the dashboard."""
from __future__ import annotations

import hashlib
import logging
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from resco.config import LADDER_KINDS, TableConfig
from resco.deck_scan import TableMatch
from resco.deck_writer import TableJob
from resco.exporter import ExportIssue
from resco.tables import BlockTexts, Ladder, LadderError

log = logging.getLogger(__name__)
TABLES_LOB = "(tables)"


def local_copy(path: Path, name: str) -> tuple[Path, str]:
    """Copy a workbook to a local temp folder (never open the OneDrive original); returns (copy, sha256[:12])."""
    dest_dir = Path(tempfile.gettempdir()) / "resco_work"
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / f"{name}{path.suffix}"
    shutil.copy2(path, dest)
    h = hashlib.sha256()
    with dest.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return dest, h.hexdigest()[:12]


@dataclass
class TableSources:
    ladder: Ladder | None = None
    ladder_error: str | None = None
    prior: Ladder | None = None  # the Q-1 netting ladder (Profit (loss) movement QoQ boxes)
    prior_error: str | None = None
    overview: dict[str, tuple[list[BlockTexts | None], list[str]]] = field(default_factory=dict)
    ibnr_error: str | None = None
    info: dict[str, str] = field(default_factory=dict)


NO_PRIOR = "no Q-1 netting file was given (--netting-prior): the Profit (loss) movement QoQ boxes are left as they are"


def load_sources(
    matches: list[TableMatch], tcfg: TableConfig, netting_path: Path | None, ibnr_path: Path | None, xl: Any | None, netting_prior_path: Path | None = None,
) -> TableSources:
    """`xl` is the open ExcelSession (needed for the IBNR workbook only)."""
    src = TableSources()
    if any(m.target.kind in LADDER_KINDS for m in matches):
        if netting_path is None:
            src.ladder_error = "no netting workbook was given (--netting)"
        else:
            try:
                src.ladder = Ladder(netting_path, tcfg.netting_sheet)
                src.info["netting"] = netting_path.name
            except Exception as exc:  # noqa: BLE001 - a broken file must not stop the pictures
                src.ladder_error = f"netting workbook could not be read: {exc}"
    if any(m.target.kind == "profit_qoq" for m in matches):
        if netting_prior_path is None:
            src.prior_error = NO_PRIOR
        else:
            try:
                src.prior = Ladder(netting_prior_path, tcfg.netting_sheet)
                src.info["netting_prior"] = netting_prior_path.name
            except Exception as exc:  # noqa: BLE001
                src.prior_error = f"Q-1 netting file could not be read: {exc}"
    overview = [m for m in matches if m.target.kind == "overview"]
    if overview:
        if ibnr_path is None:
            src.ibnr_error = "no IBNR allocation workbook was given (--ibnr)"
        elif xl is None:
            src.ibnr_error = "Excel is not available for the IBNR allocation workbook"
        else:
            from resco.tables_excel import read_blocks

            try:
                copy, sha = local_copy(ibnr_path, "ibnr_copy")
                log.info("opening a local copy of %s (can take half a minute)", ibnr_path.name)
                book = xl.open_aux(copy)
                try:
                    for m in overview:
                        src.overview[m.target.id] = read_blocks(book, tcfg.ibnr_sheet, list(m.target.blocks))
                finally:
                    book.close()
                src.info["ibnr"] = f"{ibnr_path.name} ({sha})"
            except Exception as exc:  # noqa: BLE001
                log.exception("IBNR workbook")
                src.ibnr_error = f"IBNR allocation workbook could not be read: {exc}"
    return src


def make_jobs(matches: list[TableMatch], src: TableSources) -> list[TableJob]:
    jobs: list[TableJob] = []
    for m in matches:
        t = m.target
        job = TableJob(m.slide, m.title, m.shape_id, t.id, t.kind, lob=t.select.get("lob", ""))
        if t.kind == "profit_qoq":
            if src.ladder is None:
                job.error = src.ladder_error or "no netting data"
            elif src.prior is None:
                job.error, job.warn_only = src.prior_error or NO_PRIOR, src.prior_error in (None, NO_PRIOR)
            else:
                try:
                    job.block = src.ladder.block(t.select)
                    job.ladder_groups = src.ladder.groups
                except LadderError as exc:
                    job.error = str(exc)
                try:  # a LoB that is new this quarter has no Q-1 rows: the boxes are left empty and that is reported
                    job.prior_block = src.prior.block(t.select)
                except LadderError:
                    job.prior_block = None
        elif t.kind == "ladder":
            if src.ladder is None:
                job.error = src.ladder_error or "no netting data"
            else:
                try:
                    job.block = src.ladder.block(t.select)
                    job.ladder_groups = src.ladder.groups
                except LadderError as exc:
                    job.error = str(exc)
        else:
            if src.ibnr_error:
                job.error = src.ibnr_error
            else:
                blocks, notes = src.overview.get(t.id, ([None] * len(t.blocks), ["not read"]))
                job.blocks = blocks
                if all(b is None for b in blocks):
                    job.error = "; ".join(notes) or "no source block found"
        jobs.append(job)
    return jobs


def describe_source(m: TableMatch) -> str:
    t = m.target
    if t.kind == "ladder":
        return "netting ladder: " + ", ".join(f"{k}={v}" for k, v in t.select.items())
    if t.kind == "profit_qoq":
        return "Profit (loss) movement QoQ, netting ladder Q vs Q-1, Total row: " + ", ".join(f"{k}={v}" for k, v in t.select.items())
    return "IBNR 'Overview': " + "; ".join(str(b.get("title") or " ".join(b.get("header", []))) + f" #{b.get('occurrence', 1)}" for b in t.blocks)


def planned_entries(matches: list[TableMatch]) -> list[dict[str, Any]]:
    return [{"id": m.target.id, "kind": m.target.kind, "slide": m.slide, "title": m.title, "table": m.table_name, "rows": m.rows, "cols": m.cols,
             "status": "planned", "source": describe_source(m)} for m in matches]  # fmt: skip


def report(jobs: list[TableJob], matches: list[TableMatch], applied: bool = True, fresh_slides: frozenset[int] = frozenset()) -> tuple[list[dict[str, Any]], list[ExportIssue], dict[int, set[int]]]:
    """Per-table entries for the JSON / dashboard, the issues they raise, and slide -> filled shape ids for the verification."""
    by_id = {(m.target.id, m.slide): m for m in matches}
    entries: list[dict[str, Any]] = []
    issues: list[ExportIssue] = []
    filled: dict[int, set[int]] = {}
    for j in jobs:
        m = by_id[(j.target_id, j.slide)]
        kind = f"table:{j.target_id}"
        res = j.result
        entry: dict[str, Any] = {"id": j.target_id, "kind": j.kind, "slide": j.slide, "title": j.title, "table": m.table_name, "rows": m.rows, "cols": m.cols, "source": describe_source(m)}
        if j.error or res is None:
            msg = j.error or "the table was not filled"
            entry["status"] = f"skipped: {msg}" if j.warn_only else f"failed: {msg}"
            issues.append(ExportIssue("WARN" if j.warn_only else "ERROR", TABLES_LOB, kind, f"slide {j.slide}: {msg}"))
        elif j.kind == "profit_qoq":
            entry.update(
                status="filled" if applied else "checked", changed=len(res.changed), unchanged=res.unchanged, inserted=len(j.inserted_ids),
                unmatched_rows=[], unmatched_cols=[], unused_source_rows=[], format_drift=0, notes=res.notes, grid=res.grid, overflow=[],
            )  # fmt: skip
            if j.box_ids:
                filled.setdefault(j.slide, set()).update(j.box_ids)
            if j.inserted_ids and applied:
                issues.append(ExportIssue("INFO", TABLES_LOB, kind, f"slide {j.slide}: {len(j.inserted_ids)} Profit (loss) movement QoQ box(es) inserted under the Gross vs Net table (next quarter they are updated in place)"))
            for n in res.notes:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"slide {j.slide}: {n}"))
        else:
            entry.update(
                status="filled" if applied else "checked", changed=len(res.changed), unchanged=res.unchanged, unmatched_rows=res.unmatched_rows,
                unmatched_cols=res.unmatched_cols, unused_source_rows=res.unused_source_rows, format_drift=len(res.format_drift), notes=res.notes, grid=res.grid,
                overflow=[[r, c, t] for r, c, t, _ in res.overflow],
            )  # fmt: skip
            filled.setdefault(j.slide, set()).add(j.shape_id)
            where = f"slide {j.slide}, table '{m.table_name}'"
            if res.unmatched_rows:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: no source data for row(s) {', '.join(res.unmatched_rows)}: left as they were"))
            if res.unmatched_cols:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: no source column for header(s) {', '.join(sorted(set(res.unmatched_cols)))}: left as they were"))
            for n in res.notes:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: {n}"))
            if res.format_drift and j.slide not in fresh_slides:  # a slide added from a model starts as a copy of another LoB's table: nothing to compare with
                ex = "; ".join(f"'{o}' -> '{n}'" for _, _, o, n in res.format_drift[:4])
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: {len(res.format_drift)} figure(s) are written in another number format than the template had (e.g. {ex})"))
            if res.overflow:
                ex = "; ".join(f"'{t}' in a {w:.0f} pt column" for _, _, t, w in res.overflow[:3])
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: {len(res.overflow)} figure(s) are probably too wide for their column (e.g. {ex}): widen those columns in PowerPoint before printing"))
            new_only = sorted(set(res.new_signs) - set(res.old_signs)) if (res.old_signs and j.slide not in fresh_slides) else []
            if new_only:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: signed figures are now written with the style {new_only} which the template table does not use (template: {sorted(set(res.old_signs))})"))
            if not res.changed and res.unchanged == 0:
                issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"{where}: nothing was filled"))
        checks = [c for b in j.blocks if b is not None for c in b.checks]
        if checks:
            issues.append(ExportIssue("WARN", TABLES_LOB, kind, f"the workbook's own 'checks' column shows FALSE at {', '.join(a for a, _ in checks[:6])}"))
            entry["workbook_checks_false"] = [a for a, _ in checks]
        entries.append(entry)
    return entries, issues, filled
