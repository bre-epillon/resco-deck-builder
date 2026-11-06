"""Step 4: scan -> plan -> export from Excel -> write into a copy of the template -> verify -> compare."""
from __future__ import annotations

import hashlib
import json
import logging
import shutil
import tempfile
import time
from contextlib import nullcontext
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any

from resco.compare import write_compare
from resco.dashboard import render_dashboard
from resco.runs import deck_of, summarize
from resco.config import DEFAULT_CONFIG_DIR, LADDER_KINDS, load_deck_rules, load_lobs, load_ranges, load_slides_to_add, load_table_config
from resco.deck_scan import MappedSlide, TemplateLocked, build_map, copy_for_reading, scan_deck
from resco.deck_writer import Job, apply_table_jobs, write_deck
from resco.excel_reader import ExcelSession, SelectorError
from resco.exporter import RANGE_KINDS, ExportedImage, ExportIssue, ej_capacity, export_lob, prepare_ej_template
from resco.layout import Layout, discover_layout, from_config
from resco.preflight import preflight, slot_problem
from resco.regress import coverage, describe as describe_regression, find_regressions, template_changes
from resco.selection import Selection, SelectionError
from resco.slide_adder import AddResult, add_missing_slides, missing_specs
from resco.table_plan import load_sources, local_copy, make_jobs, planned_entries, report as table_report
from resco.validate import check_table_totals, visible_error_cells
from resco.values import decode_grid
from resco.verify import compare, compare_prepared, snapshot

log = logging.getLogger(__name__)
DEFAULT_EJ_TEMPLATE = DEFAULT_CONFIG_DIR.parent / "assets" / "ej_render_template.xlsx"


def check_lob(xl: ExcelSession, lob: Any, kinds: set[str], ranges: Any, layout: Layout) -> tuple[dict[str, Any], list[ExportIssue]]:
    """Data checks for the LoB currently selected in the workbook (call before selecting the next one)."""
    sheet = ranges.selector_sheet
    out: dict[str, Any] = {"c5": "ok"}  # select() already refused anything but a matching code
    issues: list[ExportIssue] = []
    if kinds & {"table", "table_life"}:
        grid, _ = decode_grid(xl.read_grid(sheet, layout.table))
        t = check_table_totals(grid)
        out["totals"] = {"status": t.status, "comparisons": t.comparisons, "max_rel_diff": t.max_rel_diff, "failures": t.failures[:10]}
        if t.status == "fail":
            issues.append(ExportIssue("ERROR", lob.code, "table", f"totals check failed ({len(t.failures)} mismatches), e.g. {t.failures[0]}"))
        elif t.status == "n/a":
            issues.append(ExportIssue("WARN", lob.code, "table", f"totals check not applicable: {t.failures[0]}"))
    visible: list[dict[str, str]] = []
    imaged = {"table": layout.table, "table_life": layout.table_short, "profit": layout.profit, "triangles": layout.triangles}
    for kind in sorted(kinds & set(imaged)):
        grid = xl.read_grid(sheet, imaged[kind])
        hidden = xl.hidden_rows_cols(sheet, imaged[kind])
        cells = visible_error_cells(grid, imaged[kind], *hidden)
        visible += [{"kind": kind, "cell": c, "error": e} for c, e in cells]
        if cells:
            issues.append(ExportIssue("WARN", lob.code, kind, f"{len(cells)} error value(s) visible in the image: " + ", ".join(f"{c} {e}" for c, e in cells[:6])))
    out["visible_errors"] = visible
    return out, issues


def _distortion(slot_geo: tuple[float, float, float, float], size: tuple[int, int]) -> float:
    """Percent by which a picture of `size` px is stretched horizontally relative to vertically when forced into the slot."""
    return (slot_geo[2] / slot_geo[3]) / (size[0] / size[1]) * 100 - 100


def _local_copy(workbook: Path) -> tuple[Path, str]:
    """Copy the workbook to a local temp folder (never open the OneDrive original) and return (copy, sha256[:12])."""
    return local_copy(workbook, "workbook_copy")


def _target_in_scope(target_id: str | None, tcfg: Any, wanted: list[str], sel: Selection) -> bool:
    """Whether a table target belongs to this run (its LoB is requested; IBNR tables when the run is not limited to some LoBs)."""
    t = next((x for x in tcfg.targets if x.id == target_id), None)
    if t is None:
        return False
    return t.select.get("lob", "").upper() in wanted if t.kind in LADDER_KINDS else not sel.lobs


def run_build(*args: Any, **kwargs: Any) -> int:
    """Scan -> (add missing slides) -> export -> write -> verify. A temporary folder holds the template copy with the added slides."""
    with tempfile.TemporaryDirectory(prefix="resco_prep_") as tmp:
        return _run_build(*args, workdir=Path(tmp), **kwargs)


def _run_build(
    workbook: Path,
    template: Path,
    out: Path,
    config_dir: Path = DEFAULT_CONFIG_DIR,
    lobs_arg: str | None = None,
    max_year_arg: int | None = None,
    ej_template: Path = DEFAULT_EJ_TEMPLATE,
    dry_run: bool = False,
    copy_workbook: bool = True,
    picture_mode: str = "auto",
    check: bool = False,
    netting_path: Path | None = None,
    ibnr_path: Path | None = None,
    parts_arg: str | None = None,
    slides_arg: str | None = None,
    workdir: Path | None = None,
    netting_prior_path: Path | None = None,
) -> int:
    t_start = time.time()
    workbook, template, out, ej_template = (p.resolve() for p in (workbook, template, out, ej_template))
    netting_path = netting_path.resolve() if netting_path else None
    netting_prior_path = netting_prior_path.resolve() if netting_prior_path else None
    ibnr_path = ibnr_path.resolve() if ibnr_path else None
    tcfg = load_table_config(config_dir)
    rules, lobs, ranges = load_deck_rules(config_dir), load_lobs(config_dir), load_ranges(config_dir)
    try:  # a copy: the run is not disturbed when the template is saved again meanwhile, and a locked file is reported clearly
        template_copy = copy_for_reading(template, workdir) if workdir is not None else template
    except TemplateLocked as exc:
        log.error("%s", exc)
        return 2
    original_name, source_template = template.name, template
    template_modified = datetime.fromtimestamp(source_template.stat().st_mtime).isoformat(timespec="seconds")
    template = template_copy
    scans = scan_deck(template, rules)
    dm = build_map(scans, lobs, ranges.life_resclasses, rules, tcfg.targets)
    deck_titles = [sc.title or "" for sc in scans]  # the template as given (before any slide is added to the working copy)

    try:
        sel = Selection.parse(lobs_arg, parts_arg, slides_arg)
    except SelectionError as exc:
        log.error("%s", exc)
        return 2
    # slides the template lacks (config/slides_add.yaml) are added to a copy of the template; everything below works on that copy
    original_template, added = template, AddResult()  # (after the copy: original_template is the working copy, original_name the user's file)
    specs = [sp for sp in load_slides_to_add(config_dir) if sel.allows_new(sp.lob, sp.part)]
    todo = missing_specs([m.title for m in dm.mapped] + [sc.title for sc in dm.narrative], specs)
    if todo and workdir is not None:
        try:
            added = add_missing_slides(template, workdir / f"{template.stem}_with_new_slides{template.suffix}", todo, rules)
        except Exception as exc:  # noqa: BLE001 - the build goes on without the new slides
            log.exception("adding slides")
            added.entries.append({"id": "(all)", "title": "(slides to add)", "lob": "", "status": f"failed: {exc}", "position": None, "like": None})
        if added.saved:
            template = workdir / f"{original_template.stem}_with_new_slides{original_template.suffix}"
            scans = scan_deck(template, rules)
            dm = build_map(scans, lobs, ranges.life_resclasses, rules, tcfg.targets, added.kinds_by_title)
    available = {m.lob for m in dm.mapped}
    wanted = list(sel.lobs) or [lob.code for lob in lobs if lob.code in available]
    unknown = [c for c in wanted if c not in available]
    if unknown:
        log.error("no slides in the template for LoB(s): %s", unknown)
        return 2
    plan: list[MappedSlide] = sel.restrict_plan([m for m in dm.mapped if m.lob in wanted and m.selector])
    export_lobs = [c for c in wanted if any(m.lob == c for m in plan)]  # LoBs with at least one picture to export
    kinds_by_lob: dict[str, set[str]] = {}
    for m in plan:
        kinds_by_lob.setdefault(m.lob, set()).update(s.kind for s in m.slots)
    lob_by_code = {lob.code: lob for lob in lobs}
    table_matches = sel.restrict_tables(dm.table_matches, wanted)
    if sel.active and not plan and not table_matches:
        log.error("the selection (%s) matches nothing in the template", sel.describe())
        return 2
    table_lobs = [c for c in dict.fromkeys(m.target.select.get("lob", "").upper() for m in table_matches if m.target.kind in LADDER_KINDS) if c and c not in export_lobs]
    touched = export_lobs + table_lobs
    needs_excel = bool(export_lobs) or any(m.target.kind == "overview" for m in table_matches)  # netting-only runs never open Excel
    if sel.active:
        log.info("partial run: %s -> %d slide(s), %d picture slot(s), %d table(s)", sel.describe(), len({m.slide for m in plan} | {m.slide for m in table_matches}), sum(len(m.slots) for m in plan), len(table_matches))

    changes = template_changes(out, deck_of(config_dir), deck_titles)
    if changes and (changes["added"] or changes["removed"]):
        log.info("template changes since %s: +%d / -%d slide(s)", changes["since"], len(changes["added"]), len(changes["removed"]))
    from resco.serve import candidates, deck_folders  # a newer copy of the deck in the usual folders is reported

    newer = [c for c in candidates("ReservingPresentation*.pptx", deck_folders(), uk=(deck_of(config_dir) == "uk")) if c.stat().st_mtime > source_template.stat().st_mtime + 2 and c.resolve() != source_template.resolve()]
    cov = coverage(dm)  # what the template offers, compared with recent full builds of this deck
    regs = find_regressions(out, deck_of(config_dir), cov, original_name, template_modified)
    for r in regs:
        log.warning("%s", describe_regression(r))
    template_notes = [ExportIssue("WARN", "(template)", "template", f"a newer copy of the deck exists: {c.name} saved {datetime.fromtimestamp(c.stat().st_mtime):%d %b %H:%M} in {c.parent} (this run uses {source_template.name} saved {template_modified})") for c in newer[:2]]
    if changes and (changes["added"] or changes["removed"]):
        template_notes.append(ExportIssue("INFO", "(template)", "template", f"template changes since {changes['since']}: " + "; ".join(([f"added {t!r}" for t in changes["added"][:6]]) + ([f"removed {t!r}" for t in changes["removed"][:6]]))))
    out_dir = out / f"build_{datetime.now():%Y%m%d_%H%M%S}"
    out_dir.mkdir(parents=True, exist_ok=True)
    source_workbook = workbook
    workbook_sha: str | None = None
    if copy_workbook and not dry_run and needs_excel:
        workbook, workbook_sha = _local_copy(source_workbook)
        log.info("working on a local copy of %s", source_workbook.name)
    result: dict[str, Any] = {
        "kind": "check" if check else ("dry-run" if dry_run else "build"),
        "deck": deck_of(config_dir),
        "created": datetime.now().isoformat(timespec="seconds"),
        "template": original_name,
        "template_path": str(source_template),
        "template_modified": template_modified,
        "newer_templates": [str(c) for c in newer[:3]],
        "coverage": cov,
        "deck_titles": deck_titles,
        "template_changes": changes,
        "regressions": regs,
        "workbook": source_workbook.name,
        "workbook_sha256": workbook_sha,
        "dry_run": dry_run,
        "lobs": touched,
        "selection": sel.as_dict(),
        "added_slides": [e for e in added.entries if e["status"] != "present"],
        "lobs_without_slides": [] if sel.active else dm.lobs_without_slides,
        "scan_flags": [asdict(f) for f in dm.flags if f.lob in touched or ("table target" in f.message and _target_in_scope(f.lob, tcfg, wanted, sel))],
        "slides": [],
        "issues": [],
        "tables": [],
    }

    if dry_run:
        for m in plan:
            result["slides"].append(
                {
                    "slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "title_type": m.title_type,
                    "slots": [{"kind": s.kind, "shape_id": s.shape_id, "source": rules.kinds[s.kind].source, "status": "planned", **({"new": True} if s.new else {})} for s in m.slots],
                }
            )  # fmt: skip
        result["tables"] = planned_entries(table_matches)
        result["issues"] = [asdict(ExportIssue("WARN", "(regression)", "template", describe_regression(r))) for r in regs] + [asdict(i) for i in template_notes]
        _write_outputs(out_dir, result, dry_run=True, seconds=time.time() - t_start)
        log.info("dry run: %d slides, %d slots planned -> %s", len(plan), sum(len(m.slots) for m in plan), out_dir)
        return 0

    # ---- Excel side ------------------------------------------------------------------------------------
    images: dict[str, dict[str, ExportedImage]] = {}
    issues: list[ExportIssue] = []
    issues += [ExportIssue("WARN", "(regression)", "template", describe_regression(r)) for r in regs] + template_notes
    for e in result["added_slides"]:
        detail = f"slide '{e['title']}': {e['status']}" + (f" at position {e['position']} (copy of '{e['like']}')" if e.get("position") else "")
        issues.append(ExportIssue("WARN" if e["status"].startswith("failed") else "INFO", e.get("lob") or "(deck)", "slides", detail + (f"; {e['notes']}" if e.get("notes") else "")))
    seconds_per_lob: dict[str, float] = {}
    checks: dict[str, dict[str, Any]] = {}
    max_year, layout, pre, ej_book = 0, from_config(ranges), [], None
    with (ExcelSession(workbook, picture_mode=picture_mode) if needs_excel else nullcontext()) as xl:
        if export_lobs:
            max_year = max_year_arg or ranges.ej_max_year or xl.asof_year()
            ej_book = None
            if ranges.hide_rows:
                xl.hide_rows(ranges.selector_sheet, ranges.hide_rows)  # in memory only: the workbook copy is never saved

            # where things are in THIS workbook version: blocks are found by their labels / LoB codes, not by fixed addresses
            layout = from_config(ranges)
            for code in export_lobs:  # the sheet labels are the same for every LoB: the first one that selects cleanly will do
                try:
                    xl.select(ranges.selector_sheet, ranges.selector_cell, ranges.check_cell, lob_by_code[code].selector, lob_by_code[code].expected_check)
                    layout = discover_layout(xl, ranges, max_year)
                    break
                except SelectorError as exc:
                    log.warning("layout: %s did not select (%s), trying the next LoB", code, exc)
            else:
                issues.append(ExportIssue("ERROR", "(workbook)", "layout", "no LoB could be selected, so the sheet layout could not be checked"))
            result["layout"] = layout.as_report()
            issues += [ExportIssue(level, "(workbook)", "layout", msg) for level, msg in layout.notes if level != "INFO"]
            pre = preflight(xl, ranges, layout, lob_by_code, kinds_by_lob)
            ej_lobs = [lob_by_code[c].resclass for c, ks in kinds_by_lob.items() if {"ej", "ej_life"} & ks]
            if not check and ej_lobs:
                needed = max([layout.ej[r].n_rows for r in ej_lobs if r in layout.ej] or [ej_capacity(ranges)])
                ej_book = prepare_ej_template(xl, ej_template, ranges, needed)
        sources = load_sources(table_matches, tcfg, netting_path, ibnr_path, xl, netting_prior_path)
        table_jobs = make_jobs(table_matches, sources)
        result["table_sources"] = sources.info
        if check:
            issues += pre  # in a build the same problems surface slot by slot, so they are not listed twice
            for m in plan:
                slide_entry: dict[str, Any] = {"slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "title_type": m.title_type, "slots": []}
                for slot in m.slots:
                    problem = slot_problem(issues, m.lob, slot.kind)
                    slide_entry["slots"].append({"kind": slot.kind, "shape_id": slot.shape_id, "source": rules.kinds[slot.kind].source, "status": f"failed: {problem}" if problem else "found"})
                result["slides"].append(slide_entry)
            if table_jobs:  # try the fill on the template in memory: reports unmatched rows / columns without writing anything
                from pptx import Presentation

                apply_table_jobs(Presentation(str(template)), table_jobs)
                entries, t_issues, _ = table_report(table_jobs, table_matches, applied=False)
                result["tables"] = entries
                issues += t_issues
                for e in entries:
                    problem = e["status"] if e["status"].startswith("failed") else None
                    log.info("table %s (slide %s): %s", e["id"], e["slide"], problem or f"ok, {e.get('changed', 0)} figure(s) would change")
            result["issues"] = [asdict(i) for i in issues]
            _write_outputs(out_dir, result, dry_run=False, seconds=time.time() - t_start)
            n_err = sum(1 for i in issues if i.level == "ERROR")
            log.info("workbook check: %d error(s), %d warning(s) -> %s", n_err, sum(1 for i in issues if i.level == "WARN"), out_dir)
            return 1 if n_err else 0
        for n, code in enumerate(export_lobs, start=1):
            t0 = time.time()
            log.info("PROGRESS %d/%d %s", n, len(export_lobs), code)
            try:
                images[code], lob_issues = export_lob(xl, lob_by_code[code], kinds_by_lob[code], ranges, layout, out_dir / "images", max_year, ej_book)
                issues += lob_issues
                checks[code], check_issues = check_lob(xl, lob_by_code[code], kinds_by_lob[code], ranges, layout)
                issues += check_issues
            except SelectorError as exc:
                log.error("%s: %s", code, exc)
                issues.append(ExportIssue("ERROR", code, "*", str(exc)))
                images.setdefault(code, {})
                checks[code] = {"c5": "FAILED"}
            except Exception as exc:  # noqa: BLE001 - keep going with the other LoBs
                log.exception("%s: unexpected error", code)
                issues.append(ExportIssue("ERROR", code, "*", f"unexpected error: {exc}"))
                images.setdefault(code, {})
                checks.setdefault(code, {"c5": "unknown"})
            seconds_per_lob[code] = round(time.time() - t0, 1)
            log.info("%s exported (%d images) in %.1fs", code, len(images[code]), seconds_per_lob[code])
        # second pass: anything that failed (typically a busy clipboard) gets one more go after a short pause
        for code in export_lobs:
            missing = kinds_by_lob.get(code, set()) - set(images.get(code, {}))
            if not missing or checks.get(code, {}).get("c5") == "FAILED":
                continue
            log.warning("%s: retrying %s", code, sorted(missing))
            time.sleep(5)
            try:
                more, more_issues = export_lob(xl, lob_by_code[code], missing, ranges, layout, out_dir / "images", max_year, ej_book)
            except Exception:  # noqa: BLE001
                log.exception("%s: retry failed", code)
                continue
            images[code].update(more)
            issues[:] = [i for i in issues if not (i.lob == code and i.kind in more and i.level == "ERROR")]
            issues += more_issues
        if ej_book is not None:
            ej_book.book.close()

    # ---- deck side -------------------------------------------------------------------------------------
    jobs: list[Job] = []
    slot_status: dict[tuple[int, int], str] = {}
    for m in plan:
        for s in m.slots:
            img = images.get(m.lob, {}).get(s.kind)
            if img is None:
                slot_status[(m.slide, s.shape_id)] = "failed: no image"
                continue
            jobs.append(Job(m.slide, m.title, s.shape_id, s.kind, m.lob, img.path, new=s.new, geometry=s.geometry if s.new else None))
            slot_status[(m.slide, s.shape_id)] = "inserted" if s.new else "replaced"

    built = out_dir / f"{Path(original_name).stem}_auto{Path(original_name).suffix}"
    job_results = write_deck(template, built, jobs, table_jobs) if (jobs or table_jobs) else []
    table_entries, table_issues, filled = table_report(table_jobs, table_matches, fresh_slides=frozenset(e["position"] for e in added.entries if e.get("position")))
    result["tables"] = table_entries
    issues += table_issues
    by_key = {(r.job.slide, r.job.shape_id): r for r in job_results}

    rows: list[dict[str, Any]] = []
    stretch: dict[tuple[str, int, int], list[tuple[int, str]]] = {}
    for m in plan:
        slide_entry: dict[str, Any] = {"slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "title_type": m.title_type, "slots": []}
        for s in m.slots:
            r = by_key.get((m.slide, s.shape_id))
            entry: dict[str, Any] = {"kind": s.kind, "shape_id": s.shape_id, "status": slot_status[(m.slide, s.shape_id)]}
            if r and r.old_size is None:  # inserted into an empty slide: nothing to compare with
                d_new = _distortion(s.geometry, r.new_size)
                entry.update(shape_id=r.shape_id, new_size=list(r.new_size), distortion_new=round(d_new, 1), inserted=True)
                rows.append({"slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "kind": s.kind, "shape_id": r.shape_id,
                             "old_size": None, "new_size": r.new_size, "distortion_old": None, "distortion_new": d_new, "inserted": True})  # fmt: skip
            elif r:
                d_old, d_new = _distortion(s.geometry, r.old_size), _distortion(s.geometry, r.new_size)
                entry.update(old_size=list(r.old_size), new_size=list(r.new_size), distortion_old=round(d_old, 1), distortion_new=round(d_new, 1))
                rows.append({"slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "kind": s.kind, "shape_id": s.shape_id,
                             "old_size": r.old_size, "new_size": r.new_size, "distortion_old": d_old, "distortion_new": d_new})  # fmt: skip
                if abs(d_new - d_old) > 10:
                    stretch.setdefault((s.kind, round(d_old), round(d_new)), []).append((m.slide, m.lob))
            slide_entry["slots"].append(entry)
        result["slides"].append(slide_entry)

    methods = [img.method for per_lob in images.values() for img in per_lob.values() if img.kind in RANGE_KINDS]
    result["picture_methods"] = {m: methods.count(m) for m in sorted(set(methods))}
    if "pdf" in methods:
        issues.append(ExportIssue("WARN", "(multiple)", "pictures", f"{methods.count('pdf')} of {len(methods)} range pictures were rendered through the PDF route because CopyPicture was unavailable (no live display or busy clipboard). They are sharper and use print metrics, so they look slightly different from the usual screen pictures."))
    for (kind, d_old, d_new), where in stretch.items():
        what = f"the new image has a different aspect ratio from the one in the template (stretch in the slot {d_old:+d}% -> {d_new:+d}%)"
        if len(where) >= 3:  # one finding, not one line per slide: the source block itself changed
            codes = ", ".join(sorted({lob for _, lob in where}))
            issues.append(ExportIssue("WARN", "(multiple)", kind, f"{len(where)} slides: {what}. LoBs: {codes}"))
        else:
            issues += [ExportIssue("WARN", lob, kind, f"slide {slide}: {what}") for slide, lob in where]

    # ---- verification ----------------------------------------------------------------------------------
    replaced: dict[int, set[int]] = {}
    inserted: dict[int, int] = {}
    for j in jobs:
        if j.new:
            inserted[j.slide] = inserted.get(j.slide, 0) + 1
        else:
            replaced.setdefault(j.slide, set()).add(j.shape_id)
    inserted_tables: dict[int, int] = {}
    for tj in table_jobs:  # the Profit (loss) movement QoQ boxes added under a Gross vs Net table
        if tj.inserted_ids and not tj.error:
            inserted_tables[tj.slide] = inserted_tables.get(tj.slide, 0) + len(tj.inserted_ids)
    ver = compare(snapshot(template), snapshot(built), replaced, filled, inserted, inserted_tables) if (jobs or filled or inserted_tables) else None
    if ver is not None and added.saved:  # the added slides must not have changed any original slide
        ver.messages += compare_prepared(snapshot(original_template), snapshot(template), added.added_positions)
        ver.stats["slides_added"] = added.added
    result["verify"] = {"ok": ver.ok, "stats": ver.stats, "messages": ver.messages} if ver else None
    result["issues"] = [asdict(i) for i in issues]
    result["seconds_per_lob"] = seconds_per_lob
    result["checks"] = checks
    result["output_deck"] = built.name

    if jobs:
        compare_path = write_compare(template, built, rows, out_dir, original_name)
        result["compare_html"] = str(compare_path.relative_to(out_dir))
    _write_outputs(out_dir, result, dry_run=False, seconds=time.time() - t_start)

    data_error = lambda i: i.level == "ERROR" and i.message.startswith("totals check failed")  # noqa: E731  workbook data, not the tool
    hard = any(i.level == "ERROR" and not data_error(i) for i in issues) or any(v.startswith("failed") for v in slot_status.values()) or (ver is not None and not ver.ok) or any(e["status"].startswith("failed") for e in table_entries)
    soft = any(data_error(i) for i in issues)
    log.info("built %s: %d pictures replaced, verify %s -> %s", built.name, len(jobs), "n/a" if ver is None else "OK" if ver.ok else "FAILED", out_dir)
    return 1 if hard else 3 if soft else 0  # 0 ok, 3 built but the workbook data checks reported errors, 1 failure


def _write_outputs(out_dir: Path, result: dict[str, Any], dry_run: bool, seconds: float) -> None:
    result["seconds"] = round(seconds, 1)
    result["summary"] = summarize(result)
    (out_dir / "build_result.json").write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    heading = {"check": "Workbook check (nothing was exported or written)", "dry-run": "Dry run (nothing written to Excel or the deck)"}.get(result["kind"], "")
    L = [f"# Build report{': ' + heading if heading else ''}", ""]
    L += [f"- Template: `{result['template']}`; workbook: `{result['workbook']}`; LoBs: {', '.join(result['lobs'])}"]
    if result.get("added_slides"):
        L += ["", "## Slides added to the deck (not in the template)", "", "| Slide | Title | Copy of | After | Status |", "|---|---|---|---|---|"]
        L += [f"| {e.get('position') or '-'} | {e['title']} | {e.get('like') or '-'} | {e.get('after') or '-'} | {e['status']} |" for e in result["added_slides"]]
        L += [""]
    narrowed = result.get("selection") or {}
    if narrowed.get("lobs") or narrowed.get("parts") or narrowed.get("slides"):
        L += [f"- **Partial run**, only: {narrowed['text']}. Everything else in the deck is left exactly as the template had it."]
    n_slots = sum(len(s["slots"]) for s in result["slides"])
    L += [f"- Slides in scope: {len(result['slides'])}; picture slots: {n_slots}; run time: {result['seconds']} s"]
    if result["kind"] == "build":
        done = sum(1 for s in result["slides"] for x in s["slots"] if x["status"] in ("replaced", "inserted"))
        n_new = sum(1 for s in result["slides"] for x in s["slots"] if x["status"] == "inserted")
        L += [f"- Pictures replaced: {done} of {n_slots}" + (f" ({n_new} inserted into empty slides)" if n_new else "") + f"; output deck: `{result.get('output_deck', '-')}`"]
        v = result.get("verify")
        if v:
            L += [f"- Verification: {'PASSED' if v['ok'] else 'FAILED'} {v['stats']}"]
        if result.get("compare_html"):
            L += [f"- Old-vs-new comparison: `{result['compare_html']}`"]
    L += ["", "## Slides", "", "| Slide | LoB / type | Slot | Status | Old px | New px | Stretch old / new |", "|---|---|---|---|---|---|---|"]
    for s in result["slides"]:
        for x in s["slots"]:
            old = "x".join(map(str, x["old_size"])) if "old_size" in x else "-"
            new = "x".join(map(str, x["new_size"])) if "new_size" in x else "-"
            st = f"{x['distortion_old']:+.0f}% / {x['distortion_new']:+.0f}%" if "distortion_old" in x else x.get("source", "-")
            note = "" if s["type"] == s["title_type"] else f" (title says {s['title_type']})"
            L.append(f"| {s['slide']} | {s['lob']} / {s['type']}{note} | {x['kind']} | {x['status']} | {old} | {new} | {st} |")
    if result.get("tables"):
        L += ["", "## Tables filled from Excel (native slide tables, text only)", "", "| Slide | Table | Source | Status | Figures changed | Unchanged | Unmatched rows / columns |", "|---|---|---|---|---|---|---|"]
        for t in result["tables"]:
            un = ", ".join(t.get("unmatched_rows", []) + t.get("unmatched_cols", [])) or "-"
            L.append(f"| {t['slide']} | {t['id']} ({t['table']}) | {t['source']} | {t['status']} | {t.get('changed', '-')} | {t.get('unchanged', '-')} | {un} |")
    lay = result.get("layout")
    if lay:
        L += ["", "## Where things were found in the workbook", "", "| What | Address used | Configured | |", "|---|---|---|---|"]
        for key, label in (("table", "Results table"), ("table_short", "Results table, first block only (Life)"), ("profit", "P&L box"), ("triangles", "Triangles")):
            it = lay[key]
            L.append(f"| {label} | {it['address']} | {it['configured']} | {'MOVED' if it['moved'] else 'as configured'} |")
        L += ["", "Expert Judgement blocks (found by LoB code; rows used): " + ", ".join(f"{c} {b['rows'][0]}-{b['rows'][1]}" for c, b in lay["expert_judgement"].items())]
    if result.get("checks"):
        L += ["", "## Data checks per LoB", "", "| LoB | LoB code on sheet | Totals vs components | Max rel. diff | Error values visible in images |", "|---|---|---|---|---|"]
        for code, c in result["checks"].items():
            t = c.get("totals", {})
            L.append(
                f"| {code} | {c.get('c5', '-')} | {t.get('status', '-')} ({t.get('comparisons', 0)} comparisons) | "
                f"{t.get('max_rel_diff', 0):.1e} | {len(c.get('visible_errors', []))} |"
            )
    L += ["", "## Issues", ""]
    items = result["issues"] + [{"level": f["level"], "lob": f["lob"], "kind": "scan", "message": f["message"]} for f in result["scan_flags"] if f["level"] != "INFO"]
    for lvl in ("ERROR", "WARN", "INFO"):
        sel = [i for i in items if i["level"] == lvl]
        L += [f"- **{lvl}** [{i['lob']}/{i['kind']}] {i['message']}" for i in sel]
    if not items:
        L.append("- none")
    if result.get("verify") and result["verify"]["messages"]:
        L += ["", "## Verification messages", ""] + [f"- {lvl}: {m}" for lvl, m in result["verify"]["messages"]]
    (out_dir / "build_report.md").write_text("\n".join(L) + "\n", encoding="utf-8")
    render_dashboard(out_dir, result)
