from __future__ import annotations

import argparse
import csv
import logging
from datetime import datetime
from pathlib import Path

from resco.config import DEFAULT_CONFIG_DIR, load_lobs, load_ranges
from resco.dump import dump_all, write_manifest
from resco.excel_reader import ExcelSession

log = logging.getLogger("resco")


def _cmd_dump(args: argparse.Namespace) -> int:
    lobs = load_lobs(args.config)
    ranges = load_ranges(args.config)
    if args.lobs:
        wanted = {c.strip().upper() for c in args.lobs.split(",")}
        lobs = [lob for lob in lobs if lob.code.upper() in wanted]
        unknown = wanted - {lob.code.upper() for lob in lobs}
        if unknown:
            log.error("unknown LoB codes: %s", sorted(unknown))
            return 2
    out_dir = args.out / f"dump_{datetime.now():%Y%m%d_%H%M%S}"
    out_dir.mkdir(parents=True, exist_ok=True)

    with ExcelSession(args.workbook, visible=args.visible) as xl:
        max_year = args.max_year or xl.asof_year()
        log.info("max year %s, %d LoBs", max_year, len(lobs))
        from resco.layout import discover_layout

        xl.select(ranges.selector_sheet, ranges.selector_cell, ranges.check_cell, lobs[0].selector, lobs[0].expected_check)
        res = dump_all(xl, lobs, ranges, out_dir, max_year, discover_layout(xl, ranges, max_year))

        # chart inventory: which chart objects exist under which names (legacy macro addresses them by name)
        inv: list[dict] = []
        sheets = [ranges.selector_sheet] + sorted({lob.resclass for lob in lobs if lob.resclass})
        for sheet in sheets:
            if xl.has_sheet(sheet):
                inv += xl.chart_inventory(sheet)
            else:
                res.failures.append((sheet, "sheet not found for chart inventory"))
    if inv:
        with (out_dir / "charts.csv").open("w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=list(inv[0]))
            w.writeheader()
            w.writerows(inv)
    write_manifest(res)
    _write_report(res, inv, lobs, ranges, out_dir, max_year)
    log.info("report: %s", out_dir / "dump_report.md")
    return 1 if res.failures else 0


def _write_report(res, inv, lobs, ranges, out_dir: Path, max_year: int) -> None:
    ok = {b.lob for b in res.blocks}
    lines = [
        "# Dump report (Step 2)",
        "",
        f"- LoBs requested: {len(lobs)}; read OK: {len(ok)}; failed: {len(res.failures)}",
        f"- Max year used for Expert Judgement block length: {max_year}",
        f"- Blocks written: {len(res.blocks)}; error cells recorded: {len(res.errors)} "
        f"(of which in hidden rows/columns, i.e. not rendered in images: {sum(1 for e in res.errors if e[4])})",
        "",
        "## Failures",
        *(f"- {c}: {m}" for c, m in res.failures),
        *([] if res.failures else ["- none"]),
        "",
        "## Expert Judgement checks",
        *(f"- {m}" for _, m in res.ej_checks),
        *([] if res.ej_checks else ["- none"]),
        "",
        "## Error cells per LoB / block",
    ]
    visible: dict[tuple[str, str], list[str]] = {}
    for lob, blk, cell, _text, hidden in res.errors:
        if not hidden:
            visible.setdefault((lob, blk), []).append(cell)
    lines += [f"- VISIBLE {lob} / {blk}: {', '.join(cells)}" for (lob, blk), cells in sorted(visible.items())] or [
        "- none visible (all error cells are in hidden rows/columns)"
    ]

    lines += ["", "## Chart objects"]
    by_sheet: dict[str, list[dict]] = {}
    for row in inv:
        by_sheet.setdefault(row["sheet"], []).append(row)
    need = set(ranges.chart_names_triangles)
    have = {r["name"] for r in by_sheet.get(ranges.selector_sheet, [])}
    lines.append(f"- {ranges.selector_sheet}: legacy needs {sorted(need)}; missing: {sorted(need - have) or 'none'}")
    lines.append("- Pattern charts (legacy finds them by NAME; by TITLE is what identifies them reliably). Only deviations listed:")
    for lob in lobs:
        if lob.patterns and lob.resclass:
            charts = by_sheet.get(lob.resclass, [])
            for name, title in zip(ranges.chart_names_patterns, ranges.chart_titles_patterns):
                n_name = sum(1 for c in charts if c["name"] == name)
                n_title = sum(1 for c in charts if c["title"] == title)
                if n_name != 1 or n_title != 1:
                    lines.append(f"  - {lob.resclass} '{title}': {n_title} chart(s) by title, {n_name} named '{name}'")
    lines += ["", "## Timing (s per LoB)", ", ".join(f"{k} {v}" for k, v in res.seconds_per_lob.items())]
    (out_dir / "dump_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def _cmd_scan(args: argparse.Namespace) -> int:
    import yaml  # local import: only this command writes YAML

    from resco.config import load_deck_rules
    from resco.deck_scan import build_map, matrix, render_report, scan_deck, to_yaml_dict

    import time as _time

    t0 = _time.time()
    rules = load_deck_rules(args.config)
    lobs = load_lobs(args.config)
    ranges = load_ranges(args.config)
    import tempfile

    from resco.deck_scan import TemplateLocked, copy_for_reading

    with tempfile.TemporaryDirectory(prefix="resco_scan_") as tmp:
        try:
            scans = scan_deck(copy_for_reading(args.template, Path(tmp)), rules)
        except TemplateLocked as exc:
            log.error("%s", exc)
            return 2
    from resco.config import load_table_config
    from resco.table_plan import planned_entries

    tcfg = load_table_config(args.config)
    dm = build_map(scans, lobs, ranges.life_resclasses, rules, tcfg.targets)
    from resco.config import load_slides_to_add
    from resco.slide_adder import missing_specs

    to_add = [sp.title for sp in missing_specs([m.title for m in dm.mapped] + [sc.title for sc in dm.narrative], load_slides_to_add(args.config))]

    out_dir = args.out / f"scan_{datetime.now():%Y%m%d_%H%M%S}"
    out_dir.mkdir(parents=True, exist_ok=True)
    baseline = yaml.safe_dump(to_yaml_dict(dm, args.template.name), sort_keys=False, allow_unicode=True, width=140)
    (out_dir / "slides.yaml").write_text(baseline, encoding="utf-8")
    if args.update_baseline:
        (args.config / "slides.yaml").write_text(baseline, encoding="utf-8")
        log.info("reference slide map updated: %s", args.config / "slides.yaml")
    report_text = render_report(dm, args.template.name, len(scans))
    if to_add:
        report_text += "\n## Slides a build adds (not in this template)\n\n" + "\n".join(f"- {t}" for t in to_add) + "\n"
    (out_dir / "scan_report.md").write_text(report_text, encoding="utf-8")

    from dataclasses import asdict
    import json

    from resco.dashboard import render_dashboard
    from resco.runs import deck_of, summarize

    result = {
        "kind": "scan", "deck": deck_of(args.config), "created": datetime.now().isoformat(timespec="seconds"),
        "template": args.template.name, "workbook": "-", "dry_run": False, "lobs": sorted({m.lob for m in dm.mapped}),
        "scan_flags": [asdict(f) for f in dm.flags], "issues": [], "tables": planned_entries(dm.table_matches), "lobs_without_slides": dm.lobs_without_slides, "slides_to_add": to_add, "other_slides": len(dm.narrative), "seconds": round(_time.time() - t0, 1),
        "slides": [
            {"slide": m.slide, "title": m.title, "lob": m.lob, "type": m.type, "title_type": m.title_type,
             "slots": [{"kind": x.kind, "shape_id": x.shape_id, "status": "recognised", "source": rules.kinds[x.kind].source, **({"new": True} if x.new else {})} for x in m.slots]}
            for m in dm.mapped
        ],
    }  # fmt: skip
    result["summary"] = summarize(result)
    (out_dir / "scan_result.json").write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    render_dashboard(out_dir, result)

    types, rows = matrix(dm)
    log.info("%d slides, %d LoB slides, %d other", len(scans), len(dm.mapped), len(dm.narrative))
    for lvl in ("ERROR", "WARN", "INFO"):
        log.info("%s flags: %d", lvl, sum(1 for f in dm.flags if f.level == lvl))
    log.info("written: %s", out_dir)
    return 1 if any(f.level == "ERROR" for f in dm.flags) else 0


def _cmd_build(args: argparse.Namespace) -> int:
    from resco.build import DEFAULT_EJ_TEMPLATE, run_build

    return run_build(
        workbook=args.workbook,
        template=args.template,
        out=args.out,
        config_dir=args.config,
        lobs_arg=args.lobs,
        max_year_arg=args.max_year,
        ej_template=args.ej_template or DEFAULT_EJ_TEMPLATE,
        dry_run=args.dry_run,
        picture_mode=args.picture_mode,
        check=args.check,
        netting_path=args.netting,
        ibnr_path=args.ibnr,
        parts_arg=args.parts,
        slides_arg=args.slides,
    )


def _cmd_serve(args: argparse.Namespace) -> int:
    from resco.serve import PROJECT_ROOT, candidates, deck_folders, serve

    # the input files normally sit in Inputs/; decks are also looked for in the project folder and the one above it.
    # The NEWEST file wins (by modification time), so an updated template is picked up whichever folder it was saved in.
    inputs = [PROJECT_ROOT / "Inputs", PROJECT_ROOT]

    def newest(pattern: str, where: list[Path], uk: bool | None = None) -> str:
        found = candidates(pattern, where, uk)
        return str(found[0]) if found else ""

    defaults = {
        "workbook": str(args.workbook) if args.workbook else newest("ABE*MBE*.xlsm", inputs),
        "template": str(args.template) if args.template else newest("ReservingPresentation*.pptx", deck_folders(), uk=False),
        "netting": newest("netting*.xlsx", [*inputs, PROJECT_ROOT.parent.parent / "02_ABE&MBE_Proposals"]),
        "ibnr": newest("IBNR_Allocation*.xlsm", inputs),
        "uk_template": newest("ReservingPresentation*UK*.pptx", deck_folders(), uk=True),
    }
    return serve(args.out.resolve(), args.port, defaults, open_browser=not args.no_browser)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="resco")
    sub = p.add_subparsers(dest="cmd", required=True)
    v = sub.add_parser("serve", help="Open the local dashboard (with run controls) in the browser")
    v.add_argument("--out", type=Path, default=Path("output"))
    v.add_argument("--port", type=int, default=8765)
    v.add_argument("--workbook", type=Path, help="default workbook in the run form")
    v.add_argument("--template", type=Path, help="default template deck in the run form")
    v.add_argument("--no-browser", action="store_true")
    v.set_defaults(func=_cmd_serve)
    b = sub.add_parser("build", help="Step 4: export images from Excel and write them into a copy of the template deck")
    b.add_argument("--workbook", type=Path, required=True)
    b.add_argument("--template", type=Path, required=True)
    b.add_argument("--out", type=Path, default=Path("output"))
    b.add_argument("--config", type=Path, default=DEFAULT_CONFIG_DIR)
    b.add_argument("--lobs", help="comma separated LoB codes, e.g. ENUPS,ALL (default: every LoB with slides in the template)")
    b.add_argument("--parts", help="only these content types: results, triangles, patterns, ej, netting (Gross vs Net tables), ibnr (IBNR tables); comma separated (default: all)")
    b.add_argument("--slides", help="only these slide numbers of the template, e.g. 16,17 or 16-25 (default: all)")
    b.add_argument("--max-year", type=int, help="override; default = year of the 'asof' named range")
    b.add_argument("--ej-template", type=Path, help="formatted Expert Judgement render workbook (default: assets/)")
    b.add_argument("--picture-mode", choices=["auto", "clipboard", "pdf"], default="auto", help="how range pictures are made: auto = CopyPicture, PDF route if it fails")
    b.add_argument("--netting", type=Path, help="netting workbook (sheet 'Ladder'): fills the Gross vs Net summary tables")
    b.add_argument("--ibnr", type=Path, help="IBNR allocation workbook (sheet 'Overview'): fills the Earned IBNR allocation / QoQ movement tables")
    b.add_argument("--check", action="store_true", help="check the workbook structure and the deck only: opens a copy of the workbook (about 40 s), exports nothing")
    b.add_argument("--dry-run", action="store_true", help="scan and print the plan only: no Excel, no deck written")
    b.set_defaults(func=_cmd_build)
    s = sub.add_parser("scan", help="Step 3: scan the template deck, map (LoB, slide type) -> picture slots")
    s.add_argument("--template", type=Path, required=True)
    s.add_argument("--out", type=Path, default=Path("output"))
    s.add_argument("--config", type=Path, default=DEFAULT_CONFIG_DIR)
    s.add_argument("--update-baseline", action="store_true", help="also rewrite config/slides.yaml (the reference slide map) from this template")
    s.set_defaults(func=_cmd_scan)
    d = sub.add_parser("dump", help="Step 2: loop over LoBs and dump every block to CSV (no PowerPoint)")
    d.add_argument("--workbook", type=Path, required=True)
    d.add_argument("--out", type=Path, default=Path("output"))
    d.add_argument("--config", type=Path, default=DEFAULT_CONFIG_DIR)
    d.add_argument("--lobs", help="comma separated LoB codes, e.g. ENUPS,ALL (default: all in lobs.yaml)")
    d.add_argument("--max-year", type=int, help="override; default = year of the 'asof' named range")
    d.add_argument("--visible", action="store_true", help="show the Excel window (debugging)")
    d.set_defaults(func=_cmd_dump)
    args = p.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    return args.func(args)
