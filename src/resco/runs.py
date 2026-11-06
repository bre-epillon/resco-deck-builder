"""What a run is: its deck (Lux or UK), its kind (scan / dry run / build) and a one-line summary.
Shared by the build, the scan, the dashboard pages and the local server so they never disagree."""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

DECK_LABEL = {"main": "Lux (main deck)", "uk": "UK branch deck"}
KIND_LABEL = {"build": "Build", "dry-run": "Dry run", "scan": "Template scan", "check": "Workbook check"}
_UK_IN_NAME = re.compile(r"(?<![A-Za-z])UK(?![A-Za-z])", re.IGNORECASE)
_STAMP = re.compile(r"_(\d{8})_(\d{6})$")


def deck_of(config_dir: Path) -> str:
    """The UK deck has its own config folder (config/uk); everything else is the Lux (main) deck."""
    return "uk" if Path(config_dir).name.lower() == "uk" else "main"


def name_says_uk(template_name: str) -> bool:
    return bool(_UK_IN_NAME.search(Path(template_name).stem))


def run_kind(result: dict[str, Any]) -> str:
    return result.get("kind") or ("dry-run" if result.get("dry_run") else "build")


def run_deck(result: dict[str, Any]) -> str:
    return result.get("deck") or ("uk" if name_says_uk(result.get("template", "")) else "main")


def created_of(run_dir: Path, result: dict[str, Any] | None = None) -> str:
    if result and result.get("created"):
        return str(result["created"])
    m = _STAMP.search(run_dir.name)
    if m:
        return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S").isoformat(timespec="seconds")
    return datetime.fromtimestamp(run_dir.stat().st_mtime).isoformat(timespec="seconds")


def enrich(result: dict[str, Any], run_dir: Path) -> dict[str, Any]:
    """Fill kind / deck / created for runs written before these fields existed (does not touch the file on disk)."""
    out = dict(result)
    out["kind"], out["deck"], out["created"] = run_kind(result), run_deck(result), created_of(run_dir, result)
    return out


def summarize(result: dict[str, Any]) -> dict[str, Any]:
    kind = run_kind(result)
    issues = list(result.get("issues", [])) + [f for f in result.get("scan_flags", []) if f.get("level") != "INFO"]
    verify = result.get("verify")
    verify_errors = [m for lvl, m in (verify or {}).get("messages", []) if lvl == "ERROR"]
    errors = sum(1 for i in issues if i.get("level") == "ERROR") + len(verify_errors)
    warnings = sum(1 for i in issues if i.get("level") == "WARN")
    data_errors = sum(1 for i in issues if i.get("level") == "ERROR" and str(i.get("message", "")).startswith("totals check failed"))
    slots = [x for s in result.get("slides", []) for x in s.get("slots", [])]
    replaced = sum(1 for x in slots if x.get("status") in ("replaced", "inserted"))
    failed = sum(1 for x in slots if str(x.get("status", "")).startswith("failed"))
    verify_failed = bool(verify) and not verify.get("ok", True)

    if kind == "dry-run":
        status, label = "none", "Dry run"
    elif errors or failed or verify_failed:
        hard = errors > data_errors or failed or verify_failed
        status = "err"
        label = {"scan": "Problems in the deck", "check": "Problems found in the workbook"}.get(kind, "Needs attention") if hard else "Deck built: workbook data checks failed"
    elif warnings:
        status, label = "warn", "Review warnings"
    else:
        status, label = "ok", {"scan": "Deck recognised", "check": "Workbook ready"}.get(kind, "All good")
    sel = result.get("selection") or {}
    tables = result.get("tables", [])
    return {
        "status": status, "label": label, "slides": len(result.get("slides", [])), "slots": len(slots),
        "replaced": replaced, "failed": failed, "errors": errors, "warnings": warnings,
        "tables": len(tables), "tables_filled": sum(1 for t in tables if t.get("status") == "filled"),
        "partial": sel.get("text", "") if (sel.get("lobs") or sel.get("parts") or sel.get("slides")) else "",
    }  # fmt: skip


def describe_run(run_dir: Path) -> dict[str, Any] | None:
    """One row of the run history, or None when the folder is not a finished run."""
    for fname in ("build_result.json", "scan_result.json"):
        f = run_dir / fname
        if f.is_file():
            try:
                result = json.loads(f.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                return None
            r = enrich(result, run_dir)
            return {
                "name": run_dir.name, "kind": r["kind"], "deck": r["deck"], "deck_label": DECK_LABEL[r["deck"]],
                "kind_label": KIND_LABEL[r["kind"]], "created": r["created"], "template": r.get("template", ""),
                "has_dashboard": (run_dir / "dashboard.html").is_file(), **summarize(r),
            }  # fmt: skip
    return None


def list_runs(out_dir: Path) -> list[dict[str, Any]]:
    """All runs in the output folder, newest first."""
    rows = [d for p in sorted(out_dir.glob("*_*")) if p.is_dir() and (d := describe_run(p))]
    return sorted(rows, key=lambda r: r["created"], reverse=True)
