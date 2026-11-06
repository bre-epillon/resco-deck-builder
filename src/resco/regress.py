"""Run-to-run regression guard: what the template covers now versus what recent full builds of the same deck covered.

A slide or table that was filled in a recent full build but is not in this run's plan is reported ("regression"): typically a
stale or wrong template, an input file that was swapped, or a slide that was deleted. The guard only reports; it never changes a run.
"""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

from resco.deck_scan import DeckMap
from resco.tables import norm_title

KEEP = 5  # how many recent full builds are compared with


def coverage(dm: DeckMap) -> dict[str, list[str]]:
    """What the template offers to fill: 'LoB / slide type' for every LoB slide with picture slots, and every matched table target."""
    return {
        "slides": sorted({f"{m.lob} / {m.type}" for m in dm.mapped if m.slots}),
        "tables": sorted({m.target.id for m in dm.table_matches}),
    }


def _of_result(result: dict[str, Any]) -> dict[str, list[str]] | None:
    """Coverage recorded by a run; runs from before this feature are reconstructed from their slide / table lists."""
    if "coverage" in result:
        return result["coverage"]
    slides = sorted({f"{s['lob']} / {s['type']}" for s in result.get("slides", []) if s.get("slots")})
    tables = sorted({t["id"] for t in result.get("tables", []) if str(t.get("status")) == "filled"})
    return {"slides": slides, "tables": tables} if slides else None


def recent_full_builds(out_dir: Path, deck: str, limit: int = KEEP) -> list[tuple[str, dict[str, Any]]]:
    """Newest first: builds of this deck that covered the whole template (no LoB / content / slide selection)."""
    found: list[tuple[str, dict[str, Any]]] = []
    for d in sorted((p for p in out_dir.glob("build_*") if p.is_dir()), key=lambda p: p.name, reverse=True):
        f = d / "build_result.json"
        if not f.is_file():
            continue
        try:
            r = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        sel = r.get("selection") or {}
        if r.get("kind") != "build" or r.get("deck") != deck or sel.get("lobs") or sel.get("parts") or sel.get("slides"):
            continue
        found.append((d.name, r))
        if len(found) >= limit:
            break
    return found


def find_regressions(out_dir: Path, deck: str, now: dict[str, list[str]], template_name: str, template_modified: str | None) -> list[dict[str, str]]:
    """Items present in a recent full build but not in `now`, plus a template that is older than the one used before."""
    out: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    previous = recent_full_builds(out_dir, deck)
    for name, r in previous:
        before = _of_result(r)
        if before is None:
            continue
        for kind, label in (("slides", "slide"), ("tables", "table")):
            for item in before[kind]:
                if item not in now[kind] and (kind, item) not in seen:
                    seen.add((kind, item))
                    out.append({"kind": label, "item": item, "last_seen": name})
    if previous and template_modified:
        name, r = previous[0]
        prev_mod = r.get("template_modified")
        if prev_mod and r.get("template") == template_name and template_modified < prev_mod:
            out.append({"kind": "template", "item": f"'{template_name}' is older than the one used in the previous full build (saved {template_modified} vs {prev_mod})", "last_seen": name})
    return out


def template_changes(out_dir: Path, deck: str, now_titles: list[str]) -> dict[str, Any] | None:
    """Slides added to / removed from the template since the most recent full build that recorded its slide titles."""
    for name, r in recent_full_builds(out_dir, deck):
        before = r.get("deck_titles")
        if not before:
            continue
        b, n = Counter(norm_title(t) for t in before), Counter(norm_title(t) for t in now_titles)
        shown = {norm_title(t): t for t in [*before, *now_titles]}
        return {"since": name, "added": sorted(shown[k] for k in (n - b)), "removed": sorted(shown[k] for k in (b - n))}
    return None


def describe(reg: dict[str, str]) -> str:
    if reg["kind"] == "template":
        return f"regression: the template {reg['item']} [{reg['last_seen']}]"
    return f"regression: {reg['kind']} '{reg['item']}' was filled in {reg['last_seen']} but is not in this run's plan (stale or wrong template, or a slide that was removed?)"
