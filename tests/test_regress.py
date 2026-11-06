"""The run-to-run regression guard, and a plan contract on a synthetic deck (no Excel, no confidential template)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest
from pptx import Presentation

from resco.build import run_build
from resco.regress import coverage, find_regressions, recent_full_builds, template_changes
from test_deck_scan import _add_slide  # noqa: E402

LOGO = ("Picture 4", (925, 16, 65, 56))
TABLE, PROFIT = ("T", (95, 100, 720, 288)), ("P", (693, 400, 122, 50))


def _write_run(out: Path, name: str, **extra) -> None:
    d = out / name
    d.mkdir(parents=True)
    result = {"kind": "build", "deck": "main", "template": "deck.pptx", "selection": {"lobs": [], "parts": [], "slides": []}, **extra}
    (d / "build_result.json").write_text(json.dumps(result), encoding="utf-8")


def test_recent_full_builds_skip_partial_other_deck_and_non_builds(tmp_path: Path) -> None:
    _write_run(tmp_path, "build_20260101_000001", coverage={"slides": ["A / Results"], "tables": []})
    _write_run(tmp_path, "build_20260101_000002", selection={"lobs": ["ENTOT"], "parts": [], "slides": []})
    _write_run(tmp_path, "build_20260101_000003", deck="uk")
    _write_run(tmp_path, "build_20260101_000004", kind="dry-run")
    assert [n for n, _ in recent_full_builds(tmp_path, "main")] == ["build_20260101_000001"]


def test_missing_slide_or_table_is_a_regression_with_the_build_it_was_last_seen_in(tmp_path: Path) -> None:
    _write_run(tmp_path, "build_20260101_000001", coverage={"slides": ["DBTOT / Results", "ENTOT / Results"], "tables": ["gvn_DBTOT", "gvn_ENTOT"]})
    now = {"slides": ["ENTOT / Results"], "tables": ["gvn_ENTOT"]}
    regs = find_regressions(tmp_path, "main", now, "deck.pptx", None)
    assert {(r["kind"], r["item"], r["last_seen"]) for r in regs} == {("slide", "DBTOT / Results", "build_20260101_000001"), ("table", "gvn_DBTOT", "build_20260101_000001")}
    assert find_regressions(tmp_path, "main", {"slides": ["DBTOT / Results", "ENTOT / Results", "NEW / Results"], "tables": ["gvn_DBTOT", "gvn_ENTOT"]}, "deck.pptx", None) == []  # new things are fine


def test_old_runs_without_a_coverage_record_are_still_used(tmp_path: Path) -> None:
    _write_run(tmp_path, "build_20260101_000001", slides=[{"lob": "DBTOT", "type": "Results", "slots": [{"kind": "table"}]}], tables=[{"id": "gvn_DBTOT", "status": "filled"}])
    regs = find_regressions(tmp_path, "main", {"slides": [], "tables": []}, "deck.pptx", None)
    assert {r["item"] for r in regs} == {"DBTOT / Results", "gvn_DBTOT"}


def test_an_older_template_than_the_previous_run_is_reported(tmp_path: Path) -> None:
    _write_run(tmp_path, "build_20260101_000001", coverage={"slides": [], "tables": []}, template_modified="2026-10-23T10:00:00")
    regs = find_regressions(tmp_path, "main", {"slides": [], "tables": []}, "deck.pptx", "2026-10-22T19:00:00")
    assert [r["kind"] for r in regs] == ["template"]
    assert find_regressions(tmp_path, "main", {"slides": [], "tables": []}, "deck.pptx", "2026-10-23T11:00:00") == []


def test_template_changes_list_added_and_removed_slides(tmp_path: Path) -> None:
    _write_run(tmp_path, "build_20260101_000001", deck_titles=["Contents", "Approvals", "Approvals", "Old slide"])
    ch = template_changes(tmp_path, "main", ["Contents", "Approvals", "New slide"])
    assert ch == {"since": "build_20260101_000001", "added": ["New slide"], "removed": ["Approvals", "Old slide"]}
    assert template_changes(tmp_path / "nothing", "main", ["x"]) is None


def _deck(path: Path, *, with_dbtot: bool) -> Path:
    prs = Presentation()
    _add_slide(prs, "Contents", [LOGO])
    _add_slide(prs, "Energy (ENTOT) - Results", [LOGO, TABLE, PROFIT])
    if with_dbtot:
        _add_slide(prs, "Dubai (DBTOT) - Results", [LOGO, TABLE, PROFIT])
    _add_slide(prs, "Dubai Energy (DBENR) - Results", [LOGO, TABLE])
    _add_slide(prs, "Dubai Life (DBLIF) - Results", [LOGO, TABLE])
    prs.save(path)
    return path


def _dry(deck: Path, out: Path) -> dict:
    code = run_build(workbook=deck, template=deck, out=out, dry_run=True)  # the workbook is not opened in a dry run
    assert code == 0
    run = max(p for p in out.glob("build_*") if (p / "build_result.json").exists())
    return json.loads((run / "build_result.json").read_text(encoding="utf-8"))


def test_plan_contract_and_losing_a_slide_on_the_next_run(tmp_path: Path) -> None:
    out = tmp_path / "out"
    good = _dry(_deck(tmp_path / "good.pptx", with_dbtot=True), out)
    assert {"DBTOT / Results", "ENTOT / Results", "DBENR / Results"} <= set(good["coverage"]["slides"])
    assert [s["lob"] for s in good["slides"]].count("DBTOT") == 1
    # the Dubai Property / Treaty Results slides are added from DBENR (table only); their netting slides have no model in this deck
    added = {e["id"]: e["status"] for e in good["added_slides"]}
    assert added["dbpro_results"] == "added" and added["dbtrt_results"] == "added" and added["dbpro_netting"].startswith("failed")
    assert "DBPRO / Results" in good["coverage"]["slides"] and good["regressions"] == []

    # make that run count as a previous full build, then run on a template that lacks the DBTOT slide: the guard must say so
    run = max(p for p in out.glob("build_*"))
    prev = json.loads((run / "build_result.json").read_text(encoding="utf-8"))
    prev["kind"] = "build"
    (run / "build_result.json").write_text(json.dumps(prev), encoding="utf-8")
    again = _dry(_deck(tmp_path / "stale.pptx", with_dbtot=False), out)
    assert [(r["kind"], r["item"]) for r in again["regressions"]] == [("slide", "DBTOT / Results")]
    assert any(i["lob"] == "(regression)" for i in again["issues"])
    assert again["template_changes"]["removed"] == ["Dubai (DBTOT) - Results"]


def test_shipped_config_keeps_the_lobs_and_tables_we_added() -> None:
    """A guard against losing a fix when config is edited: these LoBs and table targets must stay."""
    from resco.config import load_lobs, load_slides_to_add, load_table_config

    codes = {lob.code for lob in load_lobs()}
    assert {"ALL", "ENTOT", "LITOT", "PRTOT", "DCTOT", "DBTOT", "DBENR", "DBLIF", "DBPRO", "DBTRT"} <= codes
    ids = {t.id for t in load_table_config().targets}
    assert {"gvn_ENTOT", "gvn_LITOT", "gvn_PRTOT", "gvn_DCTOT", "gvn_DBTOT", "gvn_DBPRO", "gvn_DBTRT", "gvn_ALL", "ibnr_alloc_totals", "ibnr_alloc_by_class", "ibnr_qoq_by_class"} <= ids
    assert {s.id for s in load_slides_to_add()} == {"dbpro_results", "dbpro_netting", "dbtrt_results", "dbtrt_netting"}
    uk = {lob.code for lob in load_lobs(Path(__file__).resolve().parents[1] / "config" / "uk")}
    assert {"ALL", "ENTOT", "PRTOT", "DCTOT", "ENUPS", "ENDWM", "ENCON", "DCMAR", "DCFIL"} <= uk
