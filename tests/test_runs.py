from pathlib import Path

from resco.dashboard import render_dashboard, render_home
from resco.runs import deck_of, describe_run, name_says_uk, summarize


def test_deck_detection():
    assert deck_of(Path("config")) == "main" and deck_of(Path("config/uk")) == "uk"
    assert name_says_uk("ReservingPresentation_202606_UK.pptx") and not name_says_uk("ReservingPresentation_202609.pptx")
    assert not name_says_uk("Bukowski.pptx")  # 'UK' inside a word is not the UK deck


def _result(**kw):
    base = {"template": "t.pptx", "workbook": "w", "dry_run": False, "slides": [{"slots": [{"status": "replaced"}, {"status": "replaced"}]}], "issues": [], "scan_flags": []}
    base.update(kw)
    return base


def test_summary_levels():
    assert summarize(_result())["status"] == "ok"
    assert summarize(_result(issues=[{"level": "WARN", "message": "x"}]))["status"] == "warn"
    s = summarize(_result(issues=[{"level": "ERROR", "message": "totals check failed (2 mismatches)"}]))
    assert s["status"] == "err" and s["label"] == "Deck built: workbook data checks failed"
    assert summarize(_result(issues=[{"level": "ERROR", "message": "boom"}]))["label"] == "Needs attention"
    assert summarize(_result(kind="dry-run"))["status"] == "none"
    assert summarize(_result(kind="scan", scan_flags=[{"level": "WARN", "message": "x"}]))["status"] == "warn"
    assert summarize(_result(kind="scan"))["label"] == "Deck recognised"


def test_describe_run_ignores_folders_without_result(tmp_path: Path):
    (tmp_path / "dump_20261020_110000").mkdir()
    assert describe_run(tmp_path / "dump_20261020_110000") is None


def test_pages_show_the_deck(tmp_path: Path):
    page = render_dashboard(tmp_path, _result(template="Deck_UK.pptx")).read_text(encoding="utf-8")
    assert '"deck": "uk"' in page and "UK branch deck" in page  # deck inferred and embedded for old-style results
    scan = render_dashboard(tmp_path, _result(kind="scan", deck="main", other_slides=3)).read_text(encoding="utf-8")
    assert '"kind": "scan"' in scan
    assert "Previous runs" in render_home()


def test_workbook_check_summary_and_labels():
    from resco.runs import KIND_LABEL

    assert KIND_LABEL["check"] == "Workbook check"
    assert summarize(_result(kind="check"))["label"] == "Workbook ready"
    bad = summarize(_result(kind="check", issues=[{"level": "ERROR", "message": "no Expert Judgement block for 'X'"}]))
    assert bad["status"] == "err" and bad["label"] == "Problems found in the workbook"


def test_ej_capacity_from_the_render_template_ranges():
    from resco.config import load_ranges
    from resco.exporter import ej_capacity

    assert ej_capacity(load_ranges()) == 10


def test_dashboard_script_declares_tbls_before_using_it(tmp_path: Path):
    """JavaScript const variables are unusable before their declaration (the whole page script then dies and the run page is
    blank): 'tbls' must be declared before its first use."""
    page = render_dashboard(tmp_path, _result()).read_text(encoding="utf-8")
    assert 0 <= page.index("const tbls = ") < page.index("tbls.length")
