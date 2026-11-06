import json
import threading
import urllib.error
import urllib.request
from pathlib import Path

import pytest

from resco.serve import make_server


@pytest.fixture()
def server(tmp_path: Path):
    srv, _ = make_server(tmp_path, 0, {"workbook": "w.xlsm", "template": "t.pptx"})  # port 0 = any free port
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    yield f"http://127.0.0.1:{srv.server_address[1]}", tmp_path
    srv.shutdown()
    srv.server_close()


def _post(url: str, body: dict):
    req = urllib.request.Request(url + "/api/run", data=json.dumps(body).encode(), headers={"Content-Type": "application/json"})
    try:
        return urllib.request.urlopen(req, timeout=10).status, None
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read())["error"]


def test_status_and_defaults(server):
    url, _ = server
    st = json.loads(urllib.request.urlopen(url + "/api/status", timeout=10).read())
    assert st["state"] == "idle" and st["defaults"]["workbook"] == "w.xlsm"


def test_rejects_bad_parameters(server):
    url, tmp = server
    good_tpl = tmp / "t.pptx"
    good_tpl.write_bytes(b"x")
    assert _post(url, {"mode": "rm -rf", "deck": "main"})[0] == 400
    assert "LoB list" in _post(url, {"mode": "scan", "deck": "main", "lobs": "ENUPS; calc.exe", "template": str(good_tpl)})[1]
    assert "template" in _post(url, {"mode": "scan", "deck": "main", "lobs": "", "template": str(tmp / "missing.pptx")})[1]
    assert "workbook" in _post(url, {"mode": "build", "deck": "main", "lobs": "", "template": str(good_tpl), "workbook": str(tmp / "nope.xlsm")})[1]


def test_root_is_the_home_page_not_a_result(server):
    url, tmp = server
    # even with a finished run in the folder the home page must not show its results
    run = tmp / "build_20261020_120000"
    run.mkdir()
    (run / "build_result.json").write_text(json.dumps({"template": "ReservingPresentation_202606_UK.pptx", "workbook": "w", "dry_run": False, "slides": [], "issues": []}), encoding="utf-8")
    html = urllib.request.urlopen(url + "/", timeout=10).read().decode("utf-8")
    assert "Previous runs" in html and "New run" in html and "UK branch deck" in html
    assert "ReservingPresentation_202606_UK" not in html  # nothing of the run is embedded in the home page


def test_runs_listing_infers_deck_and_kind(server):
    url, tmp = server
    (tmp / "build_20261020_120000").mkdir()
    (tmp / "build_20261020_120000" / "build_result.json").write_text(json.dumps({"template": "ReservingPresentation_202606_UK.pptx", "workbook": "w", "dry_run": False, "slides": [], "issues": []}), encoding="utf-8")
    (tmp / "build_20261021_090000").mkdir()
    (tmp / "build_20261021_090000" / "build_result.json").write_text(json.dumps({"kind": "dry-run", "deck": "main", "template": "ReservingPresentation_202609.pptx", "workbook": "w", "dry_run": True, "slides": [], "issues": []}), encoding="utf-8")
    runs = json.loads(urllib.request.urlopen(url + "/api/runs", timeout=10).read())["runs"]
    assert [r["name"] for r in runs] == ["build_20261021_090000", "build_20261020_120000"]  # newest first
    assert runs[0]["deck"] == "main" and runs[0]["kind"] == "dry-run" and runs[0]["status"] == "none"
    assert runs[1]["deck"] == "uk" and runs[1]["kind"] == "build" and runs[1]["deck_label"] == "UK branch deck"


def test_deck_must_match_the_template_file(server):
    url, tmp = server
    lux, uk = tmp / "ReservingPresentation_202609.pptx", tmp / "ReservingPresentation_202606_UK.pptx"
    lux.write_bytes(b"x")
    uk.write_bytes(b"x")
    assert "says UK" in _post(url, {"mode": "scan", "deck": "main", "lobs": "", "template": str(uk)})[1]
    assert "does not say UK" in _post(url, {"mode": "scan", "deck": "uk", "lobs": "", "template": str(lux)})[1]
