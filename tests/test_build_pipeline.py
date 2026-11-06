"""deck_writer + verify on a synthetic deck (no Excel, no confidential template)."""
import io
import struct
import zlib
from pathlib import Path

import pytest
from pptx import Presentation

from resco.deck_scan import scan_deck
from resco.deck_writer import DeckWriteError, Job, write_deck
from resco.config import load_deck_rules
from resco.verify import compare, snapshot
from test_deck_scan import _add_slide  # noqa: E402

RULES = load_deck_rules()


def _png(path: Path, rgb: tuple[int, int, int], w: int = 4, h: int = 2) -> Path:
    def chunk(tag: bytes, data: bytes) -> bytes:
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    return path


@pytest.fixture()
def deck(tmp_path: Path) -> Path:
    prs = Presentation()
    logo = ("Picture 4", (925, 16, 65, 56))
    _add_slide(prs, "Executive Summary (1/3)", [logo])
    _add_slide(prs, "Energy Upstream (ENUPS) - Results", [logo, ("Picture 4", (95, 100, 720, 288)), ("Picture 7", (693, 400, 122, 50))])
    _add_slide(prs, "Energy Downstream (ENDWM) - Results", [logo, ("Picture 4", (95, 100, 720, 288))])
    path = tmp_path / "template.pptx"
    prs.save(path)
    return path


def _job(deck: Path, tmp_path: Path, slide: int, kind: str, color) -> Job:
    sc = scan_deck(deck, RULES)[slide - 1]
    pic = next(p for p in sc.pictures if p.kind == kind)
    return Job(slide, sc.title or "", pic.shape_id, kind, "ENUPS", _png(tmp_path / f"{kind}.png", color, 8, 3))


def test_replace_keeps_geometry_and_only_touches_target(deck: Path, tmp_path: Path):
    out = tmp_path / "out.pptx"
    jobs = [_job(deck, tmp_path, 2, "table", (200, 0, 0))]
    res = write_deck(deck, out, jobs)
    assert res[0].old_sha != res[0].new_sha and res[0].new_size == (8, 3)

    v = compare(snapshot(deck), snapshot(out), {2: {jobs[0].shape_id}})
    assert v.ok, v.messages
    assert v.stats["untouched_slides_identical"] == 2 and v.stats["pictures_replaced"] == 1


def test_verify_catches_unscheduled_change(deck: Path, tmp_path: Path):
    out = tmp_path / "out.pptx"
    j_table = _job(deck, tmp_path, 2, "table", (200, 0, 0))
    j_profit = _job(deck, tmp_path, 2, "profit", (0, 200, 0))
    write_deck(deck, out, [j_table, j_profit])
    # claim only the table was scheduled: the profit change must be reported
    v = compare(snapshot(deck), snapshot(out), {2: {j_table.shape_id}})
    assert not v.ok and any("not scheduled" in m for _, m in v.messages)


def test_refuses_to_overwrite_template_and_stale_title(deck: Path, tmp_path: Path):
    job = _job(deck, tmp_path, 2, "table", (1, 2, 3))
    with pytest.raises(DeckWriteError, match="overwrite"):
        write_deck(deck, deck, [job])
    stale = Job(job.slide, "Some Other Title", job.shape_id, job.kind, job.lob, job.image)
    with pytest.raises(DeckWriteError, match="title"):
        write_deck(deck, tmp_path / "o.pptx", [stale])


def test_compare_page_and_dashboard_render(deck: Path, tmp_path: Path):
    from resco.compare import write_compare
    from resco.dashboard import render_dashboard

    out = tmp_path / "out.pptx"
    job = _job(deck, tmp_path, 2, "table", (200, 0, 0))
    res = write_deck(deck, out, [job])
    rows = [{
        "slide": 2, "title": job.title, "lob": "ENUPS", "type": "Results", "kind": "table", "shape_id": job.shape_id,
        "old_size": res[0].old_size, "new_size": res[0].new_size, "distortion_old": 1.0, "distortion_new": 2.0,
    }]
    page = write_compare(deck, out, rows, tmp_path, deck.name)
    assert page.exists() and (tmp_path / "compare" / "s2_old.png").exists() and (tmp_path / "compare" / f"s2_{job.shape_id}_new.png").exists()
    html = render_dashboard(tmp_path, {"template": "t", "workbook": "w", "dry_run": False, "lobs": ["ENUPS"], "slides": [], "issues": [], "scan_flags": [], "seconds": 1})
    assert "ResCo deck builder" in html.read_text(encoding="utf-8")
