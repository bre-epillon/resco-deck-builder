"""Compare the built deck with the template: only the intended pictures may have changed (CLAUDE.md validation 4-6)."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree
from pptx import Presentation

from resco.deck_scan import _clean, _geometry, _is_picture, picture_blob


@dataclass(frozen=True)
class ShapeSnap:
    shape_id: int
    name: str
    geometry: tuple[float, float, float, float] | None
    text: str
    image_sha: str | None
    is_table: bool = False


@dataclass(frozen=True)
class SlideSnap:
    index: int
    xml_sha: str
    shapes: tuple[ShapeSnap, ...]


@dataclass
class VerifyResult:
    messages: list[tuple[str, str]] = field(default_factory=list)  # (level, text)
    stats: dict[str, int] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not any(level == "ERROR" for level, _ in self.messages)


def snapshot(path: Path) -> list[SlideSnap]:
    prs = Presentation(str(path))
    snaps: list[SlideSnap] = []
    for i, slide in enumerate(prs.slides, start=1):
        shapes = []
        for sh in slide.shapes:
            text = _clean(sh.text_frame.text) if getattr(sh, "has_text_frame", False) and sh.has_text_frame else ""
            is_table = bool(getattr(sh, "has_table", False) and sh.has_table)
            if is_table:
                text = " \u00a6 ".join(_clean(c.text) for row in sh.table.rows for c in row.cells)
            blob = picture_blob(sh) if _is_picture(sh) else None
            sha = hashlib.sha256(blob).hexdigest() if blob is not None else ("no-embedded-image" if _is_picture(sh) else None)
            shapes.append(ShapeSnap(sh.shape_id, sh.name, _geometry(sh), text, sha, is_table))
        xml = hashlib.sha256(etree.tostring(slide._element, method="c14n")).hexdigest()
        snaps.append(SlideSnap(i, xml, tuple(shapes)))
    return snaps


def compare_prepared(original: list[SlideSnap], prepared: list[SlideSnap], added_positions: set[int]) -> list[tuple[str, str]]:
    """The deck with the added slides must hold every original slide unchanged (XML-identical) in its original order."""
    if len(prepared) != len(original) + len(added_positions):
        return [("ERROR", f"slide count after adding slides is {len(prepared)}, expected {len(original)} + {len(added_positions)}")]
    kept = [s.xml_sha for i, s in enumerate(prepared, start=1) if i not in added_positions]
    if kept != [s.xml_sha for s in original]:
        return [("ERROR", "adding slides changed or reordered an original slide")]
    return []


def compare(
    template: list[SlideSnap], built: list[SlideSnap], replaced: dict[int, set[int]], filled: dict[int, set[int]] | None = None, inserted: dict[int, int] | None = None,
    inserted_tables: dict[int, int] | None = None,
) -> VerifyResult:
    """`replaced` maps slide index -> picture shape ids the build was supposed to change; `filled` the same for tables whose text was filled;
    `inserted` slide index -> number of pictures that were to be added to an empty slide; `inserted_tables` the same for tables
    (the Profit (loss) movement QoQ boxes)."""
    filled = filled or {}
    inserted = inserted or {}
    inserted_tables = inserted_tables or {}
    res = VerifyResult()
    err = lambda m: res.messages.append(("ERROR", m))  # noqa: E731
    if len(template) != len(built):
        err(f"slide count changed: {len(template)} -> {len(built)}")
        return res

    untouched_identical = replaced_ok = inserted_ok = tables_inserted_ok = 0
    for before, after in zip(template, built):
        i = before.index
        expected = replaced.get(i, set())
        tables = filled.get(i, set())
        n_new, n_tab = inserted.get(i, 0), inserted_tables.get(i, 0)
        ids_before, ids_after = [s.shape_id for s in before.shapes], [s.shape_id for s in after.shapes]
        if ids_after[: len(ids_before)] != ids_before or len(ids_after) != len(ids_before) + n_new + n_tab:
            what = [f"{n_new} new picture(s)"] * bool(n_new) + [f"{n_tab} new table(s)"] * bool(n_tab)
            err(f"slide {i}: shapes added, removed or reordered" + (f" (expected {' and '.join(what)})" if what else ""))
            continue
        extras = after.shapes[len(before.shapes):]  # pictures inserted into an empty slide, tables inserted under a Gross vs Net table
        n_pics, n_tabs = sum(1 for e in extras if e.image_sha is not None), sum(1 for e in extras if e.is_table)
        if (n_pics, n_tabs) != (n_new, n_tab):
            err(f"slide {i}: inserted {n_pics} picture(s) and {n_tabs} table(s), expected {n_new} and {n_tab}")
        else:
            inserted_ok += n_pics
            tables_inserted_ok += n_tabs
        for sb, sa in zip(before.shapes, after.shapes):
            where = f"slide {i} shape {sb.shape_id} '{sb.name}'"
            if (sb.name, sb.geometry) != (sa.name, sa.geometry) or (sb.text != sa.text and sb.shape_id not in tables):
                err(f"{where}: name, position, size or text changed")
            if sb.shape_id in expected:
                if sb.image_sha == sa.image_sha:
                    res.messages.append(("INFO", f"{where}: new image is byte-identical to the old one"))
                replaced_ok += 1
            elif sb.image_sha != sa.image_sha:
                err(f"{where}: image changed but was not scheduled for replacement")
        if not expected and not tables and not n_new and not n_tab:
            if before.xml_sha != after.xml_sha:
                err(f"slide {i}: untouched slide differs in its XML")
            else:
                untouched_identical += 1
    res.stats = {
        "slides": len(template),
        "touched_slides": len(set(replaced) | set(filled) | set(inserted) | set(inserted_tables)),
        "tables_filled": sum(len(v) for v in filled.values()),
        "untouched_slides_identical": untouched_identical,
        "pictures_replaced": replaced_ok,
        "pictures_inserted": inserted_ok,
    }
    if inserted_tables:
        res.stats["tables_inserted"] = tables_inserted_ok
    return res
