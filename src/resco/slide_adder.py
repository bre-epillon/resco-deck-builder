"""Add the slides a run needs but the template does not have yet (config/slides_add.yaml).

A new slide is a copy of a model slide of the template, retitled and placed after another slide:
  * `frame`: title, logo, slide number and every other non-picture, non-table shape; the content pictures and tables are NOT copied,
    so a failed export leaves the slide empty instead of showing another LoB's figures. The build then inserts the pictures
    (the kinds the model slide had) at the standard positions, like on any empty LoB slide.
  * `full`: a complete copy (native table, legend pictures, notes text boxes); the build fills the table from Excel.
The template file is never modified: the new deck is written to a separate file and the rest of the build works on that.
"""
from __future__ import annotations

import copy
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from resco.config import DeckRules, NewSlide
from resco.deck_scan import _clean, _geometry, _is_picture, _title_of, classify_geometry
from resco.tables import norm_title

log = logging.getLogger(__name__)


@dataclass
class AddResult:
    entries: list[dict[str, Any]] = field(default_factory=list)  # one per spec: id, title, status, position, like, after
    added_positions: set[int] = field(default_factory=set)  # 1-based positions of the new slides in the new deck
    kinds_by_title: dict[str, tuple[str, ...]] = field(default_factory=dict)  # norm title -> picture kinds to insert (frame copies)
    saved: bool = False

    @property
    def added(self) -> int:
        return len(self.added_positions)


def missing_specs(titles: Iterable[str | None], specs: Iterable[NewSlide]) -> list[NewSlide]:
    """The specs whose slide is not in the deck yet (matched by title, dashes and case ignored)."""
    have = {norm_title(t) for t in titles if t}
    return [s for s in specs if norm_title(s.title) not in have]


def _titles(prs: Any, rules: DeckRules) -> list[str | None]:
    return [_title_of(slide, rules.total_code)[0] for slide in prs.slides]


def _find(titles: list[str | None], wanted: Iterable[str]) -> int | None:
    """0-based index of the first slide whose title matches one of the wanted titles (in the order given)."""
    for w in wanted:
        for i, t in enumerate(titles):
            if t and norm_title(t) == norm_title(w):
                return i
    return None


def _retitle(slide: Any, old_title: str, new_title: str) -> bool:
    for shape in slide.shapes:
        if not (getattr(shape, "has_text_frame", False) and shape.has_text_frame):
            continue
        if norm_title(_clean(shape.text_frame.text)) != norm_title(old_title):
            continue
        paragraphs = shape.text_frame.paragraphs
        first = paragraphs[0]._p
        runs = first.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r")
        if not runs:
            return False
        runs[0].find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = new_title
        for child in list(first):  # the rest of the title (more runs, line breaks) goes; formatting of the first run stays
            tag = child.tag.rsplit("}", 1)[-1]
            if child is not runs[0] and tag in ("r", "br", "fld"):
                first.remove(child)
        for extra in paragraphs[1:]:
            extra._p.getparent().remove(extra._p)
        return True
    return False


def _clone(prs: Any, src: Any, spec: NewSlide, model_title: str, rules: DeckRules) -> tuple[Any, list[str], list[str]]:
    """Returns (new slide, picture kinds of the model that were left out, names of shapes that could not be copied)."""
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):  # the layout's placeholders: the copied shapes replace them
        shp._element.getparent().remove(shp._element)
    rid_map = {rid: new.part.relate_to(rel.target_part, rel.reltype) for rid, rel in src.part.rels.items() if rel.reltype == RT.IMAGE and not rel.is_external}
    r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    left_out: list[str] = []
    skipped: list[str] = []
    for shp in src.shapes:
        if spec.copy == "frame":
            if _is_picture(shp):
                kind = classify_geometry(_geometry(shp), rules)
                if kind != "logo":
                    if kind:
                        left_out.append(kind)
                    continue
            elif (getattr(shp, "has_table", False) and shp.has_table) or (getattr(shp, "has_chart", False) and shp.has_chart) or shp.shape_type == MSO_SHAPE_TYPE.GROUP:
                continue
        el = copy.deepcopy(shp._element)
        refs = [v for v in el.xpath(".//@r:embed | .//@r:link | .//@r:id")]
        if any(v not in rid_map for v in refs):
            skipped.append(shp.name)  # linked to something other than an image: not copied, reported
            continue
        for node in el.iter():
            for attr in (f"{r_ns}embed", f"{r_ns}link", f"{r_ns}id"):
                if node.get(attr) in rid_map:
                    node.set(attr, rid_map[node.get(attr)])
        new.shapes._spTree.insert_element_before(el, "p:extLst")
    _retitle(new, model_title, spec.title)
    return new, left_out, skipped


def add_missing_slides(template: Path, out: Path, specs: list[NewSlide], rules: DeckRules) -> AddResult:
    """Write `out` = the template plus the slides of `specs` it does not have. Nothing is written when nothing is missing."""
    if out.resolve() == template.resolve():
        raise ValueError("refusing to overwrite the template")
    result = AddResult()
    prs = Presentation(str(template))
    lst = prs.slides._sldIdLst
    new_ids: list[int] = []
    for spec in specs:
        titles = _titles(prs, rules)
        anchors = (spec.after,) if isinstance(spec.after, str) else tuple(spec.after or ())
        entry: dict[str, Any] = {"id": spec.id, "title": spec.title, "lob": spec.lob, "part": spec.part, "status": "", "position": None, "like": None, "after": anchors[0] if anchors else None}
        result.entries.append(entry)
        if _find(titles, [spec.title]) is not None:
            entry["status"] = "present"
            continue
        i_model = _find(titles, spec.like)
        if i_model is None:
            entry["status"] = f"failed: no model slide ({' / '.join(spec.like)}) in the template"
            continue
        model_title = titles[i_model] or ""
        entry["like"] = model_title
        slide, left_out, skipped = _clone(prs, prs.slides[i_model], spec, model_title, rules)
        new_el = lst[-1]
        new_ids.append(int(new_el.get("id")))
        i_after = _find(_titles(prs, rules)[:-1], anchors) if anchors else None  # the new slide is last at this point
        if i_after is not None:
            lst.remove(new_el)
            lst.insert(i_after + 1, new_el)
            entry["status"] = "added"
        else:
            entry["status"] = "added" if not anchors else f"added at the end: no slide '{anchors[0]}' to place it after"
        if skipped:
            entry["notes"] = f"not copied (linked to something other than an image): {', '.join(skipped)}"
        if spec.copy == "frame" and left_out:
            result.kinds_by_title[norm_title(spec.title)] = tuple(left_out)
    if not new_ids:
        return result
    position_of = {int(el.get("id")): n for n, el in enumerate(lst, start=1)}
    result.added_positions = {position_of[i] for i in new_ids}
    for entry in result.entries:
        if entry["status"].startswith("added"):
            titles = _titles(prs, rules)
            entry["position"] = (_find(titles, [entry["title"]]) or 0) + 1
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    result.saved = True
    log.info("added %d slide(s) to the working copy of the template", result.added)
    return result
