"""Write images into a COPY of the template deck. The picture shape (id, name, position, size, z-order) is kept;
only the image it points to is swapped. python-pptx never touches the template file: we save to a new path."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from resco.deck_scan import EMU_PER_PT, _clean, _geometry, _is_picture, picture_blob
from resco.tables import BlockTexts, FillResult, LadderBlock, fill_blocks, fill_ladder, norm_title

_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


class DeckWriteError(RuntimeError):
    pass


@dataclass(frozen=True)
class Job:
    slide: int  # 1-based position in the template at scan time
    title: str  # expected title text, re-checked before writing
    shape_id: int
    kind: str
    lob: str
    image: Path
    new: bool = False  # insert a new picture (empty slide) instead of replacing one
    geometry: tuple[float, float, float, float] | None = None  # points: where a new picture goes


@dataclass
class TableJob:
    """Fill one native table of a narrative slide from Excel figures (the table keeps the deck's styling)."""

    slide: int
    title: str
    shape_id: int
    target_id: str
    kind: str  # ladder | overview
    block: LadderBlock | None = None
    ladder_groups: list[str] = field(default_factory=list)
    blocks: list[BlockTexts | None] = field(default_factory=list)
    result: FillResult | None = None
    error: str | None = None


def apply_table_jobs(prs, table_jobs: list[TableJob]) -> None:
    """Fill the tables of an already loaded presentation; errors are recorded on the job, never raised."""
    for tj in table_jobs:
        try:
            if not 1 <= tj.slide <= len(prs.slides):
                raise DeckWriteError(f"slide {tj.slide} does not exist in the template")
            slide = prs.slides[tj.slide - 1]
            if tj.title and norm_title(tj.title) not in norm_title(_slide_text(slide)):
                raise DeckWriteError(f"slide {tj.slide}: expected title '{tj.title}' not found; the deck changed since the scan")
            shape = next((s for s in slide.shapes if s.shape_id == tj.shape_id and getattr(s, "has_table", False) and s.has_table), None)
            if shape is None:
                raise DeckWriteError(f"slide {tj.slide}: table with id {tj.shape_id} not found")
            geometry_before = _geometry(shape)
            if tj.kind == "ladder":
                if tj.block is None:
                    raise DeckWriteError("no ladder rows for this table")
                tj.result = fill_ladder(shape.table, tj.block, tj.ladder_groups)
            else:
                tj.result = fill_blocks(shape.table, tj.blocks)
            if _geometry(shape) != geometry_before:
                raise DeckWriteError(f"slide {tj.slide}: the table moved or was resized")
        except DeckWriteError as exc:
            tj.error = str(exc)


@dataclass(frozen=True)
class JobResult:
    job: Job
    old_sha: str | None  # None for an inserted picture
    new_sha: str
    old_size: tuple[int, int] | None  # pixels
    new_size: tuple[int, int]
    shape_id: int = 0  # id of the picture in the built deck (differs from job.shape_id for an inserted one)


def _find_picture(slide, shape_id: int):
    for shape in slide.shapes:
        if shape.shape_id == shape_id and _is_picture(shape):
            return shape
    return None


def _slide_text(slide) -> str:
    parts = [_clean(s.text_frame.text) for s in slide.shapes if getattr(s, "has_text_frame", False) and s.has_text_frame]
    return " | ".join(p for p in parts if p)


def write_deck(template: Path, out_path: Path, jobs: list[Job], table_jobs: list[TableJob] | None = None) -> list[JobResult]:
    if out_path.resolve() == template.resolve():
        raise DeckWriteError("refusing to overwrite the template")
    prs = Presentation(str(template))
    results: list[JobResult] = []
    for job in jobs:
        if not 1 <= job.slide <= len(prs.slides):
            raise DeckWriteError(f"slide {job.slide} does not exist in the template")
        slide = prs.slides[job.slide - 1]
        if job.title and job.title not in _slide_text(slide):
            raise DeckWriteError(f"slide {job.slide}: expected title '{job.title}' not found; deck changed since the scan")
        if job.new:
            if job.geometry is None:
                raise DeckWriteError(f"slide {job.slide}: no position given for the new {job.kind} picture")
            left, top, width, height = (round(v * EMU_PER_PT) for v in job.geometry)
            added = slide.shapes.add_picture(str(job.image), left, top, width, height)
            added.name = f"ResCo {job.kind} {job.lob}"  # not "Picture N": that is how the logo is named
            added._element.nvPicPr.cNvPr.set("descr", f"{job.lob} {job.kind} (inserted by the ResCo deck builder)")
            results.append(JobResult(job, None, hashlib.sha256(added.image.blob).hexdigest(), None, _png_size(added.image.blob), added.shape_id))
            continue
        pic = _find_picture(slide, job.shape_id)
        if pic is None:
            raise DeckWriteError(f"slide {job.slide}: picture with id {job.shape_id} not found")

        blip = pic._element.find(f".//{_NS_A}blip")
        old_rid = blip.get(f"{_NS_R}embed") if blip is not None else None
        old_blob = picture_blob(pic)
        if old_rid is None or old_blob is None:
            raise DeckWriteError(f"slide {job.slide}: picture {job.shape_id} has no embedded image to replace")
        geometry_before = _geometry(pic)

        _, new_rid = slide.part.get_or_add_image_part(str(job.image))
        blip.set(f"{_NS_R}embed", new_rid)
        # python-pptx's own reference counter only looks at r:id (not r:embed), so count the references here and only
        # drop the old relationship when no other shape on this slide still uses the same image.
        still_used = [v for v in slide._element.xpath("//@r:embed | //@r:link | //@r:id") if v == old_rid]
        if old_rid != new_rid and not still_used:
            slide.part.drop_rel(old_rid)

        if _geometry(pic) != geometry_before:
            raise DeckWriteError(f"slide {job.slide}: picture {job.shape_id} moved or resized")
        results.append(
            JobResult(
                job,
                hashlib.sha256(old_blob).hexdigest(),
                hashlib.sha256(pic.image.blob).hexdigest(),
                _png_size(old_blob),
                _png_size(pic.image.blob),
                job.shape_id,
            )
        )
    apply_table_jobs(prs, table_jobs or [])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return results


def _png_size(blob: bytes) -> tuple[int, int]:
    import io

    from PIL import Image

    with Image.open(io.BytesIO(blob)) as im:
        return im.size


__all__ = ["DeckWriteError", "Job", "JobResult", "write_deck", "EMU_PER_PT"]
