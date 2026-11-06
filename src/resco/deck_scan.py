"""Step 3: scan a template deck and map (LoB, slide type) -> picture slots.

Nothing here depends on slide numbers or shape names. A LoB slide is found by its title
("<Name> (<CODE>) - <Type>", hyphen or en dash, or "Total <Type>"); each picture on it is classified by matching
its position and size against config/slot_kinds.yaml. Anything that does not fit is reported, never guessed.
Only pptx *reading* happens here (the template is never saved).
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from resco.config import DeckRules, Lob, TableTarget
from resco.tables import norm_title, table_signature

EMU_PER_PT = 12700
APPROX_MAX_POS = 0.12  # centre distance, as a fraction of the slide diagonal
APPROX_MAX_SIZE = 0.35  # relative width / height difference
SLIDE_TYPES = ("Results", "Triangles", "Patterns", "Expert Judgement")
_TYPE_PATTERN = "|".join(SLIDE_TYPES)
# An optional note in brackets after the type is allowed, e.g. "... – Expert Judgement (old methodology)" (Q3 2026 decks).
TITLE_RE = re.compile(
    rf"^(?P<name>.+?)\s*\((?P<code>[A-Z]{{2,6}})\)\s*[-–—]\s*(?P<type>{_TYPE_PATTERN})\s*(?:\([^()]*\)\s*)?$", re.IGNORECASE
)
TOTAL_RE = re.compile(rf"^Total\s*(?:[-–—]\s*)?(?P<type>{_TYPE_PATTERN})\s*$", re.IGNORECASE)
LIFE_AWARE_KINDS = {"table": False, "table_life": True, "ej": False, "ej_life": True}


@dataclass(frozen=True)
class ParsedTitle:
    name: str
    code: str  # LoB code; the configured total code for "Total <Type>" slides
    type: str  # one of SLIDE_TYPES, normalised capitalisation


@dataclass
class PictureInfo:
    shape_id: int
    name: str
    geometry: tuple[float, float, float, float] | None  # points
    kind: str | None  # slot kind, "logo", or None when unrecognised
    approx: bool = False  # matched to the slot approximately (the picture was moved or resized in the deck)


@dataclass
class TableInfo:
    shape_id: int
    name: str
    rows: int
    cols: int
    signature: str  # normalised text of the first two rows


@dataclass
class SlideScan:
    index: int  # 1-based position in the deck at scan time (informational only, never used as a key)
    title: str | None
    parsed: ParsedTitle | None
    pictures: list[PictureInfo] = field(default_factory=list)
    other_shapes: list[str] = field(default_factory=list)
    tables: list[TableInfo] = field(default_factory=list)


@dataclass(frozen=True)
class Flag:
    level: str  # ERROR | WARN | INFO
    slide: int | None
    lob: str | None
    message: str


@dataclass
class SlotEntry:
    kind: str
    shape_id: int
    shape_name: str
    geometry: tuple[float, float, float, float]
    new: bool = False  # the slide has no picture here yet: the build inserts one at the standard position


@dataclass
class MappedSlide:
    slide: int
    title: str
    lob: str
    selector: str | None
    resclass: str
    title_type: str
    type: str  # effective type: what the pictures on the slide say it is
    slots: list[SlotEntry]


# --------------------------------------------------------------------------- pure helpers
def normalise_type(text: str) -> str:
    for t in SLIDE_TYPES:
        if t.lower() == text.lower():
            return t
    raise ValueError(text)


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\x0b", " ")).strip()


def parse_title(text: str, total_code: str = "ALL") -> ParsedTitle | None:
    text = _clean(text)
    m = TITLE_RE.match(text)
    if m:
        return ParsedTitle(m.group("name").strip(), m.group("code").upper(), normalise_type(m.group("type")))
    m = TOTAL_RE.match(text)
    if m:
        return ParsedTitle("Total", total_code, normalise_type(m.group("type")))
    return None


def _near(a: tuple[float, ...], b: tuple[float, ...], tol: float) -> bool:
    return all(abs(x - y) <= tol for x, y in zip(a, b))


def classify_geometry(geometry: tuple[float, float, float, float] | None, rules: DeckRules) -> str | None:
    """'logo', a slot kind, or None. Logos are matched first; a picture must match position AND size."""
    if geometry is None:
        return None
    if any(_near(geometry, logo, rules.tolerance_pt) for logo in rules.logos):
        return "logo"
    hits = [k for k, sk in rules.kinds.items() if _near(geometry, sk.geometry, rules.tolerance_pt)]
    return hits[0] if len(hits) == 1 else None


# --------------------------------------------------------------------------- reading the deck
def _geometry(shape: Any) -> tuple[float, float, float, float] | None:
    if None in (shape.left, shape.top, shape.width, shape.height):
        return None
    return tuple(round(v / EMU_PER_PT, 1) for v in (shape.left, shape.top, shape.width, shape.height))


def _is_picture(shape: Any) -> bool:
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return True
    return bool(getattr(shape, "is_placeholder", False)) and hasattr(shape, "image")


def picture_blob(shape: Any) -> bytes | None:
    """Image bytes of a picture shape, or None when it has no embedded image (linked picture, empty placeholder)."""
    try:
        return shape.image.blob
    except (ValueError, KeyError, AttributeError):
        return None


def _count_pictures(group: Any) -> int:
    n = 0
    for s in group.shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            n += _count_pictures(s)
        elif _is_picture(s):
            n += 1
    return n


def _title_of(slide: Any, total_code: str) -> tuple[str | None, ParsedTitle | None]:
    candidates: list[str] = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            text = _clean(shape.text_frame.text)
            if text:
                candidates.append(text)
                parsed = parse_title(text, total_code)
                if parsed:
                    return text, parsed
    placeholder = getattr(slide.shapes, "title", None)
    if placeholder is not None and placeholder.has_text_frame and _clean(placeholder.text_frame.text):
        return _clean(placeholder.text_frame.text), None
    return (candidates[0] if candidates else None), None


def _relax(scan: SlideScan, rules: DeckRules, slide_w: float, slide_h: float) -> None:
    """Second chance for pictures that no slot matches exactly (someone nudged or resized them): within the slide's own
    type, assign each to the nearest free slot if it is close in position and size. The picture keeps its own position
    and size when its image is replaced; the match is flagged as approximate."""
    if scan.parsed is None:
        return
    loose = [p for p in scan.pictures if p.kind is None and p.geometry is not None]
    if not loose:
        return
    taken = {p.kind for p in scan.pictures if p.kind and p.kind != "logo"}
    types = {rules.kind_type[k] for k in taken if k in rules.kind_type}
    slide_type = next(iter(types)) if len(types) == 1 else scan.parsed.type
    candidates = [k for k, t in rules.kind_type.items() if t == slide_type and k not in taken]
    diag = math.hypot(slide_w, slide_h)
    scored: list[tuple[float, int, str, PictureInfo]] = []
    for n, pic in enumerate(loose):
        x, y, w, h = pic.geometry  # type: ignore[misc]
        for kind in candidates:
            kx, ky, kw, kh = rules.kinds[kind].geometry
            d_pos = math.hypot(x + w / 2 - (kx + kw / 2), y + h / 2 - (ky + kh / 2)) / diag
            d_size = max(abs(w / kw - 1), abs(h / kh - 1))
            if d_pos <= APPROX_MAX_POS and d_size <= APPROX_MAX_SIZE:
                scored.append((d_pos + d_size, n, kind, pic))
    for _, _, kind, pic in sorted(scored, key=lambda t: (t[0], t[1])):
        if pic.kind is None and kind not in taken:
            pic.kind, pic.approx = kind, True
            taken.add(kind)


class TemplateLocked(OSError):
    """The template cannot be read right now (typically open in PowerPoint while it is being saved)."""


def _windows_copy(src: Path, dest: Path) -> bool:
    """CopyFileW: unlike Python's open(), it can copy a file that PowerPoint has open (the last SAVED version is copied)."""
    import ctypes
    import sys

    if sys.platform != "win32":
        return False
    return bool(ctypes.windll.kernel32.CopyFileW(str(src), str(dest), False))  # type: ignore[attr-defined]


def copy_for_reading(path: Path, dest_dir: Path, tries: int = 6, wait: float = 5.0) -> Path:
    """Copy the template to `dest_dir` and return the copy: the run then works on a stable file even if the original is saved
    again during the run. A deck that is open in PowerPoint is copied through the Windows copy call (its last saved version);
    a file that is locked at this moment is retried for about `tries * wait` seconds."""
    import shutil
    import time

    dest = dest_dir / path.name
    last: Exception | None = None
    for attempt in range(tries):
        try:
            shutil.copyfile(path, dest)
            return dest
        except OSError as exc:  # PermissionError: the deck is open in PowerPoint, or a OneDrive file is being fetched
            last = exc
            if _windows_copy(path, dest):
                return dest
            if attempt < tries - 1:
                time.sleep(wait)
    raise TemplateLocked(f"cannot read '{path.name}' ({last}): it is locked by another program. Save and close it, or work on a copy, then run again.")


def scan_deck(path: Path, rules: DeckRules) -> list[SlideScan]:
    prs = Presentation(str(path))  # read-only: never saved
    scans: list[SlideScan] = []
    for i, slide in enumerate(prs.slides, start=1):
        title, parsed = _title_of(slide, rules.total_code)
        scan = SlideScan(index=i, title=title, parsed=parsed)
        for shape in slide.shapes:
            if _is_picture(shape):
                geo = _geometry(shape)
                scan.pictures.append(PictureInfo(shape.shape_id, shape.name, geo, classify_geometry(geo, rules)))
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                n = _count_pictures(shape)
                if n:
                    scan.other_shapes.append(f"group '{shape.name}' containing {n} picture(s)")
            elif getattr(shape, "has_table", False) and shape.has_table:
                scan.other_shapes.append(f"table '{shape.name}'")
                scan.tables.append(TableInfo(shape.shape_id, shape.name, len(shape.table.rows), len(shape.table.columns), table_signature(shape.table)))
            elif getattr(shape, "has_chart", False) and shape.has_chart:
                scan.other_shapes.append(f"chart '{shape.name}'")
        _relax(scan, rules, prs.slide_width / EMU_PER_PT, prs.slide_height / EMU_PER_PT)
        scans.append(scan)
    return scans


# --------------------------------------------------------------------------- table targets
def match_table_targets(scans: list[SlideScan], targets: tuple[TableTarget, ...]) -> tuple[list[TableMatch], list[Flag]]:
    """Find, for each configured target, the slide (by title) and the native table on it (by words in its first rows)."""
    matches: list[TableMatch] = []
    flags: list[Flag] = []
    for t in targets:
        slides = [sc for sc in scans if sc.title and norm_title(sc.title) == norm_title(t.slide)]
        if not slides:
            flags.append(Flag("WARN", None, t.id, f"table target '{t.id}': no slide titled '{t.slide}' in the template: nothing is filled there"))
            continue
        for sc in slides:
            cands = [tb for tb in sc.tables if all(w in tb.signature for w in t.match)]
            if not cands:
                flags.append(Flag("WARN", sc.index, t.id, f"table target '{t.id}': slide {sc.index} has no table whose first rows contain {list(t.match)}"))
                continue
            if len(cands) > 1:
                flags.append(Flag("WARN", sc.index, t.id, f"table target '{t.id}': {len(cands)} tables on slide {sc.index} match, the first is used"))
            tb = cands[0]
            matches.append(TableMatch(t, sc.index, sc.title or "", tb.shape_id, tb.name, tb.rows, tb.cols))
    return matches, flags


# --------------------------------------------------------------------------- mapping
@dataclass
class TableMatch:
    target: TableTarget
    slide: int
    title: str
    shape_id: int
    table_name: str
    rows: int
    cols: int


@dataclass
class DeckMap:
    mapped: list[MappedSlide]
    narrative: list[SlideScan]  # slides that are not LoB slides (never touched)
    flags: list[Flag]
    lobs_without_slides: list[str]
    table_matches: list[TableMatch] = field(default_factory=list)


def build_map(
    scans: list[SlideScan], lobs: list[Lob], life_resclasses: frozenset[str], rules: DeckRules, targets: tuple[TableTarget, ...] = (), empty_kinds: dict[str, tuple[str, ...]] | None = None,
) -> DeckMap:
    """`empty_kinds` (normalised title -> picture kinds): what to insert into an empty slide that was added from a model slide."""
    by_code = {lob.code: lob for lob in lobs}
    for lob in lobs:
        for alias in lob.aliases:
            by_code.setdefault(alias, lob)
    mapped: list[MappedSlide] = []
    narrative: list[SlideScan] = []
    flags: list[Flag] = []

    for sc in scans:
        if sc.parsed is None:
            narrative.append(sc)
            continue
        p = sc.parsed
        add = lambda level, msg: flags.append(Flag(level, sc.index, p.code, msg))  # noqa: E731
        lob = by_code.get(p.code)
        if lob is None:
            add("ERROR", f"no LoB '{p.code}' in lobs.yaml: slide cannot be fed")
        elif lob.code != p.code:
            add("WARN", f"title code '{p.code}' is a known alias of '{lob.code}' (typo in the deck title?): treated as {lob.code}")

        slots: list[SlotEntry] = []
        for pic in sc.pictures:
            if pic.kind == "logo":
                continue
            if pic.kind is None:
                add("WARN", f"unrecognised picture '{pic.name}' (id {pic.shape_id}) at {pic.geometry}: left untouched")
            else:
                slots.append(SlotEntry(pic.kind, pic.shape_id, pic.name, pic.geometry))  # type: ignore[arg-type]
                if pic.approx:
                    add("WARN", f"picture '{pic.name}' (id {pic.shape_id}) was matched to slot '{pic.kind}' approximately: it was moved or resized in the deck (it keeps its own position and size)")

        if lob is not None and not any(pic.kind != "logo" for pic in sc.pictures):
            # a LoB slide with a title but no picture at all (a slide just added to the template): its pictures are inserted at the
            # standard positions of the slide type, so a new LoB only needs a titled slide
            wants_life_slide = lob.resclass in life_resclasses
            like_model = (empty_kinds or {}).get(norm_title(sc.title or ""))  # a slide added from a model: the model's kinds
            default_kinds = [next((k for k in group if LIFE_AWARE_KINDS.get(k) == wants_life_slide), group[0]) if len(group) > 1 else group[0] for group in rules.required.get(p.type, ())]
            for n, kind in enumerate(like_model or default_kinds, start=1):
                slots.append(SlotEntry(kind, -n, "(new picture)", rules.kinds[kind].geometry, new=True))
            if slots:
                add("INFO" if like_model else "WARN", f"the slide has no pictures: {', '.join(s.kind for s in slots)} will be inserted at the standard positions" + ("" if like_model else " (check the layout after the build)"))

        kinds = [s.kind for s in slots]
        for k in sorted({k for k in kinds if kinds.count(k) > 1}):
            add("WARN", f"{kinds.count(k)} pictures match slot '{k}'")

        content_types = {rules.kind_type[k] for k in kinds}
        if not content_types:
            effective = p.type
            add("WARN", f"title says {p.type} but no picture matches a known slot")
        elif len(content_types) > 1:
            effective = p.type
            add("ERROR", f"pictures belong to several slide types {sorted(content_types)}")
        else:
            effective = next(iter(content_types))
            if effective != p.type:
                add("WARN", f"title says '{p.type}' but the pictures are a {effective} layout: treated as {effective}")

        for group in rules.required.get(effective, ()):
            if not any(k in kinds for k in group):
                add("INFO", f"no '{'/'.join(group)}' slot on this {effective} slide (nothing will be written there)")

        if lob is not None:
            resclass = lob.resclass
            if effective == "Patterns" and not resclass:
                add("ERROR", "Patterns slide but the LoB has no ResClass sheet (no pattern charts available)")
            if effective == "Expert Judgement" and not resclass:
                add("ERROR", "Expert Judgement slide but the LoB has no ResClass (the Expert Judgement block is found by that code)")
            wants_life = resclass in life_resclasses
            for s in slots:
                if s.kind in LIFE_AWARE_KINDS and LIFE_AWARE_KINDS[s.kind] != wants_life:
                    add(
                        "WARN",
                        f"slot '{s.kind}' does not match the Life rule for ResClass '{resclass or '(blank)'}'"
                        f" (legacy would use the {'short' if wants_life else 'full'} range)",
                    )
        else:
            resclass = ""

        mapped.append(
            MappedSlide(sc.index, sc.title or "", lob.code if lob else p.code, lob.selector if lob else None, resclass, p.type, effective, slots)
        )

    # several slides carrying the same (LoB, type): legitimate (front section repeats the appendix), but say so.
    groups: dict[tuple[str, str], list[MappedSlide]] = {}
    for m in mapped:
        groups.setdefault((m.lob, m.type), []).append(m)
    for (lob, typ), ms in groups.items():
        if len(ms) > 1:
            same = len({tuple(sorted(s.kind for s in m.slots)) for m in ms}) == 1
            flags.append(
                Flag(
                    "INFO" if same else "WARN",
                    None,
                    lob,
                    f"{typ}: slides {[m.slide for m in ms]} share (LoB, type); all get the same images"
                    + ("" if same else " but their slot sets differ"),
                )
            )

    seen = {m.lob for m in mapped}
    missing = [lob.code for lob in lobs if lob.code not in seen]
    matches, table_flags = match_table_targets(scans, targets)
    return DeckMap(mapped, narrative, flags + table_flags, missing, matches)


# --------------------------------------------------------------------------- output
def matrix(dm: DeckMap) -> tuple[list[str], list[tuple[str, str, dict[str, list[int]]]]]:
    order: list[str] = []
    sel: dict[str, str] = {}
    cells: dict[str, dict[str, list[int]]] = {}
    for m in dm.mapped:
        if m.lob not in cells:
            order.append(m.lob)
            sel[m.lob] = m.selector or "?"
            cells[m.lob] = {t: [] for t in SLIDE_TYPES}
        cells[m.lob][m.type].append(m.slide)
    return list(SLIDE_TYPES), [(lob, sel[lob], cells[lob]) for lob in order]


def render_report(dm: DeckMap, template: str, n_slides: int) -> str:
    types, rows = matrix(dm)
    n_slots = sum(len(m.slots) for m in dm.mapped)
    count = {lvl: sum(1 for f in dm.flags if f.level == lvl) for lvl in ("ERROR", "WARN", "INFO")}
    out = [
        "# Template scan report (Step 3)",
        "",
        f"- Template: `{template}` ({n_slides} slides)",
        f"- LoB slides recognised: {len(dm.mapped)}; picture slots mapped: {n_slots}; "
        f"other slides (never touched): {len(dm.narrative)}",
        f"- Flags: {count['ERROR']} error, {count['WARN']} warning, {count['INFO']} info",
        "",
        "## LoB x slide type (slide numbers at scan time)",
        "",
        "| LoB | selector | " + " | ".join(types) + " |",
        "|---|---|" + "---|" * len(types),
    ]
    for lob, selector, cells in rows:
        out.append(f"| {lob} | {selector} | " + " | ".join(", ".join(map(str, cells[t])) or "-" for t in types) + " |")

    out += ["", "## Flags", ""]
    for lvl in ("ERROR", "WARN", "INFO"):
        items = [f for f in dm.flags if f.level == lvl]
        if items:
            out.append(f"### {lvl}")
            out += [
                f"- {'slide ' + str(f.slide) if f.slide else 'deck'} [{f.lob}]: {f.message}" for f in items
            ]
            out.append("")
    if not dm.flags:
        out += ["- none", ""]

    out += ["## Slide -> slots", ""]
    for m in dm.mapped:
        slots = ", ".join(f"{s.kind}(id {s.shape_id})" for s in m.slots) or "-"
        note = "" if m.type == m.title_type else f" [title says {m.title_type}]"
        out.append(f"- {m.slide}: {m.lob} / {m.type}{note}: {slots}")

    out += ["", "## Slides in the template that are not LoB slides (left untouched)", ""]
    for sc in dm.narrative:
        pics = [p for p in sc.pictures if p.kind != "logo"]
        extra = f"; {len(pics)} picture(s)" if pics else ""
        extra += f"; {', '.join(sc.other_shapes)}" if sc.other_shapes else ""
        out.append(f"- {sc.index}: {sc.title or '(no title)'}{extra}")

    if dm.lobs_without_slides:
        out += ["", "## LoBs in lobs.yaml with no slide in this template", "", ", ".join(dm.lobs_without_slides)]
    return "\n".join(out) + "\n"


def to_yaml_dict(dm: DeckMap, template: str) -> dict[str, Any]:
    return {
        "template": template,
        "note": "Generated by `resco scan`. Slide numbers are informational; the builder re-scans and matches by title + geometry.",
        "slides": [
            {
                "slide": m.slide,
                "title": m.title,
                "lob": m.lob,
                "selector": m.selector,
                "type": m.type,
                **({"title_type": m.title_type} if m.title_type != m.type else {}),
                "slots": [
                    {"kind": s.kind, "shape_id": s.shape_id, "shape_name": s.shape_name, "geometry": list(s.geometry), **({"new": True} if s.new else {})}
                    for s in m.slots
                ],
            }
            for m in dm.mapped
        ],
    }
