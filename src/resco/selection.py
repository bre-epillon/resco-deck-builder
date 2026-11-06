"""Which part of the deck a run touches: LoBs x content types x individual slides.

A run with no selection updates everything the template holds. A selection narrows it, e.g. "only ENTOT, only the net
(Gross vs Net) tables" = lobs ENTOT, parts netting. The three filters combine (a slide is touched when it passes all of
them), and the same code answers the dashboard's live preview, so the page and the build can never disagree.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, replace
from typing import Any

from resco.deck_scan import DeckMap, MappedSlide, TableMatch

# content types -> picture kinds (config/slot_kinds.yaml). The kind names themselves are accepted too (finer control).
GROUPS: dict[str, tuple[str, ...]] = {
    "results": ("table", "table_life", "profit"),
    "triangles": ("triangles", "chart_prem", "chart_paid", "chart_incurr"),
    "patterns": ("pattern_prem", "pattern_inc"),
    "ej": ("ej", "ej_life"),
}
# content types -> table target kinds (config/tables.yaml)
TABLE_GROUPS: dict[str, str] = {"netting": "ladder", "ibnr": "overview"}
ALIASES: dict[str, str] = {
    "result": "results", "triangle": "triangles", "pattern": "patterns", "expert_judgement": "ej", "expertjudgement": "ej",
    "net": "netting", "gvn": "netting", "gross_vs_net": "netting", "netting_tables": "netting", "ladder": "netting", "ibnr_tables": "ibnr",
}  # fmt: skip
KINDS = tuple(k for ks in GROUPS.values() for k in ks)
PART_LABEL = {
    "results": "Results (table + P&L)", "triangles": "Triangles (block + 3 charts)", "patterns": "Patterns (2 charts)", "ej": "Expert Judgement",
    "netting": "Gross vs Net tables (netting file)", "ibnr": "IBNR allocation / QoQ tables (IBNR file)",
}  # fmt: skip


class SelectionError(ValueError):
    pass


def _token(text: str) -> str:
    return re.sub(r"[\s\-]+", "_", text.strip().lower())


def parse_parts(text: str | None) -> frozenset[str]:
    out: set[str] = set()
    for raw in re.split(r"[,;]", text or ""):
        if not raw.strip():
            continue
        t = _token(raw)
        t = ALIASES.get(t, t)
        if t not in GROUPS and t not in TABLE_GROUPS and t not in KINDS:
            valid = ", ".join([*GROUPS, *TABLE_GROUPS])
            raise SelectionError(f"unknown part '{raw.strip()}': use {valid} (or a single picture kind such as profit, chart_prem)")
        out.add(t)
    return frozenset(out)


def parse_slides(text: str | None) -> frozenset[int]:
    out: set[int] = set()
    for raw in re.split(r"[,;\s]+", (text or "").strip()):
        if not raw:
            continue
        m = re.fullmatch(r"(\d{1,4})(?:\s*[-–]\s*(\d{1,4}))?", raw)
        if not m:
            raise SelectionError(f"cannot read '{raw}' as a slide number or range (use e.g. 16, 17-25)")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if a < 1 or b < a:
            raise SelectionError(f"bad slide range '{raw}'")
        out.update(range(a, b + 1))
    return frozenset(out)


def parse_lobs(text: str | None) -> tuple[str, ...]:
    seen: dict[str, None] = {}
    for raw in (text or "").split(","):
        if raw.strip():
            seen[raw.strip().upper()] = None
    return tuple(seen)


@dataclass(frozen=True)
class Selection:
    lobs: tuple[str, ...] = ()  # empty = every LoB that has slides
    parts: frozenset[str] = frozenset()  # empty = every content type
    slides: frozenset[int] = frozenset()  # empty = every slide

    @classmethod
    def parse(cls, lobs: str | None = None, parts: str | None = None, slides: str | None = None) -> Selection:
        return cls(parse_lobs(lobs), parse_parts(parts), parse_slides(slides))

    @property
    def active(self) -> bool:
        return bool(self.lobs or self.parts or self.slides)

    def picture_kinds(self) -> set[str] | None:
        """Picture kinds allowed; None = all."""
        if not self.parts:
            return None
        out: set[str] = set()
        for p in self.parts:
            out.update(GROUPS.get(p, (p,) if p in KINDS else ()))
        return out

    def table_kinds(self) -> set[str] | None:
        """Table target kinds allowed ('ladder' / 'overview'); None = all."""
        return None if not self.parts else {TABLE_GROUPS[p] for p in self.parts if p in TABLE_GROUPS}

    def allows_slide(self, number: int) -> bool:
        return not self.slides or number in self.slides

    def allows_new(self, lob: str, part: str) -> bool:
        """Whether a slide the build would add (LoB `lob`, content `part` = results | netting) belongs to this run."""
        if self.slides or (self.lobs and lob.upper() not in self.lobs):
            return False
        kinds = self.picture_kinds() if part == "results" else self.table_kinds()
        return kinds is None or bool(kinds & ({"table", "table_life"} if part == "results" else {"ladder"}))

    def restrict_plan(self, plan: list[MappedSlide]) -> list[MappedSlide]:
        """The mapped slides that survive the selection, each with only the picture slots that are selected."""
        kinds = self.picture_kinds()
        out: list[MappedSlide] = []
        for m in plan:
            if not self.allows_slide(m.slide):
                continue
            slots = [s for s in m.slots if kinds is None or s.kind in kinds]
            if slots:
                out.append(replace(m, slots=slots))
        return out

    def restrict_tables(self, matches: list[TableMatch], requested_lobs: list[str]) -> list[TableMatch]:
        """Ladder tables belong to one LoB (kept when that LoB is requested). IBNR tables are not LoB specific: they are kept
        when the run is not limited to some LoBs, or when 'ibnr' was asked for explicitly."""
        kinds = self.table_kinds()
        wanted = {w.upper() for w in requested_lobs}
        out: list[TableMatch] = []
        for m in matches:
            if not self.allows_slide(m.slide) or (kinds is not None and m.target.kind not in kinds):
                continue
            if m.target.kind == "ladder":
                if m.target.select.get("lob", "").upper() in wanted:
                    out.append(m)
            elif not self.lobs or (kinds is not None and "overview" in kinds):
                out.append(m)
        return out

    def describe(self) -> str:
        if not self.active:
            return "everything in the template"
        bits = []
        if self.lobs:
            bits.append("LoBs " + ", ".join(self.lobs))
        if self.parts:
            bits.append("only " + ", ".join(PART_LABEL.get(p, p) for p in sorted(self.parts, key=lambda p: (p not in PART_LABEL, p))))
        if self.slides:
            bits.append("slides " + _spans(sorted(self.slides)))
        return "; ".join(bits)

    def as_dict(self) -> dict[str, Any]:
        return {"lobs": list(self.lobs), "parts": sorted(self.parts), "slides": sorted(self.slides), "text": self.describe()}


def _spans(numbers: list[int]) -> str:
    out: list[str] = []
    i = 0
    while i < len(numbers):
        j = i
        while j + 1 < len(numbers) and numbers[j + 1] == numbers[j] + 1:
            j += 1
        out.append(str(numbers[i]) if i == j else f"{numbers[i]}-{numbers[j]}")
        i = j + 1
    return ", ".join(out)


# --------------------------------------------------------------------------- preview (dashboard)
def preview(dm: DeckMap, sel: Selection, will_add: list[dict[str, str]] | None = None) -> dict[str, Any]:
    """What a run with this selection would touch, slide by slide.

    `rows` = every slide that passes the LoB / content filters (the slides picker); `selected` = also passes the slide filter.
    """
    available = {m.lob for m in dm.mapped}
    requested = list(sel.lobs) or sorted(available)
    unknown = [c for c in sel.lobs if c not in available and c not in {a['lob'] for a in (will_add or [])}]  # a LoB whose slides the build adds is not unknown
    eligible = replace(sel, slides=frozenset())
    plan = eligible.restrict_plan([m for m in dm.mapped if m.lob in requested and m.selector])
    tables = eligible.restrict_tables(dm.table_matches, requested)
    rows: dict[int, dict[str, Any]] = {}
    for m in plan:
        rows[m.slide] = {"slide": m.slide, "title": m.title, "lob": m.lob, "pictures": [s.kind for s in m.slots], "tables": []}
    for t in tables:
        row = rows.setdefault(t.slide, {"slide": t.slide, "title": t.title, "lob": t.target.select.get("lob", ""), "pictures": [], "tables": []})
        row["tables"].append(t.target.id)
    for row in rows.values():
        row["selected"] = sel.allows_slide(row["slide"])
    ordered = sorted(rows.values(), key=lambda r: r["slide"])
    chosen = [r for r in ordered if r["selected"]]
    kinds_needing_excel = any(r["pictures"] for r in chosen) or any(t.target.kind == "overview" and sel.allows_slide(t.slide) for t in tables)
    return {
        "rows": ordered, "unknown_lobs": unknown, "available_lobs": sorted(available),
        "n_slides": len(chosen), "n_pictures": sum(len(r["pictures"]) for r in chosen), "n_tables": sum(len(r["tables"]) for r in chosen),
        "needs_excel": kinds_needing_excel, "text": sel.describe(), "will_add": will_add or [],
    }  # fmt: skip
