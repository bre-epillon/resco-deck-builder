"""Load config/*.yaml into plain dataclasses. No Excel or PowerPoint imports here."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

DEFAULT_CONFIG_DIR = Path(__file__).resolve().parents[2] / "config"


@dataclass(frozen=True)
class Lob:
    selector: str  # value written to the selector cell, e.g. "Summary_ENUPS"
    resclass: str  # sheet name / ExpertJudgement key; "" for totals
    table: bool
    triangles: bool
    profit: bool
    charts_triangles: bool
    patterns: bool
    exp_judgement: bool
    ej_first_row: int | None
    aliases: tuple[str, ...] = ()  # alternative codes seen in deck titles (e.g. a typo), mapped to this LoB
    code_override: str | None = None  # deck/report code when it differs from the selector (UK: Summary_ENUPS_UK -> ENUPS)
    check_code: str | None = None  # value the sheet's check cell must show (default: selector minus 'Summary_')

    @property
    def code(self) -> str:
        """LoB code used in deck titles and reports."""
        return self.code_override or self.selector.removeprefix("Summary_")

    @property
    def expected_check(self) -> str:
        """What the sheet's check cell must show after selecting this LoB (guards against a stale recalc)."""
        return self.check_code or self.selector.removeprefix("Summary_")


@dataclass(frozen=True)
class Ranges:
    selector_sheet: str
    selector_cell: str
    check_cell: str
    table: str
    table_life: str
    profit: str
    triangles: str
    chart_names_triangles: tuple[str, ...]
    chart_names_patterns: tuple[str, ...]
    chart_titles_patterns: tuple[str, ...]
    ej_sheet: str
    ej_blocks: tuple[tuple[str, str], ...]  # (name, "E:O")
    life_resclasses: frozenset[str]
    chart_by_kind: dict[str, str]  # slot kind -> chart object name on the selector sheet
    pattern_title_by_kind: dict[str, str]  # slot kind -> chart title on the ResClass sheet
    ej_scratch_sheet: str
    ej_clear: tuple[str, ...]
    ej_blocks_dest: tuple[str, ...]  # destination top-left per block, same order as ej_blocks
    ej_print: str
    ej_print_life: str
    chart_title_by_kind: dict[str, str] = field(default_factory=dict)  # slot kind -> title of the small triangle chart
    hide_rows: tuple[str, ...] = ()  # row spans on the selector sheet hidden in memory before export, e.g. '6:8' (never saved)
    ej_max_year: int | None = None  # fixed last UWY for the Expert Judgement block length (UK); default = year of 'asof'


@dataclass(frozen=True)
class SlotKind:
    kind: str
    geometry: tuple[float, float, float, float]  # left, top, width, height in points
    source: str


@dataclass(frozen=True)
class DeckRules:
    tolerance_pt: float
    kinds: dict[str, SlotKind]
    logos: tuple[tuple[float, float, float, float], ...]
    required: dict[str, tuple[tuple[str, ...], ...]]  # slide type -> groups of alternative kinds
    kind_type: dict[str, str]  # slot kind -> slide type it belongs to
    total_code: str


def load_deck_rules(config_dir: Path = DEFAULT_CONFIG_DIR) -> DeckRules:
    raw = yaml.safe_load((config_dir / "slot_kinds.yaml").read_text(encoding="utf-8"))
    kinds = {k: SlotKind(k, tuple(float(x) for x in v["geometry"]), v["source"]) for k, v in raw["kinds"].items()}
    return DeckRules(
        tolerance_pt=float(raw["tolerance_pt"]),
        kinds=kinds,
        logos=tuple(tuple(float(x) for x in g) for g in raw["logos"]),
        required={t: tuple(tuple(group) for group in v["required"]) for t, v in raw["slide_types"].items()},
        kind_type=dict(raw["kind_type"]),
        total_code=raw["total_code"],
    )


@dataclass(frozen=True)
class TableTarget:
    """A native table on a narrative slide that is filled from Excel (see config/tables.yaml)."""

    id: str
    kind: str  # ladder | overview
    slide: str  # slide title
    match: tuple[str, ...]  # words that must appear in the first two rows of the table
    select: dict[str, str] = field(default_factory=dict)  # ladder: LoB / level / basis / branch
    blocks: tuple[dict[str, Any], ...] = ()  # overview: how to find each source block


@dataclass(frozen=True)
class TableConfig:
    targets: tuple[TableTarget, ...] = ()
    netting_sheet: str = "Ladder"
    ibnr_sheet: str = "Overview"


def load_table_config(config_dir: Path = DEFAULT_CONFIG_DIR) -> TableConfig:
    f = config_dir / "tables.yaml"
    if not f.is_file():
        return TableConfig()
    raw = yaml.safe_load(f.read_text(encoding="utf-8")) or {}
    src = raw.get("sources", {})
    targets = tuple(
        TableTarget(
            id=t["id"], kind=t["kind"], slide=t["slide"], match=tuple(str(x).lower() for x in t.get("match", [])),
            select={k: str(v) for k, v in t.get("select", {}).items()}, blocks=tuple(t.get("blocks", [])),
        )
        for t in raw.get("targets", [])
    )  # fmt: skip
    return TableConfig(targets, src.get("netting", {}).get("sheet", "Ladder"), src.get("ibnr", {}).get("sheet", "Overview"))


@dataclass(frozen=True)
class NewSlide:
    """A slide the build adds when the template has none with this title (config/slides_add.yaml)."""

    id: str
    title: str
    like: tuple[str, ...]  # titles of model slides to copy, the first one found is used
    after: str | tuple[str, ...] | None  # title(s) of the slide it is placed after, the first one found counts (the end of the deck when none)
    copy: str  # frame (title, logo, slide number: the build inserts the pictures) | full (complete copy, e.g. with a native table)
    lob: str  # decides whether a LoB-limited run adds it
    part: str  # results | netting: the content that fills it (decides whether a content-limited run adds it)


def load_slides_to_add(config_dir: Path = DEFAULT_CONFIG_DIR) -> list[NewSlide]:
    f = config_dir / "slides_add.yaml"
    if not f.is_file():
        return []
    out: list[NewSlide] = []
    for d in (yaml.safe_load(f.read_text(encoding="utf-8")) or {}).get("add", []):
        like = d["like"]
        after = d.get("after")
        spec = NewSlide(d["id"], d["title"], (like,) if isinstance(like, str) else tuple(like), (after,) if isinstance(after, str) else (tuple(after) if after else None), d.get("copy", "frame"), str(d["lob"]).upper(), d.get("part", "results"))
        if spec.copy not in ("frame", "full") or spec.part not in ("results", "netting"):
            raise ValueError(f"{f.name}: '{spec.id}': copy must be frame|full and part results|netting")
        out.append(spec)
    return out


def load_lobs(config_dir: Path = DEFAULT_CONFIG_DIR) -> list[Lob]:
    raw = yaml.safe_load((config_dir / "lobs.yaml").read_text(encoding="utf-8"))
    return [
        Lob(
            selector=d["selector"],
            resclass=d["resclass"] or "",
            table=d.get("table", True),
            triangles=d.get("triangles", True),
            profit=d.get("profit", True),
            charts_triangles=d.get("charts_triangles", True),
            patterns=d.get("patterns", True),
            exp_judgement=d.get("exp_judgement", True),
            ej_first_row=d.get("ej_first_row"),
            aliases=tuple(str(a).upper() for a in d.get("aliases", [])),
            code_override=d.get("code"),
            check_code=d.get("check_code"),
        )
        for d in raw["lobs"]
    ]


def load_ranges(config_dir: Path = DEFAULT_CONFIG_DIR) -> Ranges:
    raw = yaml.safe_load((config_dir / "ranges_main.yaml").read_text(encoding="utf-8"))
    items = raw["resco_sheet_items"]
    ej = raw["expert_judgement"]
    return Ranges(
        selector_sheet=raw["selector"]["sheet"],
        selector_cell=raw["selector"]["cell"],
        check_cell=raw["selector"]["check_cell"],
        table=items["table"]["range"],
        table_life=items["table"]["range_life"],
        profit=items["profit"]["range"],
        triangles=items["triangles"]["range"],
        chart_names_triangles=tuple(c["chart"] for c in items["charts_triangles"]),
        chart_names_patterns=tuple(c["chart"] for c in raw["charts_patterns"]["charts"]),
        chart_titles_patterns=tuple(c["title"] for c in raw["charts_patterns"]["charts"]),
        ej_sheet=ej["source_sheet"],
        ej_blocks=tuple((b["name"], b["src_cols"]) for b in ej["blocks"]),
        life_resclasses=frozenset(raw["life_resclasses"]),
        chart_by_kind={c["kind"]: c["chart"] for c in items["charts_triangles"]},
        chart_title_by_kind={c["kind"]: c.get("title", "") for c in items["charts_triangles"]},
        pattern_title_by_kind={c["kind"]: c["title"] for c in raw["charts_patterns"]["charts"]},
        ej_scratch_sheet=ej["scratch_sheet"],
        ej_clear=tuple(ej["clear_before"]),
        ej_blocks_dest=tuple(b["dest"] for b in ej["blocks"]),
        ej_print=ej["print_range"],
        ej_print_life=ej["print_range_life"],
        ej_max_year=ej.get("max_year"),
        hide_rows=tuple(str(x) for x in raw["selector"].get("hide_rows", [])),
    )
