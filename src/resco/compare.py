"""Old-vs-new comparison for eyeballing: slide composites (pictures only) and per-slot image pairs, as a static HTML page.

PowerPoint is deliberately not used for rendering (the user may have the deck open); composites are drawn with Pillow
from the pictures and their slide positions, so text and native tables are not shown.
"""
from __future__ import annotations

import html
import io
from pathlib import Path
from typing import Any

from PIL import Image
from pptx import Presentation

from resco.deck_scan import EMU_PER_PT, _is_picture, picture_blob


def render_slide(prs: Any, index: int, scale: float = 1.0) -> Image.Image:
    """Draw the pictures of slide `index` of an already loaded presentation (loading a 16 MB deck is slow)."""
    slide = prs.slides[index - 1]
    w, h = prs.slide_width / EMU_PER_PT * scale, prs.slide_height / EMU_PER_PT * scale
    canvas = Image.new("RGBA", (round(w), round(h)), (255, 255, 255, 255))
    for sh in slide.shapes:
        blob = picture_blob(sh) if _is_picture(sh) else None
        if blob is None or None in (sh.left, sh.top, sh.width, sh.height):
            continue
        with Image.open(io.BytesIO(blob)) as im:
            box = tuple(round(v / EMU_PER_PT * scale) for v in (sh.left, sh.top, sh.width, sh.height))
            pic = im.convert("RGBA").resize((max(box[2], 1), max(box[3], 1)), Image.LANCZOS)
        canvas.alpha_composite(pic, (box[0], box[1]))
    return canvas.convert("RGB")


def _picture_blob(prs: Any, index: int, shape_id: int) -> bytes:
    for sh in prs.slides[index - 1].shapes:
        if sh.shape_id == shape_id and _is_picture(sh):
            blob = picture_blob(sh)
            if blob is not None:
                return blob
    raise KeyError((index, shape_id))


_CSS = """
:root{--bg:#fafaf9;--fg:#1c1917;--muted:#78716c;--card:#fff;--line:#e7e5e4;--ok:#15803d;--warn:#b45309;--err:#b91c1c;--accent:#9f1239}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#1c1917;--fg:#fafaf9;--muted:#a8a29e;--card:#292524;--line:#44403c;--ok:#4ade80;--warn:#fbbf24;--err:#f87171;--accent:#fb7185}}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 system-ui,sans-serif}
main{max-width:1200px;margin:0 auto;padding:24px 16px}
h1{font-size:22px;margin:0 0 4px}h2{font-size:16px;margin:28px 0 8px}
.sub{color:var(--muted);margin:0 0 16px}.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px;margin:0 0 12px}
.pair{display:grid;grid-template-columns:1fr 1fr;gap:12px}.pair figure{margin:0}.pair figcaption{color:var(--muted);font-size:12px;margin-bottom:4px}
.pair img{width:100%;border:1px solid var(--line);background:#fff}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{text-align:left;padding:4px 8px;border-bottom:1px solid var(--line)}
.ok{color:var(--ok)}.warn{color:var(--warn)}.err{color:var(--err)}.slots{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:8px}
.slots img{max-width:100%;border:1px solid var(--line);background:#fff}
@media (max-width:700px){.pair,.slots{grid-template-columns:1fr}}
"""


def write_compare(
    template: Path, built: Path, rows: list[dict[str, Any]], out_dir: Path, template_name: str
) -> Path:
    """`rows`: one dict per replaced slot (slide, title, lob, kind, shape_id, old_size, new_size, distortion_old/new)."""
    cdir = out_dir / "compare"
    cdir.mkdir(parents=True, exist_ok=True)
    prs_old, prs_new = Presentation(str(template)), Presentation(str(built))
    by_slide: dict[int, list[dict[str, Any]]] = {}
    for r in rows:
        by_slide.setdefault(r["slide"], []).append(r)

    parts = [
        "<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>",
        "<title>Build comparison</title>",
        f"<style>{_CSS}</style></head><body><main>",
        "<h1>Old vs new</h1>",
        f"<p class='sub'>Template <b>{html.escape(template_name)}</b> against the built deck. Composites show pictures only "
        "(titles, text and native tables are not drawn). Each picture keeps the slot position and size of the template.</p>",
    ]
    for slide, items in sorted(by_slide.items()):
        old_c, new_c = f"s{slide}_old.png", f"s{slide}_new.png"
        render_slide(prs_old, slide).save(cdir / old_c)
        render_slide(prs_new, slide).save(cdir / new_c)
        head = f"Slide {slide}: {items[0]['lob']} / {items[0]['type']}"
        parts.append(f"<h2>{html.escape(head)}</h2><div class='card'>")
        parts.append(
            f"<div class='pair'><figure><figcaption>Template (old)</figcaption><img src='{old_c}' alt='old slide {slide}'></figure>"
            f"<figure><figcaption>Built (new)</figcaption><img src='{new_c}' alt='new slide {slide}'></figure></div>"
        )
        parts.append("<table><tr><th>Slot</th><th>Old px</th><th>New px</th><th>Stretch in slot old</th><th>new</th></tr>")
        for r in sorted(items, key=lambda x: x["shape_id"]):
            d_old, d_new = r["distortion_old"], r["distortion_new"]
            cls = "warn" if d_old is not None and abs(d_new - d_old) > 10 else "ok"
            old_px = f"{r['old_size'][0]}x{r['old_size'][1]}" if r.get("old_size") else "(empty slide)"
            parts.append(
                f"<tr><td>{html.escape(r['kind'])}</td><td>{old_px}</td>"
                f"<td>{r['new_size'][0]}x{r['new_size'][1]}</td><td>{'-' if d_old is None else f'{d_old:+.0f}%'}</td><td class='{cls}'>{d_new:+.0f}%</td></tr>"
            )
        parts.append("</table><div class='slots'>")
        for r in sorted(items, key=lambda x: x["shape_id"]):
            oldf, newf = f"s{slide}_{r['shape_id']}_old.png", f"s{slide}_{r['shape_id']}_new.png"
            if not r.get("inserted"):
                (cdir / oldf).write_bytes(_picture_blob(prs_old, slide, r["shape_id"]))
            (cdir / newf).write_bytes(_picture_blob(prs_new, slide, r["shape_id"]))
            old_cell = "<div><div class='sub'>(no picture in the template)</div></div>" if r.get("inserted") else f"<div><div class='sub'>{html.escape(r['kind'])} old</div><img src='{oldf}' alt='old {html.escape(r['kind'])}'></div>"
            parts.append(old_cell + f"<div><div class='sub'>{html.escape(r['kind'])} new</div><img src='{newf}' alt='new {html.escape(r['kind'])}'></div>")
        parts.append("</div></div>")
    parts.append("</main></body></html>")
    path = cdir / "compare.html"
    path.write_text("".join(parts), encoding="utf-8")
    return path
