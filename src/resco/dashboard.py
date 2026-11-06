"""Dashboard pages. Self-contained HTML (data embedded, no external libraries).

- `render_dashboard`: one page per run (scan, dry run or build), written into the run folder. Works from a double-click
  (file://); images are referenced relatively inside the run folder. Through `resco serve` it also shows a link back to
  the home page and a switcher between runs of the same deck.
- `render_home`: the home page served at `/` by `resco serve`: start a run and browse the run history. It never shows
  results itself, so it cannot be mistaken for the result of the job you just started."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from resco.selection import PART_LABEL
from resco.runs import DECK_LABEL, KIND_LABEL, enrich

KIND_TYPE = {
    "table": "Results", "table_life": "Results", "profit": "Results",
    "triangles": "Triangles", "chart_prem": "Triangles", "chart_paid": "Triangles", "chart_incurr": "Triangles",
    "pattern_prem": "Patterns", "pattern_inc": "Patterns", "ej": "Expert Judgement", "ej_life": "Expert Judgement",
}  # fmt: skip


def render_dashboard(run_dir: Path, result: dict[str, Any], title: str = "ResCo deck builder") -> Path:
    result = enrich(result, run_dir)  # kind / deck / created also for runs written by older versions
    payload = {"title": title, "result": result, "kind_type": KIND_TYPE, "run_dir": run_dir.name, "deck_label": DECK_LABEL, "kind_label": KIND_LABEL}
    data = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    page = _PAGE.replace("/*__STYLE__*/", _STYLE).replace("/*__DATA__*/null", data)
    path = run_dir / "dashboard.html"
    path.write_text(page, encoding="utf-8")
    return path


def render_home() -> str:
    data = json.dumps({"deck_label": DECK_LABEL, "kind_label": KIND_LABEL, "parts": PART_LABEL}).replace("</", "<\\/")
    return _HOME.replace("/*__STYLE__*/", _STYLE).replace("/*__DATA__*/null", data)


_STYLE = r"""
:root{--bg:#faf9f7;--panel:#fff;--fg:#1c1917;--muted:#6b6560;--line:#e5e1dc;--accent:#9f1239;--ok:#15803d;--okbg:#dcfce7;--warn:#b45309;--warnbg:#fef3c7;--err:#b91c1c;--errbg:#fee2e2;--none:#f1efec;--lux:#9f1239;--luxbg:#fde8ec;--uk:#1d4ed8;--ukbg:#e0ecff;--shadow:0 1px 2px rgba(0,0,0,.06)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#171412;--panel:#221e1b;--fg:#f5f3f0;--muted:#a39d96;--line:#3a3430;--accent:#fb7185;--ok:#4ade80;--okbg:#14301f;--warn:#fbbf24;--warnbg:#352a0e;--err:#f87171;--errbg:#3a1616;--none:#2a2522;--lux:#fb7185;--luxbg:#3d1620;--uk:#93c5fd;--ukbg:#172a4a;--shadow:none}}
:root[data-theme="dark"]{--bg:#171412;--panel:#221e1b;--fg:#f5f3f0;--muted:#a39d96;--line:#3a3430;--accent:#fb7185;--ok:#4ade80;--okbg:#14301f;--warn:#fbbf24;--warnbg:#352a0e;--err:#f87171;--errbg:#3a1616;--none:#2a2522;--lux:#fb7185;--luxbg:#3d1620;--uk:#93c5fd;--ukbg:#172a4a;--shadow:none}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
header{background:var(--panel);border-bottom:1px solid var(--line);padding:12px 16px}
.wrap{max-width:1280px;margin:0 auto;padding:0 16px}header .wrap{padding:0;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
a{color:var(--accent)}h1{font-size:18px;margin:0}h2{font-size:15px;margin:0 0 10px}.sub{color:var(--muted);font-size:13px}
.badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:600;white-space:nowrap}
.b-ok{background:var(--okbg);color:var(--ok)}.b-warn{background:var(--warnbg);color:var(--warn)}.b-err{background:var(--errbg);color:var(--err)}.b-none{background:var(--none);color:var(--muted)}
.d-main{background:var(--luxbg);color:var(--lux)}.d-uk{background:var(--ukbg);color:var(--uk)}.deck{font-size:14px;padding:4px 14px}
main{padding:16px 0 48px}.grid{display:grid;gap:12px}.kpis{grid-template-columns:repeat(auto-fit,minmax(150px,1fr));margin-bottom:16px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:14px;box-shadow:var(--shadow)}
.kpi b{display:block;font-size:24px;line-height:1.1}.kpi span{color:var(--muted);font-size:12px}
nav{display:flex;gap:4px;margin:0 0 12px;flex-wrap:wrap}nav button{background:none;border:1px solid transparent;color:var(--muted);padding:6px 12px;border-radius:6px;cursor:pointer;font:inherit}
nav button[aria-selected=true]{background:var(--panel);border-color:var(--line);color:var(--fg);font-weight:600}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{text-align:left;padding:6px 8px;border-bottom:1px solid var(--line);vertical-align:top}th{color:var(--muted);font-weight:600;position:sticky;top:0;background:var(--panel)}
.scroll{overflow:auto;max-height:640px}.mx td,.mx th{text-align:center}.mx td:first-child,.mx th:first-child{text-align:left}
.cell{display:inline-block;min-width:54px;padding:3px 8px;border-radius:6px;border:1px solid transparent;cursor:pointer;font:inherit;font-size:12px}
.cell.ok{background:var(--okbg);color:var(--ok)}.cell.warn{background:var(--warnbg);color:var(--warn)}.cell.err{background:var(--errbg);color:var(--err)}.cell.none{background:var(--none);color:var(--muted);cursor:default}
.cell:focus-visible,button:focus-visible,select:focus-visible,input:focus-visible,a:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.lvl-ERROR{color:var(--err);font-weight:600}.lvl-WARN{color:var(--warn);font-weight:600}.lvl-INFO{color:var(--muted)}
.row{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
select,input[type=text]{background:var(--panel);color:var(--fg);border:1px solid var(--line);border-radius:6px;padding:5px 8px;font:inherit}
.btn{background:var(--accent);color:#fff;border:0;border-radius:6px;padding:7px 14px;font:inherit;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}.btn.sec{background:var(--none);color:var(--fg);border:1px solid var(--line)}
.btn:disabled{opacity:.5;cursor:not-allowed}
.pair{display:grid;grid-template-columns:1fr 1fr;gap:12px}.pair img,.slots img{width:100%;border:1px solid var(--line);background:#fff;border-radius:4px}.pair figcaption{font-size:12px;color:var(--muted);margin-bottom:3px}
.slots{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px}.slots .t{font-size:12px;color:var(--muted)}
pre{background:var(--none);padding:10px;border-radius:6px;max-height:280px;overflow:auto;font-size:12px;margin:8px 0 0;white-space:pre-wrap}
.hint{color:var(--muted);font-size:12px}.hide{display:none!important}.err-text{color:var(--err);font-weight:600}
.bar{height:8px;background:var(--none);border-radius:4px;overflow:hidden;margin-top:6px}.bar i{display:block;height:100%;background:var(--accent);width:0}
.pill{border:1px solid var(--line);border-radius:99px;padding:6px 16px;cursor:pointer;display:inline-flex;gap:8px;align-items:center;background:var(--panel)}
.pill:has(input:checked){border-color:var(--accent);background:var(--none);font-weight:600}
fieldset{border:0;padding:0;margin:0 0 10px}legend{font-size:12px;color:var(--muted);padding:0;margin-bottom:4px}
.chg{background:var(--warnbg)}.gridtbl td,.gridtbl th{text-align:right;padding:2px 6px;font-size:12px;white-space:nowrap}.gridtbl td:first-child,.gridtbl th:first-child{text-align:left}.trow{cursor:pointer}.trow:hover{background:var(--none)}
#picker td,#picker th{padding:3px 8px}#r-parts{gap:6px}
.actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px;margin-top:6px}.actions .card{padding:10px}.actions p{margin:6px 0 0}
@media (max-width:760px){.pair,.slots{grid-template-columns:1fr}}
"""

_PAGE = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ResCo deck builder</title>
<style>/*__STYLE__*/</style></head><body>
<header><div class="wrap">
<a id="home" class="btn sec hide" href="/">← All runs</a><h1 id="title"></h1>
<span id="deckb" class="badge deck"></span><span id="kindb" class="badge b-none"></span><span id="partb" class="badge b-warn hide"></span><span id="status" class="badge"></span>
<span class="sub" id="meta"></span><span style="flex:1"></span>
<select id="runsel" class="hide" aria-label="Switch to another run of this deck"></select>
<button class="btn sec" id="theme" aria-label="Toggle light/dark theme">Theme</button></div></header>
<main><div class="wrap">
<div class="grid kpis" id="kpis"></div>
<nav role="tablist" id="tabs"></nav>
<section id="tab-overview"></section><section id="tab-issues" class="hide"></section><section id="tab-checks" class="hide"></section><section id="tab-slots" class="hide"></section><section id="tab-tables" class="hide"></section>
<section id="detail" class="card hide" style="margin-top:16px"></section>
</div></main>
<script>
const D = /*__DATA__*/null; const R = D.result; const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const TYPES = ['Results','Triangles','Patterns','Expert Judgement'];
const kind = R.kind, deck = R.deck, isScan = kind === 'scan', isDry = kind === 'dry-run', isCheck = kind === 'check';
const when = iso => { const d = new Date(iso); return isNaN(d) ? iso : d.toLocaleString([], {day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'}); };
const DECKNAME = D.deck_label[deck] || deck, KINDNAME = D.kind_label[kind] || kind;
document.title = `${DECKNAME} · ${KINDNAME} · ResCo deck builder`;
$('#title').textContent = 'ResCo deck builder';
$('#deckb').className = 'badge deck d-' + deck; $('#deckb').textContent = DECKNAME; $('#kindb').textContent = KINDNAME;
const SEL = R.selection || {}, partial = !!((SEL.lobs || []).length || (SEL.parts || []).length || (SEL.slides || []).length);
if (partial) { $('#partb').textContent = 'Partial run'; $('#partb').title = SEL.text; $('#partb').classList.remove('hide'); }
$('#meta').textContent = [R.template, R.workbook && R.workbook !== '-' ? R.workbook : '', partial ? 'only: ' + SEL.text : '', R.template_modified ? 'template saved ' + when(R.template_modified) : '', when(R.created), R.seconds != null ? R.seconds + ' s' : ''].filter(Boolean).join(' · ');
$('#theme').onclick = () => { const r = document.documentElement; const cur = r.dataset.theme || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light'); r.dataset.theme = cur === 'dark' ? 'light' : 'dark'; };

// ---- derived data
const issues = [].concat(
  (R.issues || []).map(i => ({level: i.level, lob: i.lob, kind: i.kind, message: i.message, src: 'export/checks'})),
  (R.scan_flags || []).filter(f => isScan || f.level !== 'INFO').map(f => ({level: f.level, lob: f.lob, kind: f.slide ? 'slide ' + f.slide : 'deck', message: f.message, src: 'template scan'})),
  ((R.verify && R.verify.messages) || []).map(m => ({level: m[0], lob: '-', kind: 'deck check', message: m[1], src: 'verification'}))
);
const cnt = l => issues.filter(i => i.level === l).length;
const slots = []; (R.slides || []).forEach(s => s.slots.forEach(x => slots.push(Object.assign({slide: s.slide, lob: s.lob, type: s.type, title: s.title}, x))));
const replaced = slots.filter(x => x.status === 'replaced' || x.status === 'inserted').length, nInserted = slots.filter(x => x.status === 'inserted').length, failed = slots.filter(x => String(x.status).startsWith('failed')).length;
const verOk = R.verify ? R.verify.ok : null;
const tbls = R.tables || [];
let state = 'ok', label = isScan ? 'Deck recognised' : (isCheck ? 'Workbook ready' : 'All good');
if (isDry) { state = 'none'; label = 'Dry run: nothing written'; }
else if (cnt('ERROR') || failed || verOk === false) {
  const dataErr = issues.filter(i => i.level === 'ERROR' && /^totals check failed/.test(i.message)).length;
  state = 'err'; label = isScan ? 'Problems in the deck' : isCheck ? 'Problems found in the workbook' : ((dataErr === cnt('ERROR') && !failed && verOk !== false) ? 'Deck built: workbook data checks failed' : 'Needs attention'); }
else if (cnt('WARN')) { state = 'warn'; label = 'Review warnings'; }
$('#status').className = 'badge b-' + state; $('#status').textContent = label;
const kpi = (v, l) => `<div class="card kpi"><b>${esc(v)}</b><span>${esc(l)}</span></div>`;
const found = slots.filter(x => x.status === 'found').length;
$('#kpis').innerHTML = isCheck
  ? kpi((R.slides || []).length, 'slides in scope') + kpi(found + ' / ' + slots.length, 'picture slots ready') + kpi(cnt('ERROR'), 'errors') + kpi(cnt('WARN'), 'warnings') + kpi(R.seconds != null ? R.seconds + ' s' : '-', 'check took')
  : isScan
  ? kpi((R.slides || []).length, 'LoB slides recognised') + kpi(slots.length, 'picture slots found') + kpi(cnt('ERROR'), 'errors') + kpi(cnt('WARN'), 'warnings') + kpi(R.other_slides ?? '-', 'other slides (never touched)')
  : isDry
  ? kpi((R.slides || []).length, 'slides in scope') + kpi(slots.length, 'pictures planned') + kpi(cnt('ERROR'), 'errors') + kpi(cnt('WARN'), 'warnings')
  : kpi((R.slides || []).length, 'slides in scope') + kpi(replaced + ' / ' + slots.length, nInserted ? `pictures replaced (${nInserted} inserted into empty slides)` : 'pictures replaced') +
    kpi(cnt('ERROR'), 'errors') + kpi(cnt('WARN'), 'warnings') + kpi(verOk === null ? 'n/a' : (verOk ? 'passed' : 'FAILED'), 'deck verification') +
    kpi(R.verify ? R.verify.stats.untouched_slides_identical : '-', 'untouched slides identical') + (tbls.length ? kpi(tbls.filter(t => t.status === 'filled').length + ' / ' + tbls.length, 'slide tables filled from Excel') : '');

// ---- tabs
const tabs = [['overview', 'Overview'], ['issues', `Issues (${issues.length})`], ['checks', 'Data checks'], ['slots', 'All slots']];
if (tbls.length) tabs.splice(1, 0, ['tables', `Tables from Excel (${tbls.length})`]);
$('#tabs').innerHTML = tabs.map(([k, t], i) => `<button role="tab" data-k="${k}" aria-selected="${i === 0}">${esc(t)}</button>`).join('');
$('#tabs').onclick = e => { const b = e.target.closest('button'); if (!b) return; document.querySelectorAll('#tabs button').forEach(x => x.setAttribute('aria-selected', x === b)); tabs.forEach(([k]) => $('#tab-' + k).classList.toggle('hide', k !== b.dataset.k)); };

// ---- overview matrix
function cellState(lob, type) {
  const ss = (R.slides || []).filter(s => s.lob === lob && s.type === type); if (!ss.length) return null;
  const sl = ss.flatMap(s => s.slots); let st = 'ok';
  if (sl.some(x => String(x.status).startsWith('failed'))) return {st: 'err', ss};
  const rel = issues.filter(i => i.lob === lob && (D.kind_type[i.kind] === type || String(i.kind).startsWith('slide')));
  if (rel.some(i => i.level === 'ERROR')) st = 'err'; else if (rel.some(i => i.level === 'WARN')) st = 'warn';
  if (isDry) st = 'none2';
  return {st, ss};
}
const lobs = [...new Set((R.slides || []).map(s => s.lob))];
let mx = `<div class="card"><h2>${esc(DECKNAME)}: LoB × slide type</h2><div class="scroll"><table class="mx"><tr><th>LoB</th>${TYPES.map(t => `<th>${t}</th>`).join('')}</tr>`;
lobs.forEach(l => { mx += `<tr><td><b>${esc(l)}</b></td>` + TYPES.map(t => { const c = cellState(l, t);
  if (!c) return '<td><span class="cell none">–</span></td>';
  const cls = c.st === 'none2' ? 'none' : c.st; return `<td><button class="cell ${cls}" data-lob="${esc(l)}" data-type="${esc(t)}">${c.ss.map(s => s.slide).join(', ')}</button></td>`; }).join('') + '</tr>'; });
mx += `</table></div><p class="hint">Cell = slide number(s) at scan time. ${isScan || isDry || isCheck ? 'Green: no warnings · amber: warning on this LoB · red: error.' : 'Green: replaced, no warnings · amber: warning on this LoB · red: failed.'} Click a cell for details${isScan || isDry || isCheck ? '' : ' and the old-vs-new comparison'}.</p></div>`;
const fileLinks = isCheck ? [['build_report.md', 'Report (markdown)', true], ['build_result.json', 'Raw results (json)', true]] : isScan ? [['scan_report.md', 'Scan report (markdown)', true], ['slides.yaml', 'Slide → picture-slot mapping (yaml)', true], ['scan_result.json', 'Raw results (json)', true]]
  : [['compare/compare.html', 'Old vs new (all slides)', R.compare_html], [R.output_deck, 'Built deck (.pptx)', R.output_deck], ['build_report.md', 'Report (markdown)', true], ['build_result.json', 'Raw results (json)', true], ['images/', 'Exported pictures', !isDry]];
const links = fileLinks.filter(x => x[2]).map(x => `<a href="${esc(x[0])}">${esc(x[1])}</a>`).join(' · ');
function regressionCard() {
  const g = R.regressions || []; if (!g.length) return '';
  return `<div class="card" style="margin-bottom:12px;border-color:var(--err)"><h2 class="err-text">Compared with recent full builds, ${g.length} thing(s) are missing in this run</h2><ul>${g.map(x => `<li>${esc(x.kind === 'template' ? x.item : x.kind + ' ' + x.item)} <span class="hint">(last seen in ${esc(x.last_seen)})</span></li>`).join('')}</ul><p class="hint">Usually this is not the template you meant (an older copy?) or a slide was removed. This check only reports; it changed nothing.</p></div>`;
}
function changesCard() {
  const c = R.template_changes; if (!c || (!c.added.length && !c.removed.length)) return '';
  const li = (a, sign) => a.slice(0, 12).map(t => `<li>${sign} ${esc(t)}</li>`).join('') + (a.length > 12 ? `<li class="hint">… and ${a.length - 12} more</li>` : '');
  return `<div class="card" style="margin-bottom:12px"><h2>The template changed since the previous full build</h2><p class="hint">Compared with ${esc(c.since)}. The tool re-reads the deck on every run and finds everything by title and picture position, so nothing needs updating; this is for your information.</p><ul>${li(c.added, '+')}${li(c.removed, '−')}</ul></div>`;
}
function addedCard() {
  const a = (R.added_slides || []), todo = (R.slides_to_add || []);
  if (!a.length && !todo.length) return '';
  const rows = a.map(e => `<tr><td>${e.position || '–'}</td><td>${esc(e.title)}</td><td class="hint">${esc(e.like || '')}</td><td class="hint">${esc(e.after || 'end of deck')}</td><td>${esc(e.status)}</td></tr>`).join('');
  const lead = a.length ? (isDry ? 'A build would add these slides (the template does not have them):' : 'Added to the deck (the template does not have them):') : 'A build adds these slides (the template does not have them):';
  return `<div class="card" style="margin-top:12px"><h2>Slides added by the build</h2><p class="hint">${lead} each is a copy of a model slide, retitled; the pictures and tables are filled like on any other slide. The template itself is never changed.</p>` +
    (a.length ? `<table><tr><th>Slide</th><th>Title</th><th>Copy of</th><th>After</th><th>Status</th></tr>${rows}</table>` : `<ul>${todo.map(t => `<li>${esc(t)}</li>`).join('')}</ul>`) + '</div>';
}
function missingCard() {
  const m = R.lobs_without_slides || []; if (!m.length) return '';
  return `<div class="card" style="margin-top:12px"><h2>LoBs without a slide in this deck</h2><p><b>${m.map(esc).join(', ')}</b></p><p class="hint">These LoBs are configured (config/lobs.yaml) but the template has no slide titled “… (CODE) - Results / Triangles / …” for them, so nothing is produced for them. Add a slide with such a title (it can be empty) and the next run inserts the pictures.</p></div>`;
}
function layoutCard() {
  const L = R.layout; if (!L) return '';
  const rows = [['table', 'Results table'], ['table_short', 'Results table, first block only (Life classes)'], ['profit', 'P&L box'], ['triangles', 'Triangles']].map(([k, t]) => { const it = L[k];
    return `<tr><td>${t}</td><td><b>${esc(it.address)}</b></td><td class="hint">${esc(it.configured)}</td><td>${it.moved ? '<span class="badge b-warn">moved: the found range is used</span>' : '<span class="badge b-ok">as configured</span>'}</td></tr>`; }).join('');
  const ej = Object.entries(L.expert_judgement || {}).map(([c, b]) => `${esc(c)} ${b.rows[0]}–${b.rows[1]}`).join(' · ');
  return `<div class="card" style="margin-top:12px"><h2>Where things were found in the workbook</h2><table><tr><th>What</th><th>Address used</th><th>Configured</th><th></th></tr>${rows}</table>
    <p class="hint">Found by their labels (the UWY header, the Total rows, the 'Profit (loss) movement QoQ' title, the 'Year 1' row), not by fixed addresses, so a longer or re-ordered sheet is followed. Expert Judgement blocks are found by LoB code (rows used):</p><p class="hint">${ej || 'none found'}</p></div>`;
}
$('#tab-overview').innerHTML = regressionCard() + changesCard() + mx + addedCard() + missingCard() + layoutCard() + `<div class="card" style="margin-top:12px"><h2>Files of this run</h2><div>${links}</div><p class="hint">${R.workbook_sha256 ? 'Workbook copy fingerprint (sha256, first 12): ' + esc(R.workbook_sha256) : ''}</p></div>`;
$('#tab-overview').onclick = e => { const b = e.target.closest('button.cell'); if (b && b.dataset.lob) showDetail(b.dataset.lob, b.dataset.type); };

function img(name) { return `compare/${name}`; }
function showDetail(lob, type) {
  const ss = (R.slides || []).filter(s => s.lob === lob && s.type === type); const box = $('#detail'); box.classList.remove('hide');
  let h = `<h2>${esc(lob)} · ${esc(type)}</h2>`;
  ss.forEach(s => {
    h += `<h3 style="font-size:14px;margin:14px 0 6px">Slide ${s.slide}: ${esc(s.title)}${s.title_type !== s.type ? ` <span class="badge b-warn">title says ${esc(s.title_type)}</span>` : ''}</h3>`;
    if (R.compare_html && !isDry && !isScan && !isCheck) h += `<div class="pair"><figure><figcaption>Template (old)</figcaption><img src="${img('s' + s.slide + '_old.png')}" alt="old slide ${s.slide}"></figure><figure><figcaption>Built (new)</figcaption><img src="${img('s' + s.slide + '_new.png')}" alt="new slide ${s.slide}"></figure></div>`;
    h += `<table style="margin-top:8px"><tr><th>Slot</th><th>Status</th><th>Old px</th><th>New px</th><th>Stretch old → new / source</th></tr>` + s.slots.map(x => `<tr><td>${esc(x.kind)}${x.new ? ' <span class="hint">(new picture)</span>' : ''}</td><td>${esc(x.status)}</td><td>${x.old_size ? x.old_size.join('×') : '–'}</td><td>${x.new_size ? x.new_size.join('×') : '–'}</td><td>${x.distortion_old != null ? x.distortion_old.toFixed(0) + '% → ' + x.distortion_new.toFixed(0) + '%' : esc(x.source || '–')}</td></tr>`).join('') + '</table>';
    const rel = issues.filter(i => i.lob === lob && (D.kind_type[i.kind] === type || i.kind === 'slide ' + s.slide));
    if (rel.length) h += '<ul>' + rel.map(i => `<li><span class="lvl-${i.level}">${i.level}</span> ${esc(i.message)}</li>`).join('') + '</ul>';
  });
  box.innerHTML = h; box.scrollIntoView({behavior: 'smooth', block: 'start'});
}

// ---- issues
function renderIssues() {
  const lv = $('#f-lvl') ? $('#f-lvl').value : '', lb = $('#f-lob') ? $('#f-lob').value : '', q = $('#f-q') ? $('#f-q').value.toLowerCase() : '';
  const rows = issues.filter(i => (!lv || i.level === lv) && (!lb || i.lob === lb) && (!q || (i.message + i.kind).toLowerCase().includes(q)));
  $('#issue-rows').innerHTML = rows.map(i => `<tr><td class="lvl-${i.level}">${i.level}</td><td>${esc(i.lob)}</td><td>${esc(i.kind)}</td><td>${esc(i.message)}</td><td class="hint">${esc(i.src)}</td></tr>`).join('') || '<tr><td colspan="5" class="hint">Nothing matches.</td></tr>';
}
$('#tab-issues').innerHTML = `<div class="card"><div class="row"><select id="f-lvl" aria-label="Level"><option value="">All levels</option><option>ERROR</option><option>WARN</option><option>INFO</option></select>
 <select id="f-lob" aria-label="LoB"><option value="">All LoBs</option>${[...new Set(issues.map(i => i.lob))].sort().map(l => `<option>${esc(l)}</option>`).join('')}</select>
 <input type="text" id="f-q" placeholder="Search" aria-label="Search issues"></div><div class="scroll"><table><tr><th>Level</th><th>LoB</th><th>Where</th><th>Message</th><th>Source</th></tr><tbody id="issue-rows"></tbody></table></div></div>`;
['f-lvl', 'f-lob', 'f-q'].forEach(id => $('#' + id).addEventListener('input', renderIssues)); renderIssues();

// ---- checks
const ch = R.checks || {}; $('#tab-checks').innerHTML = `<div class="card"><h2>Data checks per LoB</h2><div class="scroll"><table><tr><th>LoB</th><th>LoB code on sheet</th><th>Totals vs components</th><th>Max relative difference</th><th>Error values visible in images</th><th>Export time</th></tr>` +
  Object.keys(ch).map(l => { const c = ch[l], t = c.totals || {}; const ok = t.status === 'pass' && (c.c5 === 'ok');
  return `<tr><td><b>${esc(l)}</b></td><td>${esc(c.c5)}</td><td><span class="badge ${ok ? 'b-ok' : (t.status === 'fail' ? 'b-err' : 'b-warn')}">${esc(t.status || '–')}</span> <span class="hint">${t.comparisons || 0} comparisons</span></td><td>${t.max_rel_diff != null ? t.max_rel_diff.toExponential(1) : '–'}</td><td>${(c.visible_errors || []).length}</td><td>${(R.seconds_per_lob || {})[l] ?? '–'} s</td></tr>`; }).join('') +
  `</table></div>${Object.keys(ch).length ? '' : '<p class="hint">No data checks in this run: they only run in a build (a scan, dry run or workbook check does not export or check numbers).</p>'}<p class="hint">Totals: every amount column of the Results table must equal the sum of its rows (relative tolerance 1e-6). “LoB code on sheet” confirms Excel recalculated for the selected LoB before anything was exported.</p></div>`;

// ---- all slots
function renderSlots() {
  const q = ($('#s-q').value || '').toLowerCase(), st = $('#s-st').value;
  $('#slot-rows').innerHTML = slots.filter(x => (!st || (st === 'failed' ? String(x.status).startsWith('failed') : x.status === st)) && (!q || (x.lob + x.type + x.kind + x.slide).toLowerCase().includes(q))).map(x =>
   `<tr><td>${x.slide}</td><td>${esc(x.lob)}</td><td>${esc(x.type)}</td><td>${esc(x.kind)}${x.new ? ' <span class="hint">(new picture)</span>' : ''}</td><td>${esc(x.status)}</td><td>${x.old_size ? x.old_size.join('×') : '–'}</td><td>${x.new_size ? x.new_size.join('×') : '–'}</td><td>${x.distortion_old != null ? x.distortion_old.toFixed(0) + '% → ' + x.distortion_new.toFixed(0) + '%' : '–'}</td></tr>`).join('');
}
$('#tab-slots').innerHTML = `<div class="card"><div class="row"><input type="text" id="s-q" placeholder="Filter (LoB, type, slot, slide)" aria-label="Filter slots"><select id="s-st" aria-label="Status"><option value="">All statuses</option><option>replaced</option><option>inserted</option><option>failed</option><option>planned</option><option>recognised</option><option>found</option></select></div><div class="scroll"><table><tr><th>Slide</th><th>LoB</th><th>Type</th><th>Slot</th><th>Status</th><th>Old px</th><th>New px</th><th>Stretch old → new</th></tr><tbody id="slot-rows"></tbody></table></div></div>`;
['s-q', 's-st'].forEach(id => $('#' + id).addEventListener('input', renderSlots)); renderSlots();

// ---- tables filled from Excel (native slide tables): list + before/after grid
function renderTables() {
  const cls = s => String(s).startsWith('failed') ? 'b-err' : (s === 'filled' || s === 'checked' ? 'b-ok' : 'b-none');
  $('#tab-tables').innerHTML = `<div class="card"><h2>Native slide tables filled from Excel</h2><p class="hint">Only the text of the cells is replaced; the table keeps the deck's styling. Rows are matched by UWY label and columns by header text. Click a row to see every cell before and after (changed cells are highlighted).</p>
    <div class="scroll"><table><tr><th>Slide</th><th>Table</th><th>Source</th><th>Status</th><th>Figures changed</th><th>Unchanged</th><th>Not matched</th></tr>` +
    tbls.map((t, i) => `<tr class="trow" data-i="${i}"><td>${t.slide}</td><td>${esc(t.id)} <span class="hint">(${esc(t.table)}, ${t.rows}×${t.cols})</span></td><td class="hint">${esc(t.source)}</td><td><span class="badge ${cls(t.status)}">${esc(String(t.status).slice(0, 70))}</span></td><td>${t.changed ?? '–'}</td><td>${t.unchanged ?? '–'}</td><td>${esc([...(t.unmatched_rows || []), ...(t.unmatched_cols || [])].join(', ') || '–')}</td></tr>`).join('') +
    `</table></div></div><div id="tbl-detail"></div>`;
  $('#tab-tables').querySelectorAll('.trow').forEach(tr => tr.onclick = () => showTable(+tr.dataset.i));
}
function showTable(i) {
  const t = tbls[i]; if (!t.grid) { $('#tbl-detail').innerHTML = `<div class="card" style="margin-top:12px"><p class="hint">No cell grid for this table (${esc(t.status)}).</p></div>`; return; }
  const body = t.grid.map(row => '<tr>' + row.map(([o, n]) => `<td class="${o !== n ? 'chg' : ''}" title="${o !== n ? 'was: ' + esc(o) : ''}">${esc(n)}</td>`).join('') + '</tr>').join('');
  $('#tbl-detail').innerHTML = `<div class="card" style="margin-top:12px"><h2>Slide ${t.slide}: ${esc(t.id)} after filling</h2><div class="scroll"><table class="gridtbl">${body}</table></div><p class="hint">${t.changed} figure(s) changed, ${t.unchanged} identical to the template. Hover a highlighted cell to see its previous text.</p></div>`;
  $('#tbl-detail').scrollIntoView({behavior: 'smooth', block: 'start'});
}
if (tbls.length) renderTables();

// ---- served mode only: way back to the home page, and a switcher between runs of THIS deck
(async () => {
  let j; try { j = await (await fetch('/api/runs', {cache: 'no-store'})).json(); } catch (e) { return; }
  $('#home').classList.remove('hide');
  const same = (j.runs || []).filter(r => r.deck === deck && r.has_dashboard);
  $('#runsel').innerHTML = same.map(r => `<option value="${esc(r.name)}" ${r.name === D.run_dir ? 'selected' : ''}>${esc(when(r.created))} · ${esc(r.kind_label)} · ${esc(r.label)}</option>`).join('');
  $('#runsel').classList.toggle('hide', same.length < 2);
  $('#runsel').onchange = e => { location.href = '/' + encodeURIComponent(e.target.value) + '/dashboard.html'; };
})();
</script></body></html>
"""

_HOME = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ResCo deck builder: all runs</title>
<style>/*__STYLE__*/</style></head><body>
<header><div class="wrap"><h1>ResCo deck builder</h1><span class="sub">Start a run or open an earlier one. Results are shown on the page of each run.</span>
<span style="flex:1"></span><button class="btn sec" id="theme" aria-label="Toggle light/dark theme">Theme</button></div></header>
<main><div class="wrap">
<section class="card" style="margin-bottom:16px"><h2>New run</h2>
 <fieldset><legend>Deck</legend>
  <label class="pill"><input type="radio" name="deck" value="main" checked> Lux (main deck)</label>
  <label class="pill"><input type="radio" name="deck" value="uk"> UK branch deck</label></fieldset>
 <fieldset id="what"><legend>What to update (leave everything as it is to update the whole deck)</legend>
  <div class="row"><label>LoBs <input type="text" id="r-lobs" placeholder="all, or e.g. ENTOT or ENUPS,ALL" size="28"></label> <span class="hint">codes as in the slide titles</span></div>
  <div class="row" id="r-parts" role="group" aria-label="Content types"></div>
  <div class="row"><label>Slides <input type="text" id="r-slides" placeholder="all, or e.g. 16, 17, 55-58" size="28"></label> <button class="btn sec" type="button" id="b-pick">Pick slides…</button></div>
  <div id="sel-sum" class="hint" aria-live="polite"></div>
  <div id="picker" class="card hide" style="margin-top:8px"><div class="row"><button class="btn sec" type="button" id="p-all">Tick all</button><span class="hint">Only the slides that pass the LoB and content filters are listed; untick the ones to leave alone.</span></div>
   <div class="scroll" style="max-height:320px"><table id="p-tbl"></table></div></div>
 </fieldset>
 <div class="row"><label>Workbook <input type="text" id="r-wb" size="70"></label></div>
 <div class="row"><label>Template deck <input type="text" id="r-tpl" size="70"></label></div>
 <div class="row"><label>Netting workbook <input type="text" id="r-net" size="70" placeholder="optional: fills the Gross vs Net summary tables"></label></div>
 <div class="row"><label>Netting Q-1 <input type="text" id="r-netq1" size="70" placeholder="UK deck: Q-1 ladder as presented (Inputs/Q-1_netting_UK_….xlsx), fills the Profit (loss) movement QoQ boxes"></label></div>
 <div class="row"><label>IBNR allocation <input type="text" id="r-ibnr" size="70" placeholder="optional: fills the Earned IBNR allocation and QoQ movement tables"></label></div>
 <div id="tpl-hint" class="hint err-text hide"></div>
 <div id="tpl-info" class="hint"></div>
 <div class="actions">
  <div class="card"><button class="btn sec" id="b-scan">Scan template</button><p class="hint">Reads the template deck only: which LoB slides and picture slots the tool recognises. No Excel.</p></div>
  <div class="card"><button class="btn sec" id="b-check">Check workbook</button><p class="hint">Opens a copy of the workbook (about 40 s, no pictures) and checks that everything the build needs is there: ranges, Expert Judgement blocks, charts. Run it first when the workbook or the deck changed.</p></div>
  <div class="card"><button class="btn sec" id="b-dry">Dry run</button><p class="hint">The scan plus the plan: which picture goes where. No Excel, nothing is written.</p></div>
  <div class="card"><button class="btn" id="b-build">Build deck</button><p class="hint">Full run: exports from Excel, writes a copy of the deck and checks everything. About 7 min (Lux) or 4 min (UK).</p></div>
 </div>
 <p class="row" style="margin:10px 0 0"><span id="r-msg" class="hint"></span></p>
 <div class="bar hide" id="r-bar"><i></i></div><pre id="r-log" class="hide"></pre>
 <div class="hint" style="margin-top:8px">Runs on your machine only. Excel runs hidden on a copy of the workbook; your own Excel and PowerPoint are not touched. When the job finishes, its result opens automatically.</div>
</section>
<section class="card"><h2>Previous runs</h2>
 <div class="row"><label>Show <select id="h-deck"><option value="">both decks</option><option value="main">Lux (main deck)</option><option value="uk">UK branch deck</option></select></label>
 <label><select id="h-kind" aria-label="Run type"><option value="">all run types</option><option value="build">Builds</option><option value="dry-run">Dry runs</option><option value="scan">Scans</option><option value="check">Workbook checks</option></select></label></div>
 <div class="scroll"><table><tr><th>When</th><th>Deck</th><th>Run type</th><th>Template</th><th>Outcome</th><th>Pictures</th><th>Errors</th><th>Warnings</th><th></th></tr><tbody id="hist"></tbody></table></div>
</section>
</div></main>
<script>
const D = /*__DATA__*/null; const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const when = iso => { const d = new Date(iso); return isNaN(d) ? iso : d.toLocaleString([], {day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'}); };
$('#theme').onclick = () => { const r = document.documentElement; const cur = r.dataset.theme || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light'); r.dataset.theme = cur === 'dark' ? 'light' : 'dark'; };
const deckNow = () => document.querySelector('input[name=deck]:checked').value;
// ---- what to update: filters (LoBs, content types), the slide picker and a live preview computed by the server
const PARTS = Object.keys(D.parts);
$('#r-parts').innerHTML = PARTS.map(p => `<label class="pill"><input type="checkbox" value="${p}" checked> ${esc(D.parts[p])}</label>`).join('');
const partsValue = () => { const on = [...document.querySelectorAll('#r-parts input:checked')].map(i => i.value); return !on.length ? null : (on.length === PARTS.length ? '' : on.join(',')); };
const spans = nums => { nums = [...nums].sort((a, b) => a - b); const out = []; for (let i = 0; i < nums.length;) { let j = i; while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++; out.push(i === j ? nums[i] : nums[i] + '-' + nums[j]); i = j + 1; } return out.join(', '); };
let pvRows = [], pvTimer = null, pvSeq = 0;
const narrowed = () => !!($('#r-lobs').value.trim() || partsValue() !== '' || $('#r-slides').value.trim());
const schedulePreview = () => { clearTimeout(pvTimer); pvTimer = setTimeout(refreshPreview, 350); };
const resetSelection = () => { $('#r-slides').value = ''; schedulePreview(); };
async function refreshPreview() {
  const seq = ++pvSeq, tpl = $('#r-tpl').value.trim(), pv = partsValue();
  $('#b-build').textContent = narrowed() ? 'Build selected' : 'Build deck';
  if (!tpl) { $('#sel-sum').textContent = ''; return; }
  if (pv === null) { $('#sel-sum').innerHTML = '<span class="err-text">Tick at least one content type.</span>'; return; }
  $('#sel-sum').textContent = 'checking the template…';
  const q = new URLSearchParams({deck: deckNow(), template: tpl, lobs: $('#r-lobs').value.trim(), parts: pv, slides: $('#r-slides').value.trim()});
  let j; try { j = await (await fetch('/api/preview?' + q, {cache: 'no-store'})).json(); } catch (e) { j = {error: 'preview not available'}; }
  if (seq !== pvSeq) return;  // a newer request is on its way
  if (j.error) { pvRows = []; $('#sel-sum').innerHTML = `<span class="err-text">${esc(j.error)}</span>`; renderPicker(); return; }
  pvRows = j.rows;
  const T = j.template || {}, newer = (T.newer || [])[0];
  $('#tpl-info').innerHTML = T.name ? `Template in use: <b>${esc(T.name)}</b>, saved ${esc(when(T.modified))}.` + (newer ? ` <span class="err-text">A newer deck exists: ${esc(newer.name)} (saved ${esc(when(newer.modified))}) in ${esc(newer.path.replace(/[\\/][^\\/]*$/, ''))}.</span> <button class="btn sec" type="button" id="b-usenew">Use it</button>` : '') : '';
  if (newer) $('#b-usenew').onclick = () => { $('#r-tpl').value = newer.path; resetSelection(); };
  const unk = j.unknown_lobs.length ? ` <span class="err-text">No slide for LoB ${esc(j.unknown_lobs.join(', '))} in this template.</span>` : '';
  const adds = (j.will_add || []).length ? ` <b>Adds ${j.will_add.length} slide(s) to the deck:</b> ${j.will_add.map(a => esc(a.title)).join('; ')}.` : '';
  $('#sel-sum').innerHTML = (j.n_slides || (j.will_add || []).length ? `${narrowed() ? '<b>Partial run:</b>' : 'Whole deck:'} ${j.n_slides} slide(s), ${j.n_pictures} picture(s), ${j.n_tables} table(s) filled from the netting / IBNR files. ${j.needs_excel ? 'Opens the workbook in Excel.' : 'No Excel needed, so it takes seconds.'}` : '<span class="err-text">Nothing is selected.</span>') + adds + unk;
  renderPicker();
}
function renderPicker() {
  $('#p-tbl').innerHTML = '<tr><th></th><th>Slide</th><th>LoB</th><th>Title</th><th>Updates</th></tr>' + pvRows.map(r => `<tr><td><input type="checkbox" data-s="${r.slide}" ${r.selected ? 'checked' : ''} aria-label="Slide ${r.slide}"></td><td>${r.slide}</td><td>${esc(r.lob)}</td><td>${esc(r.title)}</td><td class="hint">${esc([r.pictures.length ? r.pictures.length + ' picture(s)' : '', r.tables.length ? r.tables.length + ' table(s)' : ''].filter(Boolean).join(' + '))}</td></tr>`).join('');
  $('#p-tbl').querySelectorAll('input').forEach(cb => cb.onchange = () => {
    const all = [...$('#p-tbl').querySelectorAll('input')], on = all.filter(c => c.checked).map(c => +c.dataset.s);
    if (on.length) $('#r-slides').value = on.length === all.length ? '' : spans(on);
    schedulePreview();  // with nothing ticked the preview ticks the rows again: an empty selection is not a run
  });
}
$('#b-pick').onclick = () => $('#picker').classList.toggle('hide');
$('#p-all').onclick = resetSelection;
$('#r-lobs').addEventListener('input', resetSelection);
$('#r-parts').addEventListener('change', resetSelection);
$('#r-slides').addEventListener('input', schedulePreview);
let defaults = {}, runs = [], watching = false, timer = null;
const looksUK = p => /(^|[^A-Za-z])UK([^A-Za-z]|$)/i.test((p || '').split(/[\\/]/).pop().replace(/\.pptx$/i, ''));
function hint() { const t = $('#r-tpl').value.trim(), uk = looksUK(t), d = deckNow();
  const msg = t && d === 'main' && uk ? 'This template file looks like the UK deck, but "Lux (main deck)" is selected.' : (t && d === 'uk' && !uk ? 'This template file does not look like the UK deck, but "UK branch deck" is selected.' : '');
  $('#tpl-hint').textContent = msg; $('#tpl-hint').classList.toggle('hide', !msg); }
const save = () => { try { localStorage.setItem('resco.form', JSON.stringify({deck: deckNow()})); } catch (e) {} };
function renderHist() {
  const d = $('#h-deck').value, k = $('#h-kind').value;
  const rows = runs.filter(r => (!d || r.deck === d) && (!k || r.kind === k));
  $('#hist').innerHTML = rows.map(r => `<tr><td>${esc(when(r.created))}</td><td><span class="badge d-${esc(r.deck)}">${esc(r.deck_label)}</span></td><td>${esc(r.kind_label)}</td><td class="hint">${esc(r.template)}</td>
   <td><span class="badge b-${esc(r.status)}">${esc(r.label)}</span>${r.partial ? `<div class="hint">Partial: ${esc(r.partial)}</div>` : ''}</td><td>${r.kind === 'build' ? r.replaced + ' / ' + r.slots + (r.tables ? ' + ' + r.tables_filled + ' / ' + r.tables + ' tables' : '') : (r.kind === 'dry-run' ? r.slots + ' planned' : (r.kind === 'check' ? (r.slots - r.failed) + ' / ' + r.slots + ' ready' : r.slots + ' found'))}</td><td>${r.errors}</td><td>${r.warnings}</td>
   <td>${r.has_dashboard ? `<a href="/${encodeURIComponent(r.name)}/dashboard.html">Open</a>` : ''}</td></tr>`).join('') || '<tr><td colspan="9" class="hint">No runs yet: start with a scan of the template.</td></tr>';
}
async function loadRuns() { try { runs = (await (await fetch('/api/runs', {cache: 'no-store'})).json()).runs || []; } catch (e) { runs = []; } renderHist(); }
['h-deck', 'h-kind'].forEach(id => $('#' + id).addEventListener('input', renderHist));
(async () => {
  const st = await (await fetch('/api/status', {cache: 'no-store'})).json(); defaults = st.defaults || {};
  let saved = {}; try { saved = JSON.parse(localStorage.getItem('resco.form') || '{}'); } catch (e) {}
  const d0 = saved.deck || 'main'; document.querySelector(`input[name=deck][value=${d0}]`).checked = true;
  $('#r-lobs').value = ''; $('#r-wb').value = defaults.workbook || '';  // file paths are not remembered: the newest file found is always offered
  $('#r-net').value = defaults.netting || ''; $('#r-netq1').value = (d0 === 'uk' ? defaults.netting_prior_uk : '') || ''; $('#r-ibnr').value = defaults.ibnr || '';
  $('#r-tpl').value = (d0 === 'uk' ? defaults.uk_template : defaults.template) || ''; hint(); schedulePreview();
  document.querySelectorAll('input[name=deck]').forEach(r => r.addEventListener('change', () => { $('#r-tpl').value = (deckNow() === 'uk' ? defaults.uk_template : defaults.template) || $('#r-tpl').value; $('#r-netq1').value = deckNow() === 'uk' ? (defaults.netting_prior_uk || '') : ''; hint(); save(); resetSelection(); }));
  ['r-lobs', 'r-wb', 'r-tpl', 'r-net', 'r-netq1', 'r-ibnr'].forEach(id => $('#' + id).addEventListener('input', () => { hint(); save(); }));
  $('#r-tpl').addEventListener('input', resetSelection);
  const post = mode => async () => { $('#r-msg').className = 'hint'; $('#r-msg').textContent = 'starting…'; save();
    const pv = partsValue(); if (pv === null) { $('#r-msg').className = 'err-text'; $('#r-msg').textContent = 'Tick at least one content type.'; return; }
    const body = {mode, deck: deckNow(), lobs: $('#r-lobs').value.trim(), parts: pv, slides: $('#r-slides').value.trim(), workbook: $('#r-wb').value.trim(), template: $('#r-tpl').value.trim(), netting: $('#r-net').value.trim(), netting_prior: $('#r-netq1').value.trim(), ibnr: $('#r-ibnr').value.trim()};
    const r = await fetch('/api/run', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)}); const j = await r.json();
    if (j.error) { $('#r-msg').className = 'err-text'; $('#r-msg').textContent = j.error; return; } watching = true; poll(); };
  $('#b-scan').onclick = post('scan'); $('#b-check').onclick = post('check'); $('#b-dry').onclick = post('dry-run'); $('#b-build').onclick = post('build');
  if (st.state === 'running') { watching = true; poll(); }
  loadRuns();
})();
async function poll() {
  clearTimeout(timer); const s = await (await fetch('/api/status', {cache: 'no-store'})).json();
  const busy = s.state === 'running'; ['b-scan', 'b-check', 'b-dry', 'b-build'].forEach(id => $('#' + id).disabled = busy);
  $('#r-log').classList.toggle('hide', !s.log.length); $('#r-log').textContent = s.log.slice(-40).join('\n'); $('#r-log').scrollTop = 1e9;
  $('#r-bar').classList.toggle('hide', !busy); $('#r-bar i').style.width = (s.progress || 0) + '%';
  const name = {scan: 'template scan', check: 'workbook check', 'dry-run': 'dry run', build: 'build'}[s.mode] || s.mode, dk = D.deck_label[s.deck] || '';
  if (busy) { $('#r-msg').className = 'hint'; $('#r-msg').textContent = `running ${name} (${dk})…`; timer = setTimeout(poll, 1500); return; }
  if (s.state !== 'done') return;
  const ok = s.exit_code === 0 || s.exit_code === 3 || (s.mode === 'check' && s.exit_code === 1 && !!s.run_dir);  // a check that found problems still has a result page
  $('#r-msg').className = ok ? 'hint' : 'err-text';
  $('#r-msg').textContent = s.exit_code === 0 ? `${name} (${dk}) finished.` : s.exit_code === 3 ? `${name} (${dk}) finished: deck built, but the workbook data checks reported errors.` : (s.mode === 'check' && s.run_dir) ? `${name} (${dk}) found problems in the workbook.` : `${name} (${dk}) failed (exit ${s.exit_code}): see the log below.`;
  loadRuns();
  if (ok && s.run_dir && watching) { $('#r-msg').textContent += ' Opening the result…'; setTimeout(() => { location.href = '/' + encodeURIComponent(s.run_dir) + '/dashboard.html'; }, 1200); }
}
</script></body></html>
"""
