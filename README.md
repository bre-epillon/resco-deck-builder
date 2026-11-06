# ResCo deck builder: how to run it each quarter

Fills the pictures on the LoB slides of last quarter's ResCo deck from the ABE/MBE workbook, checks the data, and
writes a **copy** of the deck. Narrative slides are never touched; the original deck and workbook are never modified.
`CLAUDE.md` is the specification; this file is the operating guide.

## The easy way: the dashboard

Double-click **`Start dashboard.cmd`** (restart it after updating the tool). A page opens in your browser (local only,
`http://127.0.0.1:8765`). It always opens on the **home page**, which starts runs and lists earlier ones, and never shows a
result itself:

1. Choose the deck: **Lux (main deck)** or **UK branch deck**. The template path follows the deck, and a template whose file
   name says UK is refused for the Lux deck (and the other way round). Optionally limit the LoBs (e.g. `ENUPS,ALL`).
2. **Scan template** shows which slides and picture slots the tool recognises (no Excel). **Check workbook** opens a copy of
   the workbook and verifies that everything the build needs is there (about 1.5 minutes, no pictures). **Dry run** adds the plan,
   which picture goes where (no Excel, nothing written). **Build deck** does the real run (about 7 minutes Lux, 4 minutes UK).
3. When the job finishes, **its own result opens automatically**. Every result page names the deck and the run type in the
   header, and has "All runs" (back to the home page) and a switcher between the earlier runs of the same deck.

While a build runs, Excel works hidden on a *copy* of the workbook. Your own Excel and PowerPoint sessions are not touched.

**Range pictures and the display.** Pictures of cell ranges (results table, P&L box, triangles, Expert Judgement) are made
like the legacy macro does, with Excel's `CopyPicture`. That needs a live display and a free clipboard. If it is not
available (screen off or locked, remote-desktop window minimised, another program using the clipboard) the tool switches
to a fallback by itself: Excel prints the range to PDF and the PDF is converted to an image. It works without a display,
is about 30 seconds slower per LoB and looks slightly different (sharper, print metrics). The report then shows a
warning with the number of pictures made that way. Force a mode with `--picture-mode clipboard|pdf|auto`. Chart pictures
never need the clipboard.

## The same from a terminal

```bash
set PYTHONPATH=src
C:\Users\ErikPillon\venvs\resco\Scripts\python.exe -m resco scan  --template <deck.pptx>
C:\Users\ErikPillon\venvs\resco\Scripts\python.exe -m resco build --workbook <abe_mbe.xlsm> --template <deck.pptx> --dry-run
C:\Users\ErikPillon\venvs\resco\Scripts\python.exe -m resco build --workbook <abe_mbe.xlsm> --template <deck.pptx> [--lobs ENUPS,ALL] [--config config\uk]
```

## Updating only part of the deck

By default a build replaces everything the template holds. To touch only some slides, narrow the run down; everything else in
the deck stays exactly as the template had it (verified: the other slides are XML-identical).

| Filter | Dashboard | Terminal | Meaning |
|---|---|---|---|
| LoBs | "LoBs" field | `--lobs ENTOT,DBTOT` | the codes used in the slide titles (also `ALL` for the Total slides). A LoB needs an entry in `config/lobs.yaml` and a slide in the deck |
| Content type | tick boxes | `--parts results,net` | `results` (table + P&L), `triangles` (block + 3 charts), `patterns`, `ej` (Expert Judgement), `netting` (alias `net`: the Gross vs Net tables), `ibnr` (the IBNR allocation / QoQ tables). A single picture kind also works: `profit`, `table`, `chart_prem`... |
| Slides | "Slides" field / "Pick slides..." | `--slides 16,17,55-58` | slide numbers of the template; the picker lists exactly the slides that pass the other two filters, untick the ones to leave alone |

The filters combine: a slide is touched when it passes all of them. Examples:

```bash
:: only the Gross vs Net table of ENTOT (netting file only: no Excel, takes seconds)
python -m resco build --workbook <abe_mbe.xlsm> --template <deck.pptx> --netting <netting.xlsx> --lobs ENTOT --parts net
:: only the Results slides (table + P&L) of ENTOT, both occurrences
python -m resco build ... --lobs ENTOT --parts results
:: only the Triangles slides of ENTOT and ENUPS
python -m resco build ... --lobs ENTOT,ENUPS --parts triangles
:: only the first Results slide of ENTOT (front section), not the appendix copy
python -m resco build ... --lobs ENTOT --parts results --slides 16
```

The dashboard shows a live preview ("Partial run: 3 slides, 6 pictures, 1 table ...") computed by the same code as the build, says
whether Excel is needed, and the button reads "Build selected". Partial runs are marked "Partial run" on their result page and in the
run history. Notes: Excel is only opened when a picture or an IBNR table is selected; the IBNR tables follow `--lobs` (skipped in a
LoB-limited run) unless `ibnr` is asked for; slide numbers belong to the template of that run (they move when slides are added).

Each run writes `output/build_<timestamp>/` with `<deck>_auto.pptx`, `dashboard.html`, `build_report.md`,
`build_result.json`, `images/` (every exported picture) and `compare/` (old vs new).

## A new LoB or a new slide in the template

The deck decides what is produced. To add a LoB, give it an entry in `config/lobs.yaml` (selector `Summary_<CODE>`; `resclass` for
LoBs with Patterns / Expert Judgement) and add a slide titled like the others, e.g. `Dubai (DBTOT) - Results`. **The slide can be empty**:
a LoB slide without any picture gets its pictures inserted at the standard positions of its slide type (Results: table + P&L box;
Triangles: block + three charts; Patterns: two charts; Expert Judgement: one picture; Life classes use the shorter table). The scan
and the build report warn about such slides ("the slide has no pictures ... will be inserted"), the run page counts them as
"inserted into empty slides", and the verification checks that nothing else on the slide or in the deck changed. Next quarter the
pictures are there and the slide is treated like any other. A slide with some pictures that the tool does not recognise is never
treated as empty: the unrecognised pictures are reported and left alone.

### Slides the build adds itself (`config/slides_add.yaml`)

For slides that the template does not have yet, the build can add them: each entry names a **model slide** of the template, the new
title, and the slide it goes after. Currently: Dubai Property (DBPRO) and Dubai Treaty (DBTRT), each a Results slide like Dubai
Energy (table only) followed by its "Gross vs Net summary - MBE" slide (a full copy of the DBTOT / DCTOT one, whose table is filled
from the netting file), placed after the Dubai Life Results slide. The new slide is a copy: a `frame` copy keeps only the title, logo
and slide number (its pictures are exported and inserted, so a failed export leaves it empty and never shows another LoB's
figures); a `full` copy keeps everything (the table is then filled from Excel). The template file is never changed: the slides are
added to a working copy, listed in the report ("Slides added to the deck") and the verification checks that every original slide is
still identical. Once the slides exist in the template (next quarter's deck), nothing is added. A run limited with `--lobs`,
`--parts` or `--slides` only adds the slides that belong to it (`--slides` never adds any). To add another slide, add an entry to
`slides_add.yaml` (and, for a Gross vs Net slide, a `ladder` target in `config/tables.yaml`).

The tool reads a **copy** of the template, so it is not disturbed if you save the deck while a run is going. If the template is open
in PowerPoint and locked at that moment the run stops with "cannot read ... it is probably open in PowerPoint": save and close it
(or work on a copy) and run again. The dashboard form is prefilled with the newest template / workbook it finds (by modification
time) in `Inputs/`, in the project folder and in the folder above it.

The scan and the run pages also list the configured LoBs that have **no slide** in the template, so a missing LoB is visible.
Currently configured: ALL, ENTOT, PRTOT, LITOT, DCTOT, DBTOT, ENUPS, ENDWM, ENCON, LI*, DC*, DBENR, DBLIF, DBPRO, DBTRT.
Gross vs Net tables: one `config/tables.yaml` target per slide (DBTOT: `gvn_DBTOT`, ladder level Aggregate).

## Figures from the netting and IBNR allocation workbooks (native slide tables)

Besides the ABE/MBE workbook, the build can fill the **native tables** of some narrative slides. Give the two files in the
dashboard form (or `--netting` / `--ibnr`); leave a field empty to skip that source.

| Slides (Lux deck) | Source | What is filled |
|---|---|---|
| 17, 19, 21, 23, 25 "Gross vs Net summary - MBE" | netting workbook, sheet `Ladder` | one ladder block per slide: LoB, level, basis (MBE), branch (Total) |
| 27 "Earned IBNR allocation" (two tables) | IBNR allocation workbook, sheet `Overview` | the UWY/GEP table and the Gross/Net by-class table |
| 43 "QoQ movements Earned Statutory IBNR gross and net" | IBNR allocation workbook, sheet `Overview` | the QoQ movement tables, Gross and Net |

The UK deck uses the `UK branch` rows of the ladder and the UK blocks of `Overview` (further right on the sheet).
Which slide, table and source rows is in `config/tables.yaml` (UK: `config/uk/tables.yaml`).

### Profit (loss) movement QoQ boxes (UK deck)

Under each UK "Gross vs Net summary - MBE" table the build puts the two small **Profit (loss) movement QoQ** tables that
the main deck shows (slide 17 of the 202609 deck): one for Net before and one for Net after internal QS, "MBE / A+L", in
millions, red for a loss and green for a profit. The figure is the netting workbook's `Summary_Tot` formula on the Total row:

    (premium ultimate Q - premium ultimate Q-1) - (claims ultimate Q - claims ultimate Q-1)

Q comes from the netting file (`--netting`), Q-1 from a second file in the same `Ladder` layout: **Netting Q-1** in the
dashboard, `--netting-prior` in a terminal. For the UK deck that is the Q-1 UK figures **as presented** in the last UK
deck (policies moved between UK and Lux since then), with premium and claims ultimate amounts, not only the ratios.

- The first build **inserts** the two boxes, right-aligned under each block's QoQ column, the header in the colour of the
  block's band. From then on they are part of the template: they are found by their title and only the figure is updated.
- They belong to the `netting` content type, so a netting-only run (no Excel, seconds) refreshes them with the tables.
- Without a Q-1 file the boxes are skipped with a warning; a LoB without Q-1 figures gets an empty grey box and a warning.
- The verification allows exactly these inserted tables and nothing else on those slides.
- Targets: `pq_*` in `config/uk/tables.yaml`. The main deck still has its pasted pictures; give it `profit_qoq` targets
  (and delete the pictures once) to have them built as well.

## Running it from this repository

The tool no longer has to sit inside the quarter's `06_ReservingPack` folder. Copy `local.cmd.example` to `local.cmd`
(git-ignored) and set `RESCO_PACK_DIR` to the quarter's pack folder; `Start dashboard.cmd` reads it. The dashboard then
offers the newest template decks from that folder, the newest workbook and netting file from the quarter's
`02_ABE&MBE_Proposals` and the newest IBNR file from `03_FinanceAllocation`. Without `local.cmd` the folder above the
tool is used, as before. `output/`, `Inputs/` and all Office files are git-ignored.

- Only the **text** of the existing table cells is replaced, so the slide keeps its styling and the table stays editable.
  Rows are matched by UWY label and columns by header text; anything without a match is left as it was and reported.
- `Overview` blocks are found by their titles (not by row numbers) and read **as Excel displays them** (`+0.6`, `(0.1)`, `65%`).
  A column too narrow to show a figure (`#####`) is widened in memory before reading; your workbook is never changed.
- The netting figures are formatted like the slides (millions with one decimal, ratios as whole percent). The tool also checks
  that each ladder ratio equals claims ultimate / premium ultimate.
- A **fit check** warns about figures probably too wide for their slide-table column ("widen those columns before printing").
- The `Tables from Excel` tab of a run shows every cell before and after. The workbook's own `checks` column (TRUE/FALSE
  on `Overview`) is reported when it shows FALSE.
- With `--lobs` only the ladder tables of those LoBs are filled; the IBNR tables are filled only in a run without `--lobs` (or with `--parts ibnr`).

## Changes in the workbook or the deck: what the tool follows and what it reports

Run **Check workbook** first when the workbook or the deck changed (about 1.5 minutes, no pictures). The build does the same checks at its start.

Followed automatically (the run report lists what was found where, and flags anything that moved):
- **Expert Judgement blocks** are found by LoB code, so blocks that are re-ordered, added or removed (v2.13b dropped ENOTH) are fine. A new UWY row each January is picked up and the render template grows by itself.
- **Results table, P&L box and triangles** are found by their labels (the `UWY` header, the `Total` rows, `Profit (loss) movement QoQ`, `Year 1`, the last `GULR` header), so a longer or shifted sheet is followed; helper columns next to the table are ignored.
- **Charts** are found by name and title; pattern charts by title (the topmost copy wins).
- **Deck:** slides by title (hyphen or en dash), pictures by position and size. A picture that was nudged or resized is still matched (flagged "approximately") and keeps its own position and size.

Reported clearly instead of guessed:
- A **renamed label or chart title** that the tool searches for (for example the `UWY` header): the configured address is used with a warning, or the slot fails with a clear message.
- A **new LoB** or a slide with a LoB code that is not in `config/lobs.yaml`.
- **Reordered columns inside** the Expert Judgement blocks, or a deck slide that is not a LoB slide pattern the tool knows.
- The UK rows that are hidden in memory (`hide_rows` in `config/uk/ranges_main.yaml`) are configured, not detected.

## What the checks mean

| Check | Fails when |
|---|---|
| LoB code on sheet | Excel did not recalculate for the selected LoB (a stale copy would have been pasted) |
| Totals vs components | A Results-table Total (or the `<=2019` roll-up) differs from the sum of its rows, relative tolerance 1e-6 |
| Error values visible | `#DIV/0!`, `#N/A`... sit in rows/columns that would be drawn in the picture (hidden rows are ignored) |
| Deck verification | Anything other than the intended pictures changed: position, size, text, other slides, shape count |
| Duplicate charts | Two charts share a title on a sheet and differ visibly: the topmost one is used |
| Stretch in the slot | The new image's aspect ratio is far from the old one, so the picture is stretched differently |

## How the deck is read (it can change between quarters)

Slides are found by **title** (`<Name> (<CODE>) - <Type>`, hyphen or en dash, or `Total <Type>`). Pictures are classified
by **position and size** (`config/slot_kinds.yaml`), never by slide number or shape name. A slide title can be wrong
(e.g. "Results" on a Triangles layout): the pictures decide, and a warning is raised. Anything unrecognised is listed and
left alone. Typos in a LoB code can be mapped with `aliases:` in `lobs.yaml`.

## Files

| Path | Purpose |
|---|---|
| `Inputs/` | The workbooks (ABE&MBE, IBNR allocation, netting): the dashboard form is prefilled with the newest file found here |
| `config/lobs.yaml`, `ranges_main.yaml`, `slot_kinds.yaml` | Main deck: LoBs, what to read from Excel, picture slots |
| `src/resco/selection.py` | "Update only part of the deck": LoB / content type / slide filters (CLI, build and dashboard preview share it) |
| `src/resco/profit.py` | The Profit (loss) movement QoQ boxes: figures from Q and Q-1 ladders, insert / update under a Gross vs Net table |
| `local.cmd.example` | Template for `local.cmd`: the quarter's pack folder (`RESCO_PACK_DIR`) and the Python to use |
| `config/uk/` | Same three files for the UK branch deck |
| `assets/ej_render_template.xlsx` | Formatted sheet used to render the Expert Judgement picture |
| `legacy/Module1.bas` | The original VBA, for reference |
| `src/resco/` | The code (`excel_reader.py` is the only module that talks to Excel) |
| `tests/` | `pytest` (no Excel or real deck needed) |
