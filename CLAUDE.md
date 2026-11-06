# ResCo Deck Builder

## Purpose

Automate the quarterly population of the Reserving Committee (ResCo) slide deck from the ABE/MBE reserving workbook.

Today the process is: a VBA macro in a separate Excel file loops over the lines of business (LoBs), screenshots ranges of the `ResCo tables_new` sheet, and pastes them as images into a scratch presentation. The images are then copied by hand into the real deck. This project replaces that with one Python command that writes directly into a copy of last quarter's deck and produces a validation report.

There are two outputs:
1. **Main deck**, built from `ResCo tables_new` and `ExpertJudgement`.
2. **UK branch deck** (separate presentation), built from `ResCo tables UK New`. Build this second, once the main deck works.

## Working rules for Claude

- **Confidential data.** The workbook contains company reserving figures. Develop and test against files in `data/sample/` unless I explicitly say otherwise. Do not print large blocks of real figures into the conversation; summarise (shapes, totals, pass/fail).
- **Never modify inputs.** Never save over the workbook or the template deck. Always write to `output/`. Open the workbook read-only where possible, and never call `Save` on it.
- **Facts vs assumptions.** Items marked **[VERIFIED]** below were confirmed by inspecting the files. Items marked **[TO VERIFY]** are hypotheses: check them (against the VBA macro, the workbook, or by asking me) before relying on them. Do not silently fill gaps; ask.
- **Plan first.** For any non-trivial step, propose the plan and wait for my OK before writing large amounts of code.
- **Small steps.** Make each step runnable and checkable on its own (see "Build plan").
- **Nothing hard-coded per quarter.** No quarter, date, or file name in the code. Use CLI arguments and `config/*.yaml`.

## Inputs

| Input | Example (Q3 2026) | Notes |
|---|---|---|
| Workbook | `ABE_MBE_202609_v2_12.xlsm` | ~35 MB, macro-enabled, contains `xl/vbaProject.bin` and @RISK settings |
| Template deck | `ReservingPresentation_202606.pptx` | Previous quarter's final deck, 103 slides, 16:9 |
| Legacy VBA | the separate screenshot workbook | The source of truth for which ranges are copied today. I will add it to `legacy/`. |

## Workbook facts

### ResCo tables_new (main deck source)
- **[VERIFIED]** The sheet shows **one LoB at a time**. Cell `B1` holds the name of a summary sheet (e.g. `Summary_DBENR`), and the formulas pull from it via `INDIRECT($B$1 & "!...")`.
- **[VERIFIED]** The saved file only contains calculated values for the LoB currently in `B1`. **Looping requires Excel to recalculate** after each change of `B1`. Reading with openpyxl alone will NOT give per-LoB values from this sheet.
- **[VERIFIED]** Valid selector values are the `Summary_<CODE>` sheet names. Existing sheets: `Summary_ALL`, `Summary_Live`, `Summary_PRTOT`, `Summary_DBTOT`, `Summary_ENTOT`, `Summary_LITOT`, `Summary_DCTOT`, `Summary_ENUPS`, `Summary_ENDWM`, `Summary_ENCON`, `Summary_DBENR`, `Summary_DBLIF`, `Summary_DBPRO`, `Summary_DBTRT`, `Summary_Life_health`, `Summary_LIIND`, `Summary_LISHT`, `Summary_LIMED`, `Summary_LIMLT`, `Summary_DCFRO`, `Summary_DCMAR`, `Summary_DCBON`, `Summary_DCFIL`, `Summary_DCPRO`, `Summary_DCENR`, `Summary_DCSPE`, `Summary_DCLIA`, `Summary_TRADX`.
- **[VERIFIED]** Header dates come from named ranges `asof` (`'Landing Page'!$D$3`) and `lastq` (`'Landing Page'!$D$5`).
- **[VERIFIED]** Several cells in the ResCo ranges evaluate to `#DIV/0!`, `#N/A`, or `#VALUE!` (e.g. right-hand blocks from column AR, future development periods). The tool must handle these explicitly (see "Display rules").
- **[TO VERIFY]** Approximate block layout, inferred from cell contents. **Confirm exact ranges against the legacy VBA before use:**
  - Results table, ATTRITIONAL + LARGE: rows ~4–17, columns C:AA (with comparison columns AC:AK)
  - Results table, ATTRITIONAL: rows ~19–30
  - Profit (loss) movement QoQ / YTD side tables: around Y24:AA32
  - Triangles (cumulative by UWY × development quarter 3…>=60):
    - `GPrmB` rows ~36–48
    - `GClmP` rows ~50–62
    - `GClmI` rows ~64–76
    - `ClmI/GPrmB` rows ~78–90
    - Percentage versions in columns AR onwards
  - Patterns-like blocks (`GPrmB`, `GClmP`, `GClmI` as percentages): rows ~97–137
  - Summary % table: T142:W151
- **[TO VERIFY]** Which `Summary_*` value corresponds to the deck's "Total" pages (probably `Summary_ALL`).

### ExpertJudgement (EJ slides)
- **[VERIFIED]** The sheet is hidden. Its used range is A1:AL185, with the header at row 4. Each LoB occupies a 10-row block, and column E holds the LoB code: ENUPS 5–14, ENDWM 15–24, ENCON 25–34, ENOTH 35–44, PRTOT 45–54, LIIND 55–64, LIMED 65–74, LIMLT 75–84, LISHT 85–94, DCMAR 95–104, DCENR 105–114, DCPRO 115–124, DCBON 125–134, DCFIL 135–144, DCSPE 145–154, DCLIA 155–164, DCFRO 165–174, DBENR 175–184.
- **[TO VERIFY]** Which columns appear on the slide, and whether the slide shows the header plus the LoB block. Check against the legacy VBA and the Q2 deck.

### ResCo tables UK New (UK deck source)
- **[VERIFIED]** It uses the same selector mechanism in `B1` (e.g. `Summary_ENOFF_UK`), but its **column layout differs** from the main sheet: UWY labels start in column B, not D. It needs its own range config.
- **[VERIFIED]** In the saved file, `B1 = Summary_ENOFF_UK`. That sheet does not exist (the UK summary sheets are `Summary_ALL_UK`, `Summary_ENTOT_UK`, `Summary_PRTOT_UK`, `Summary_DCTOT_UK`, `Summary_ENUPS_UK`, `Summary_ENDWM_UK`, `Summary_ENCON_UK`, `Summary_DCMAR_UK`, `Summary_DCFIL_UK`), so the sheet currently shows `#REF!`. Do not treat that as a bug in our code.
- **[TO VERIFY]** The UK LoB list and the UK template deck. I will provide the UK template.

## Template deck facts

- **[VERIFIED]** LoB slides are identified by title (shape named `Shape 158`) in the form `<Name> (<CODE>) - <Type>`, where Type ∈ {`Results`, `Triangles`, `Patterns`, `Expert Judgement`}. The separator is sometimes a hyphen `-` and sometimes an en dash `–`. **Match both.**
- **[VERIFIED]** Images per slide type are typically: Results 2, Triangles 4, Patterns 2, Expert Judgement 1. There are exceptions: some Results slides have 1 image (DCMAR, DCFIL, DBENR, DBLIF), and some section-header Results slides carry extra shapes.
- **[VERIFIED]** Every slide has a logo picture named `Picture 4`. **It must never be replaced.** Identify content images by position and size, or by a config mapping, not just by "is a picture".
- **[VERIFIED]** One slide (81, LIMED) has a title that says "Results" but contains 4 images in a Triangles layout. This is likely a typo in the template. Flag such inconsistencies in the report rather than guessing.
- **[VERIFIED]** Some slides earlier in the deck (e.g. 16–25 "Energy (ENTOT) - Results", "Total Results") reuse the same images as the appendix. They are also in scope.
- **[VERIFIED]** Not every LoB has every slide type (e.g. TOT-level LoBs have no Patterns/EJ slides). The mapping must be driven by what exists in the template, and the report must list any expected slide that wasn't found.

## Scope (confirmed with me)

- Results, Triangles, **Patterns**, and Expert Judgement slides for every LoB present in the template.
- Total pages (TOT LoBs and Total), via the same selector mechanism.
- UK branch deck as a second output, built later.
- Narrative slides (exec summary, large losses, etc.) remain manual and **must not be touched**.

## Architecture

- **Engine:** Python driving the installed Excel via `xlwings` (set `B1` → `Calculate` → read ranges as values / export ranges as images). The reserving logic stays in the workbook.
- **Deck writing:** `python-pptx` operating on a copy of the template.
- **Phase 1 output:** replace existing images in place, keeping each image's position and size (use `Range.to_png()` or the equivalent in xlwings; confirm the API in the installed version).
- **Phase 2 output (optional, decide later):** native PowerPoint tables for selected slides.
- **Open risk:** recalculation behaviour with @RISK loaded is unknown. Test early. If recalc is slow, recalculate once per LoB and read all blocks in one pass.

### Suggested layout
```
resco/
  CLAUDE.md
  config/
    lobs.yaml          # LoB code -> selector value, display name, which slide types exist
    ranges_main.yaml   # block name -> sheet + range (from legacy VBA)
    ranges_uk.yaml
    slides.yaml        # (lob, slide type) -> ordered list of image slots
  src/resco/
    excel_reader.py    # xlwings: open, set selector, recalc, read/export blocks
    deck_writer.py     # python-pptx: locate slide by title, replace image slots
    validate.py        # checks + report
    cli.py
  legacy/              # VBA module(s) exported as .bas
  data/sample/         # anonymised/dummy copies for development
  output/
  tests/
```

### CLI target
```
python -m resco build --workbook <path.xlsm> --template <last_quarter.pptx> --out output/ [--deck main|uk] [--lobs ENUPS,ENDWM]
```
Outputs: the new `.pptx`, `run_report.md`, and a per-LoB CSV dump of every block (audit trail).

## Display rules

- Excel error values (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`) → **[TO VERIFY with me]** blank or "–". Record every occurrence in the report.
- Number formats and units should follow the workbook's formatting as rendered today. In image mode, Excel's rendering already does this.

## Validation (must pass before a deck is considered done)

1. Every LoB in `lobs.yaml` was selected, recalculated, and read without an exception.
2. After each recalculation, the LoB code shown in the sheet (`C5` in the main sheet) equals the expected code. This guards against a failed recalc producing a stale copy.
3. Totals in each block equal the row/column sums from the source data (tolerance 1e-6 relative).
4. Every expected image slot in `slides.yaml` was found and replaced. List missing or extra slots.
5. Image count and positions on every touched slide are unchanged versus the template.
6. Untouched slides are byte-identical in content (spot check: text and shape count).
7. The report lists all error-value cells and any template inconsistencies.

## Build plan (do in order, checkpoint with me after each)

1. Export the legacy VBA module(s) to `legacy/` (e.g. with `oletools`/`olevba`) and summarise exactly which ranges and LoBs it processes. Turn that into `ranges_main.yaml` and `lobs.yaml`.
2. Create `excel_reader.py`: loop over LoBs and dump every block to CSV. No PowerPoint yet.
3. Create `slides.yaml` by scanning the template and mapping `(lob, type)` → image slots. Print the mapping for my review.
4. Replace images for **one LoB** and compare against the Q2 deck by eye.
5. Run for all LoBs, run validation, and produce the report.
6. UK deck: `ranges_uk.yaml`, then repeat steps 2–5.
7. (Optional) Native tables for selected slide types.

## Conventions

- Python 3.11+, type hints, `pathlib`, `logging` (no prints in library code), `pytest` for tests.
- Pin dependency versions in `requirements.txt` once a working set is confirmed.
- Keep functions small and pure where possible; isolate all xlwings calls in `excel_reader.py` so the rest can be tested without Excel.

## Status (updated 2026-10-20)

Build plan steps 1-6 are implemented (step 7, native tables, is not). Operating guide: `README.md`; start the dashboard with
`Start dashboard.cmd`. Main deck: `config/`; UK deck: `config/uk/` (ranges inferred, to be confirmed). Findings that deviate from
the notes above: the slot geometry lives in `config/slot_kinds.yaml`; pattern charts are found by title, not by the names
`Premiums`/`Incurred`; the legacy `PrintPNGs` flags are informational (the deck decides what is produced); range pictures fall
back to an Excel -> PDF route when `CopyPicture` is unavailable (`--picture-mode`).

Since 2026-10-21 the build can also fill native tables on narrative slides from the netting workbook (`--netting`) and the IBNR allocation workbook (`--ibnr`): see README.md and `config/tables.yaml`.

Since 2026-10-22: `--parts` / `--slides` (with `--lobs`) update only part of the deck (`src/resco/selection.py`, dashboard "What to update" with live preview); DBTOT is in `config/lobs.yaml`; input workbooks live in `Inputs/`.

Since 2026-10-22 (later): a LoB slide with a title but no picture is filled by inserting the pictures at the standard positions (`SlotEntry.new`, `Job.new`; verification counts `pictures_inserted`); DBTOT, DBPRO and DBTRT are in `config/lobs.yaml`; `gvn_DBTOT` in `config/tables.yaml`; scan and run pages list configured LoBs with no slide.

Since 2026-10-23: `config/slides_add.yaml` + `src/resco/slide_adder.py`: the build adds missing slides (Dubai Property / Dubai Treaty Results + Gross vs Net) to a working copy of the template; templates are read from a copy (`copy_for_reading`); dashboard defaults pick the newest file by mtime.

Since 2026-11-06: the tool is its own git repository (outside the quarterly SharePoint folders; `local.cmd` sets `RESCO_PACK_DIR`, see `serve.pack_dir`). New table target kind `profit_qoq` (`src/resco/profit.py`, `pq_*` in `config/uk/tables.yaml`): the Profit (loss) movement QoQ boxes under the UK Gross vs Net tables, from the netting file (Q) and `--netting-prior` (Q-1, UK: as presented). Inserted as native tables on the first build (`verify.compare(..., inserted_tables)`), updated in place afterwards. Part of the `netting` content type.
