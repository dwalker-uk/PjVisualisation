# PjVisualisation

Standalone, single-file HTML tools that load an Excel workbook in the browser
and render programme/project visualisations. Nothing is uploaded — the workbook
is read locally with [SheetJS](https://sheetjs.com/) (the only external
dependency, pulled from a CDN at runtime).

## Files

| File | What it is |
|------|------------|
| **`roadmap.html`** | The main tool: an interactive programme **roadmap** (LOE → Theme → Activity timeline with milestones, benefits and risks). This is the file you almost certainly want. |
| `index.html` | An earlier/alternative roadmap visualiser. Kept for reference. |
| `sp-test.html` | Throwaway harness proving a SharePoint-hosted `.xlsx` can be fetched and parsed in the browser. Safe to delete. |
| `test/` | Node unit tests for the roadmap's pure data layer, plus a script to (re)generate the sample workbook. |

To use a tool, just open the `.html` file in a browser and choose an Excel file.

---

## `roadmap.html` at a glance

The file has two clearly separated halves:

1. **Pure data layer** — between the `PURE-DATA-LAYER START` / `END` comment
   markers. No DOM or SheetJS dependencies; it turns plain arrays-of-arrays
   (what SheetJS produces from each sheet) into the in-memory model via
   `assembleModel()`. This is the part exercised by the unit tests.
2. **Browser layer** — workbook loading, filter state, vertical/horizontal
   layout, SVG rendering and all UI interaction.

### Expected workbook shape

Read by `buildSheetsFromWorkbook()`. Sheet names are matched case-insensitively,
and each accepts a few aliases (see `findSheet(...)` calls):

| Sheet | Aliases accepted | Columns (matched by header text) |
|-------|------------------|----------------------------------|
| **LOEs & Themes** | `loes and themes`, `loes&themes` (or separate `LOEs` + `Themes` sheets) | An LOEs table (`LOE ID`, `Title`) and a Themes table (`LOE ID`, `Theme ID`, `Title`, `Description`) stacked in column A |
| **Activities** | `activity` | `Theme ID`, `Activity ID`, `UniqueAct ID`, `Level`, `Title`, `Description`, `Start Date`, `End Date`, `Resourcing` |
| **Milestones** | `milestone` | `Activity ID`, `UniqueMSt ID`, `Milestone Title`, `Description`, `Date`, `Owner`, `Delivery Confidence` |
| **Benefits** | `benefit` | `Milestone ID`, `Benefit Title`, `Category`, `Beneficiary`, `Impact` |
| **Risks** | `risk` | `Milestone ID`, `Risk Title`, `Risk Owner`, `RAG` |
| **Lookups** | `lookup` | A table anchored at cell **F1** that drives list-value colours (see below) |

The Activity hierarchy is encoded in the dash-delimited `Activity ID`
(`Top` → `Top-First` → `Top-First-ABC`); `UniqueAct ID` is the primary key.
Tables are read downward from their header row until the first fully blank row.

---

## Where to change key parameters

Everything you're likely to want to tune lives in a small number of named spots,
all inside `roadmap.html`.

### Colour scheme

- **UI chrome / theme** (page background, panels, text, borders): the CSS
  custom properties in the `:root { … }` block near the top of the file
  (`--bg`, `--bg-2`, `--fg`, `--secondary`, the accent colours, etc.).
- **Brand palette used by data colours**: the `BRAND_COLOURS` map (named colours
  → hex) and `DEFAULT_PALETTE` (the fallback colour cycle) near the top of the
  data layer. `NEUTRAL_COLOUR` is the fallback for unmatched values.
- **Which columns are coloured, and their defaults**: colours come from the
  **Lookups** sheet via `colourFor(...)`. The columns currently coloured are:
  - `Activities` → `Resourcing` (activity bars)
  - `Milestones` → `Delivery Confidence` (milestone markers)
  - `Risks` → `RAG` (risk icons)

  Each call passes a hard-coded fallback hex used when the Lookups sheet has no
  matching value.

### The Lookups sheet (data-driven colours)

Parsed by `parseLookups()`. The table starts at **column F** (index 5); each
column rightward is one list:

- row 1 = Table/Sheet name (e.g. `Risks`)
- row 2 = Column name (e.g. `RAG`)
- row 3 = optional `;`-separated colour names (e.g. `red;amber;green`)
- rows 4+ = the allowed values, in order, until the first blank cell

Colour names resolve through `BRAND_COLOURS`; anything not in the map falls
through to a literal CSS colour (so `grey`, `white`, etc. still work). If no
colours are given, `DEFAULT_PALETTE` is cycled.

### Column & sheet names

- **Sheet names / aliases**: the `findSheet(wb, [...])` calls in
  `buildSheetsFromWorkbook()`.
- **Column header aliases**: the `ALIASES` object in the data layer. Each field
  lists the header strings it will accept (first match wins, case-insensitive),
  e.g. `start: ['start date', 'start']`. Add an alias here to support a
  differently-named column.

### Layout & sizing

The `LAYOUT` object (top of the browser layer) holds the geometry constants:

- `labelWidth` — width of the activity-label column
- `rowHeight`, `barHeight`, `themeHeaderHeight` — row/bar sizing
- `markerSize` — milestone/risk glyph size
- `labelFont` — activity title font size
- `minMonthWidth` — narrowest a month column gets before the chart scrolls
- `minScale` / `maxScale` — bounds on the auto-fit vertical scaling
- `mstLabelMaxWidth`, `mstLabelMaxLines`, `mstLabelLineHeight` — milestone-label
  wrapping. Label vertical space is **adaptive**: it's reserved per activity,
  only for activities that have milestones, and only as tall as each label is
  *estimated* to wrap (see `mstLabelZoneFor()` / `estimateLabelLines()`).

### Default date window

In `loadWorkbook()`: defaults to the current month through +12 months, clamped
to the data's date extent. The "reset" link in the config bar uses the same
rule.

---

## Tests

The pure data layer is unit-tested under Node (no browser needed):

```bash
node test/data-layer.test.js
```

The test harness slices the code between the `PURE-DATA-LAYER` markers out of
`roadmap.html` and runs it in an isolated VM, then feeds it hand-built sheet
data. `test/make-sample-xlsx.py` regenerates `test/sample-roadmap.xlsx`, a small
workbook you can load into the tool by hand.
