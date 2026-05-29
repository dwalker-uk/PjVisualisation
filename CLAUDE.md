# CLAUDE.md

Guidance for working in this repository.

## What this project is

Standalone, single-file HTML tools that load an Excel workbook **in the
browser** and render programme/project visualisations. The workbook is parsed
locally with [SheetJS](https://sheetjs.com/) — the only external dependency,
loaded from a CDN at runtime. Nothing is uploaded; there is no build step, no
server, and no package manager.

## Files that matter

- **`roadmap.html`** — the main, current tool and the file you should almost
  always be editing: an interactive LOE → Theme → Activity timeline with
  milestones, benefits and risks. Self-contained (HTML + CSS + JS in one file).
- `index.html` — a **previous/experimental** version. Not the current tool;
  leave it alone unless explicitly asked. Treat `roadmap.html` as the source of
  truth.
- `sp-test.html` — throwaway harness that proved a SharePoint-hosted `.xlsx`
  can be fetched and parsed in the browser. Safe to ignore/delete.
- `test/data-layer.test.js` — Node unit tests for the roadmap's pure data layer.
- `test/make-sample-xlsx.py` — regenerates `test/sample-roadmap.xlsx`.
- `README.md` — human-facing summary and a "where to change key parameters"
  guide (colours, sheet/column names, layout). Keep it in sync with real
  changes.

## Architecture of `roadmap.html`

The single file has two deliberately separated halves:

1. **Pure data layer** — between the `PURE-DATA-LAYER START` / `END` comment
   markers. No DOM or SheetJS dependencies; it turns arrays-of-arrays (what
   SheetJS yields per sheet) into the in-memory model via `assembleModel()`.
   This section is what the unit tests extract and run under Node, so **keep it
   DOM-free** and keep the markers intact.
2. **Browser layer** — an IIFE handling workbook loading, filter state,
   layout, SVG rendering and UI. It early-returns under Node (`if (typeof
   document === 'undefined') return;`).

Data model shape, expected sheets/columns, and the Lookups-driven colour system
are documented in the header comment of `roadmap.html` and in `README.md`.

### Key places to change things

- **Colours**: CSS `:root { … }` vars (UI chrome); `BRAND_COLOURS` /
  `DEFAULT_PALETTE` / `NEUTRAL_COLOUR` (data palette); the **Lookups** sheet +
  `colourFor(...)` calls (which columns are coloured).
- **Sheet/column names**: `findSheet(wb, [...])` aliases and the `ALIASES`
  object.
- **Layout & sizing**: the `LAYOUT` object near the top of the browser layer.

## Testing & verifying

Run the data-layer tests (no browser needed):

```bash
node test/data-layer.test.js
```

When touching the data layer, add/adjust a check in `test/data-layer.test.js`.
The browser layer has no automated tests; sanity-check its syntax with
`node --check` on the extracted script, and where possible verify behaviour by
opening `roadmap.html` in a browser and loading `test/sample-roadmap.xlsx`.

## Conventions

- Match the existing ES5-style, well-commented, framework-free code. No new
  dependencies, no build tooling, no bundlers — everything stays in the one
  file.
- Preserve the `PURE-DATA-LAYER` markers and the data/browser split.
- Update `README.md` (and this file) when behaviour or key parameters change.
