# PowerPoint export (VBA proof-of-concept)

An enterprise-friendly companion to `roadmap.html` that stays inside the
Microsoft 365 ecosystem. Instead of a standalone web page, an **Excel VBA
macro** reads the same data and generates a **PowerPoint slide of native,
editable shapes** — bars, milestone diamonds, benefit stars and risk
warning-triangles, coloured from the Lookups sheet, positioned on a
month-proportional timeline.

This is a deliberate **vertical slice** to prove fidelity before building the
whole thing: by default it renders **Level-1 activities only**. It mirrors the
pure-data-layer logic of `roadmap.html` (dash-delimited Activity-ID hierarchy,
the Lookups colour system, and the date→x mapping), so the full data model and
deeper levels can be layered on once the approach is signed off.

> Tested by inspection. VBA can only run on Windows desktop Office, so verify
> the visual output there.

## Why this approach

For the stated goal — auto-generated, high-quality **static** roadmap slides,
staying in M365, on Windows desktop Office with macros allowed — VBA driving
PowerPoint gives the highest fidelity to `roadmap.html` with no new
infrastructure, and produces a real editable deck. (Office Scripts can't create
PowerPoint; an Office Add-in is the more future-proof but heavier alternative.)

## Requirements

- **Windows** desktop Excel **and** PowerPoint (M365). VBA cross-app automation
  does not work on the web or, reliably, on Mac.
- Macros allowed to run (a Trusted Location or signed macro is fine).
- No references to add — PowerPoint and `Scripting.Dictionary` are bound late
  via `CreateObject`.

## Install

1. Open the data workbook in Excel.
2. `Alt`+`F11` to open the VBA editor, then add the module either way:
   - **Import (cleanest):** **File ▸ Import File…** → choose `RoadmapExport.bas`.
   - **Paste:** **Insert ▸ Module**, then paste in the file's contents.
3. (Optional) add a settings sheet — see below.
4. Save the workbook as **.xlsm** (macro-enabled) if you want to keep the macro.

> If you see a syntax error on the very first line, you pasted a copy that
> still had an `Attribute VB_Name = …` line at the top — delete that one line.
> (The committed `.bas` no longer has it, so a fresh copy pastes cleanly.)

## Run

`Alt`+`F8` → select **`GenerateRoadmap`** → **Run**. PowerPoint opens with the
generated slide on a navy (`#071D49`) background. Re-run after editing data or
settings to regenerate.

You can also wire it to a button: Developer ▸ Insert ▸ Button, assign
`GenerateRoadmap`.

## Expected data

Same as the web tool — the named tables **LOEs, Themes, Activities,
Milestones, Benefits, Risks** (matched by table name, case-insensitive,
anywhere in the workbook) and a **Lookups** worksheet with the lookup block
anchored at cell **F1**. Column headers are matched by the same aliases as
`roadmap.html`. See the root `README.md` / `roadmap.html` header for the full
schema.

## Settings sheet (optional)

Add a worksheet named **`Roadmap Settings`** (or `Settings`) with labels in
column **A** and values in column **B**:

| A (label)               | B (value)                                   | Default                  |
|-------------------------|---------------------------------------------|--------------------------|
| `Title`                 | Slide title text                            | `Programme Roadmap`      |
| `Start Date`            | A date (any day in the month)               | First of current month   |
| `End Date`              | A date                                      | Current month + 12       |
| `Max Level`             | Deepest activity level to show              | `1`                      |
| `Show Milestone Labels` | `TRUE` / `FALSE`                            | `FALSE`                  |
| `LOEs`                  | Comma-separated LOE titles (or IDs); blank = all | all                 |
| `Themes`                | Comma-separated Theme titles (or IDs); blank = all | all               |
| `Only With Benefit`     | `TRUE` / `FALSE` — hide activities with no benefit milestone | `FALSE`  |
| `Only With Risk`        | `TRUE` / `FALSE` — hide activities with no risk milestone     | `FALSE`  |

The date window is clamped to the data's own min/max, matching the web tool.
If no settings sheet exists, all defaults apply. Filters match on either the
title or the ID, case-insensitively, so `LOEs` = `Digital, Estates` works.

## What it draws

- **Title** across the top.
- **Date axis** — month abbreviations with year markers — across the top of
  the plot, plus faint month gridlines.
- **LOE gutter** — rotated (upward) LOE titles down the left edge, one per LOE
  band.
- **Theme headers** — full-width blocks introducing each Theme's activities.
- **Activity rows** — a left-hand label (indented by level) and a rounded bar
  spanning Start→End, clamped to the window, coloured by **Resourcing**.
- **Milestones** — a **diamond**, or a **star** if the milestone has a Benefit,
  coloured by **Delivery Confidence**; with a small **warning triangle**
  adjacent if it has a Risk, coloured by the Risk **RAG**.
- **Milestone labels** (when enabled) now sit in dedicated space *below* the bar,
  so they no longer hide under the activity bars.
- **Hover detail** — every bar and milestone carries a pop-up **ScreenTip**:
  activity dates/resourcing/description, and for milestones the date, delivery
  confidence, owner, plus any benefit and risk detail. ScreenTips appear on
  hover in **Slide Show** view. The same text is also stored as the shape's
  **alt text** (a reliable fallback, visible via right-click ▸ Edit Alt Text).
- **Multiple slides** — when the rows don't fit one slide (especially with
  `Max Level` > 1), the roadmap **splits across slides**. The date axis repeats
  on each slide, a Theme that spans a break is repeated with a `(cont.)` marker,
  and a `Slide x of y` footer is added.

## Known limitations of this PoC

- **Hover ScreenTips show in Slide Show view**, not while editing — that's a
  PowerPoint constraint. The detail is always available as alt text too.
- **Expand/collapse** is not interactive; instead, `Max Level` chooses the depth
  up front (and deeper levels paginate across slides).
- Milestone-label collision avoidance is simpler than the web tool's: labels are
  width-bounded by their neighbours but may still abbreviate when crowded.
- Filters are driven from the settings sheet rather than live dropdowns.
- Text wrapping/fit is PowerPoint's own, so very long labels may clip.

These are all straightforward to extend further once the approach is approved.
