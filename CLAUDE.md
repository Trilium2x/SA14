# TIS Tracker — Project Guide

> **Workbook:** `SA14.xlsm`
> **Modules:** 9 `.bas` files in `C:\Users\razzl\Desktop\SA\14\`
> **Excel Version:** Microsoft 365 / Excel 2021+
> **Current Revision:** Rev14 — clean rewrite from Rev12 baseline (Update-in-place, Status, Health, Per-milestone Deviations, WhatIf)
> **Baseline:** Rev12 in `C:\Users\razzl\Desktop\SA\12\`
> **Last reviewed:** 2026-03-25

---

## How to Read This File

This file is written for two audiences:

- **Human editors** (you): Read the top sections freely. Skip anything inside a `<!-- TECHNICAL -->` block — those are dense reference details meant for Claude, not for casual reading or editing.
- **Claude**: Read everything. The technical blocks contain critical implementation contracts you must not violate.

**Always update this file** when: a module is revised, a public function is renamed, sheet structure changes, or a new bug/gotcha is discovered. Treat it as part of every code change.

---

## Project Introduction

The **TIS Tracker** is an Excel-based operations tool used by ramp managers and specialists at semiconductor fabs (primarily Intel). Its job is to track the installation of hundreds of tools (equipment systems) through a multi-phase lifecycle — from dock date to final qualification.

Every week, the customer provides a new **TIS file** (Tool Install Schedule). The tracker loads it, compares it against the previous week, and updates the **Working Sheet** — which is the team's operational database. TIS dates are reference data. The team's committed schedule lives in the **Our Dates** columns and is never overwritten by automation.

From the Working Sheet, the tool generates:
- A **Gantt chart** showing all systems and their milestone timelines
- **HC (headcount) analyzer tables** showing staffing gaps per group
- A **management Dashboard** with KPI cards, group filters, and charts
- **Ramp Alignment reports** formatted for Intel meetings

The tool is built entirely in Excel VBA with no external dependencies. A companion web app (React/TypeScript) lives at `C:\Users\razzl\Desktop\SA\TISGANTTAPP\` and may eventually replace the Excel version.

---

## Ground Rules

These rules apply to every code change, no exceptions.

### Data Safety (Highest Priority)
- **Never overwrite user-entered data.** The committed milestone dates, NIF assignments, BOD dates, Escalated/Watched flags, and manual Status markings must survive every rebuild. Always migrate before clearing.
- **Committed dates are sacred.** The 8 committed milestone columns (Set, SL1, SL2, SQ, Conv.S, Conv.F, MRCL.S, MRCL.F) are the team's schedule. No automated process — TIS upload, WhatIf rebuild, or WorkfileBuilder — may overwrite them. Only the user types into these cells, or the one-time new-system initialization populates them on first appearance.
- **WhatIf uses a backup-restore approach.** WhatIf mode backs up Our Dates to a hidden sheet, shifts them in place for Gantt rebuild, then restores from backup when deactivated. Our Date columns always return to their committed values.

### Column Naming — No Internal Jargon on Screen
- **Committed milestone columns use clean abbreviations only.** The column headers are: `Set`, `SL1`, `SL2`, `SQ`, `Conv.S`, `Conv.F`, `MRCL.S`, `MRCL.F`. No "Our" prefix. No qualifier of any kind.
- **Why:** These ARE the dates. The team's committed schedule is the source of truth. Calling them "Our Set" implies they are a secondary copy of something more authoritative — the opposite of intent. TIS dates are the external reference feed; the team's dates are primary. The qualifier belongs on the TIS side, not on the team's side.
- **The "Our" prefix is internal developer vocabulary** (used in code constant names like `TIS_COL_OUR_SET`) and must never appear in any user-visible string — not in column headers, button labels, messages, or instructions.
- **TIS source columns** in the Working Sheet use verbose names (`Set Start`, `SL1 Signoff Finish`, etc.) which are naturally identifiable as TIS data.
- **Backup before full rebuild.** WorkfileBuilder saves the previous Working Sheet as `Old YYYY-MM-DD` before any full build (CreateWorkingSheet). Weekly TIS load (UpdateWorkingSheetFromTIS) does not create a backup — it is non-destructive.
- **Slicer caches first.** When clearing the Working Sheet, delete slicer caches BEFORE deleting ListObjects. Reversing this order corrupts the workbook.
- **Preserve external formula links.** The Working Sheet is rebuilt in-place (same sheet object) so that cross-sheet formulas like `='Working Sheet'!A15` in user sheets remain valid.

### Performance
- **No cell-by-cell loops on data.** All reads and writes to data ranges must go through `Range.Value` array bulk I/O. Cell-by-cell is 30-60x slower at 300+ rows.
- **No `For Each cell In UsedRange`.** Track specific cells in a Collection instead.
- **Create Scripting.Dictionary once**, not inside loops.
- **Wrap every public macro** with `SaveAppState` / `SetPerformanceMode` / `RestoreAppState`.
- **Cell comments are acceptable.** `AddComment` / `.Comment.Text` operations are fast with `ScreenUpdating=False`. ~1500 comment operations across 500 rows adds ~1-2 seconds — acceptable at current scale.
- **Target: 1000 rows in under 10 seconds** for any single operation (TIS load, rebuild, Gantt build).

**Bulk array I/O patterns used throughout:**
- `UpdateWorkingSheetFromTIS`: Reads both TIS and Working Sheet into arrays via `Range.Value`, compares in memory using Dictionary keyed by project key, batch writes changed cells. `Union` range for orange fill application.
- `ImportUserDataFromOldSheet`: Bulk array comparison between old and new sheets, `Union` range for orange fill on imported changes.
- `PopulateNewSystemOurDates`: Bulk read of TIS dates and Our Date columns, bulk write of initialized values, `Union` for blue border application.
- `SortWorkingSheet`: Temporary helper columns for sort keys, deleted after sort completes.

### UX / UI
- **Beautiful, not busy.** Dark navy theme (AMAT brand colors defined in TISCommon). Cards, clear section dividers, readable fonts (Segoe UI).
- **Intuitive without instructions.** Labels, descriptions, and button captions must be self-explanatory.
- **Reactive where possible.** Dashboard group filter and chart date selector update automatically via injected Worksheet_Change handler — no re-run needed.
- **Consistent visual language.** Use TISCommon `THEME_*` color constants everywhere. Never hardcode RGB values in module logic — add new constants to TISCommon instead.
- **Mode clarity.** When WhatIf mode is active, the user must know it at a glance: banner above the Gantt, button label change, sheet tab color change — all three.
- **Sort-by-color must work.** Change indicators (TIS field changes, new systems) use hardcoded `Interior.Color` fills — not conditional formatting — so Excel's "Sort by Cell Color" is available to the user.

### Code Quality
- **`Option Explicit` in every module.**
- **Every public Sub must have a full error handler** (`On Error GoTo ErrorHandler` -> `Cleanup:` -> `RestoreAppState`).
- **Do not use `On Error Resume Next` as a silent catch-all.** Scope it tightly to the one call that may fail, then immediately reset with `On Error GoTo 0` or `On Error GoTo ErrorHandler`.
- **No magic strings for cross-module contracts.** HC table titles, named ranges, and sheet names must be constants.
- **Status filter everywhere.** Every counter, HC calculation, Dashboard KPI, and DashHelper row that counts systems must filter by `Status`. Most live counts include only Active + On Hold. Cancelled, Completed, and Non IQ systems are excluded from all live counts (except the Completed KPI card which specifically counts Status="Completed"). See [KPI Card Status Filtering](#kpi-card-status-filtering) and [TECH] Status Filter — Mandatory Locations for the full compliance table.

### Maintainability
- **Update CLAUDE.md at the end of EVERY code operation.** This is mandatory, not optional. After any code change — bug fix, feature addition, refactoring, or configuration change — update this file to reflect the new state. This file is the single source of truth for the project. If CLAUDE.md doesn't match the code, CLAUDE.md is wrong and must be fixed immediately.
- **Keep color constants in TISCommon.** Do not duplicate THEME_* values in module-level Private constants.
- **Keep sheet name constants in TISCommon.** Add `TIS_SHEET_*` constants rather than using string literals in module code.
- **Keep Our Date column header constants in TISCommon.** All 8 Our Date column headers are `TIS_COL_OUR_*` constants — never string literals.

---

## Module Naming Convention

### The Rule
- **Filenames** keep the `_RevNN` suffix for history and archiving (e.g., `WorkfileBuilder_Rev14.bas`)
- **Internal VB_Name** uses a **stable name with no version suffix** (e.g., `WorkfileBuilder`)
- The Launcher and all cross-module calls always use the stable name — they never need to change when a revision is bumped

### Stable Name Table

| Filename Convention | Stable VB_Name (internal) |
|---|---|
| `TIS_Launcher_Rev14.bas` | `TIS_Launcher` |
| `TISCommon.bas` | `TISCommon` (no version — shared library) |
| `TISLoader_Rev14.bas` | `TISLoader` |
| `WorkfileBuilder_Rev14.bas` | `WorkfileBuilder` |
| `GanttBuilder_Rev14.bas` | `GanttBuilder` |
| `NIF_Builder_Rev14.bas` | `NIF_Builder` |
| `DashboardBuilder_Rev14.bas` | `DashboardBuilder` |
| `RampAlignment_Rev14.bas` | `RampAlignment` |
| `HCHeatmap_Rev14.bas` | `HCHeatmap` |

**Note:** Rev14 has no Migration module. The Migration module from Rev13 is not carried forward. Schema upgrades are handled by the CreateWorkingSheet full-build path.

### How Deployment Works

1. Copy changed `.bas` files to the new folder; rename with the new `_RevNN` suffix
2. Keep `Attribute VB_Name` as the stable name inside each file — no RevNN in VB_Name
3. Update `TIS_VERSION = "RevNN"` in `TISCommon.bas` — the only required change when bumping a version
4. Make functional code changes inside module bodies
5. In Excel: `StripAllModules` -> `LoadAllModules` (point at the new folder via file dialog)
6. Done — no Launcher edits, no cross-module call edits required

### How LoadAllModules Works

Opens a file dialog for the user to select the folder containing `.bas` files. Reads the `Attribute VB_Name = "..."` line from inside each `.bas` file (not from the filename). `WorkfileBuilder_Rev14.bas` with `Attribute VB_Name = "WorkfileBuilder"` imports as `WorkfileBuilder`, replacing the previous version correctly.

### How StripAllModules Works

Iterates all loaded standard modules, strips any `_RevNN` suffix via `GetModuleBaseName`, checks against the known TIS module base names list, removes any match. Automatically handles any past or future revision number. Self-protects `TIS_Launcher` (the running module).

---

## Module Overview

> **How they connect:** TIS_Launcher is the entry point. Each step reads from the Working Sheet (the central hub) and writes back to it or to a dedicated output sheet. TISCommon is loaded invisibly by every module.

```
Customer TIS File
       |
       v
  [Step 1] TISLoader (4-step flow)
       |  1. Backup TIS -> TISold (copy values)
       |  2. Load new TIS file into TIS sheet
       |  3. Generate TIScompare (TIS vs TISold, user review only)
       |  4. UpdateWorkingSheetFromTIS — in-place update of Working Sheet
       |     Updates changed TIS cells (orange fill + comment)
       |     Cancels removed systems (Status=Cancelled)
       |     Re-activates returning systems (Status=Active)
       |     Appends new systems with Our Dates from TIS (blue border)
       |     Sorts Working Sheet (SortWorkingSheet)
       |     Rebuilds Gantt + NIF at end
       v
  +----------------------------------------------------+
  |                 Working Sheet (PERMANENT)           |
  |  Our Dates | Status | Health | WhatIf | Lock? | ...|
  |  (ListObject table — never destroyed in normal ops) |
  +------------------------+---------------------------+
                           |
          +----------------+-------------+--------------+
          v                v             v              v
    [Step 3]          [Step 4]       [Step 5]        Standalone
    GanttBuilder      NIF_Builder    DashboardBuilder  RampAlignment
    Reads Our Dates   Filters        Filters Status    Filters Status
    (redirected from  Status         in all KPIs
     TIS date cols)

[Rare] Build Working Sheet (WorkfileBuilder.CreateWorkingSheet)
  Full rebuild from scratch — for schema changes, corruption, or first-time setup
  Backs up old sheet as Old YYYY-MM-DD, clears, rebuilds on SAME sheet object

[Standalone] WhatIf Gantt  <- user-triggered
  Backs up Our Dates to hidden sheet
  Shifts Our Dates by WhatIf delta
  Rebuilds Gantt (reads shifted Our Dates)
  Restores from backup on deactivate
```

---

## Rev14 — Design (Implemented)

> Rev14 is a clean rewrite from the Rev12 baseline. All modules in `SA\12\` are the starting point. Implemented in `C:\Users\razzl\Desktop\SA\14\`. Rev14 is NOT incremental from Rev13. Because Rev12 introduced stable VB_Names, the Launcher required zero changes when importing Rev14 modules.

Rev14 introduces an update-in-place architecture, the Status lifecycle column, Health tracking, per-milestone deviations, and WhatIf scenario mode.

### The Core Change

**Before (Rev12):** TIS file -> Working Sheet rebuilt weekly. TIS was the source of truth. User annotations were fragile overlays that required migration each rebuild. The entire sheet was cleared and reconstructed every time.

**After (Rev14):** You own 8 committed milestone dates per system. The Working Sheet is permanent — never destroyed during normal operations. Weekly TIS loads update TIS columns in-place via TISLoader's 4-step flow — only changed cells are touched. New systems are appended, removed systems are marked Cancelled. Your dates stay. The Gantt draws from your dates. You see at a glance where TIS disagrees with your plan via Health status (Match/Minor/Gap). Build Working Sheet is a rare operation for schema changes or corruption recovery.

### Key Differences from Rev13

| Feature | Rev13 (designed) | Rev14 (implemented) |
|---|---|---|
| Our Date columns | 9 (including SDD) | 8 (no SDD — SDD stays as TIS column) |
| Effective Date columns | 9 hidden columns | None — Gantt reads Our Dates directly via redirect |
| Deviation columns | 9 formula columns in sheet | Health is a live formula (Match/Minor/Gap) — no separate deviation columns in sheet |
| WhatIf mechanism | Writes to Effective Date columns | Backup-restore approach (copies Our Dates to hidden sheet, shifts, restores) |
| Update path | Full rebuild every time | TIS Load (TISLoader 4-step) for weekly updates; CreateWorkingSheet (rare) for schema changes/corruption |
| Status values | Active, Completed, On Hold, Cancelled | Active, Completed, On Hold, Non IQ, Cancelled |
| Migration module | Standalone module | Not included — CreateWorkingSheet handles schema upgrades |
| Zone header colors | Fully applied | Defined in TISCommon but not yet fully applied to Working Sheet formatting |

---

## New Columns in the Working Sheet

Inserted after the Group/base data columns, before TIS date columns.

### Our Date Columns (user-owned, values only)

| Header | Meaning | Gantt phase |
|--------|---------|-------------|
| `Set` | Our committed SET start | SET phase start |
| `SL1` | Our committed SL1 Signoff Finish | SET phase end / SL2 phase start |
| `SL2` | Our committed SL2 Signoff Finish | SL2 phase end |
| `SQ` | Our committed Supplier Qual Finish | SQ phase end |
| `Conv.S` | Our committed Conversion Start | CV phase start |
| `Conv.F` | Our committed Conversion Finish | CV phase end |
| `MRCL.S` | Our committed MRCL Start | MRCL phase start |
| `MRCL.F` | Our committed MRCL Finish | MRCL phase end |

**8 Our Date columns total.** SDD is not included — the TIS SDD column serves as the SDD source for the Gantt.

- **Values only — never formulas.** No automation writes to these unless explicitly unlocked (see Lock? column).
- **Distinct header style** (THEME_ACCENT background, white text) so users immediately recognise "our zone".
- **Auto-populated for new systems** (first time a project appears in TIS): Our Dates initialized from TIS dates. Those cells get a **blue border** to signal "auto-filled — please review and confirm."

**Demo systems exception:** Demo systems (Event Type = "Demo") do not use Our Date columns. The Gantt reads TIS date columns directly for Demo/PreFab milestones.

### Status Column

`Status` — dropdown with five values: `Active`, `Completed`, `On Hold`, `Non IQ`, `Cancelled`.

- **Active** — project is live in TIS and being tracked. Default for all new and returning projects.
- **Completed** — project has finished all milestones. Stays in Working Sheet for historical context. Replaces the old `Completed` TRUE/FALSE column (Rev14 merges it into Status).
- **On Hold** — project is paused. Included in some metrics, excluded from HC calculations.
- **Non IQ** — project is not in scope for IQ tracking. Excluded from relevant counts.
- **Cancelled** — project removed from TIS or manually cancelled. Excluded from all live counts, HC calculations, Dashboard KPIs, and RampAlignment.
- Removed systems **stay in the Working Sheet** (not archived). History preserved. When a project is removed from TIS, Status is set to `Cancelled`.
- If a system returns to a future TIS: Status resets to `Active`, TIS dates update, Our Dates unchanged.

### Lock? Column

`Lock?` — user-set boolean dropdown (TRUE/FALSE), default empty (unlocked).

When TRUE:
- **Data Validation rejects edits** to Our Date cells for that row. The user sees "Row Locked — set Lock? to FALSE to edit."
- **Automation skips locked rows** — WorkfileBuilder new-system initialization checks Lock? before writing to Our Date columns.
- Lock? itself is always editable (user can always toggle it on/off).

**Implementation notes:**
- Sheet protection has been **removed entirely** (caused CF failures on protected sheets in Excel 365). Lock? enforcement relies solely on Data Validation.
- Named ranges `OUR_DATE_START`, `OUR_DATE_END`, `LOCK_COL` are still created by WorkfileBuilder Cleanup for the optional Worksheet_Change handler.

**Real-time Lock enforcement (Data Validation — primary mechanism):**
- Each Our Date cell has a custom Data Validation formula: `=NOT($LockCol=TRUE)`
- When Lock?=TRUE: Excel rejects any edit with "Row Locked — set Lock? to FALSE to edit"
- When Lock?=FALSE/empty: edits allowed normally
- **No VBA required at runtime** — Excel evaluates the formula only when the user presses Enter
- **Survives workbook close/reopen** — validation is persistent, unlike sheet protection
- **Visual indicator:** CF rule grays out locked Our Date cells (SLATE_200 background, SLATE_500 text)
- **Performance:** Near-zero. One formula eval per cell edit attempt.

**Secondary enforcement (Worksheet_Change handler — optional):**
- A `Worksheet_Change` handler can be injected by `InstallSheetEvents` for belt-and-suspenders protection
- **Requires:** "Trust access to the VBA project object model" in Trust Center
- **Not required** — Data Validation alone is sufficient for Lock? enforcement

Typical use: lock a row once the committed date has been formally agreed with the customer and should not drift.

### Health Column

`Health` — a **live formula** that auto-recalculates from the maximum deviation between Our Dates and TIS Dates across all 8 milestone pairs. No rebuild needed — editing an Our Date cell immediately updates Health.

| Condition | Health Value | CF |
|---|---|---|
| All deviations <= 0 | `Match` | Green background |
| Any deviation 1-3 days | `Minor` | Amber background |
| Any deviation > 3 days | `Gap` | Red background |

Health is a formula column — users do not edit it. It provides an instant visual summary of schedule alignment. Demo rows get blank Health.

### WhatIf Column

`WhatIf` — user enters a hypothetical new project start date. See [WhatIf Feature](#whatif-feature) below.

---

## TIS Upload Behavior

What happens when TISLoader Step 4 (`UpdateWorkingSheetFromTIS`) runs during a weekly TIS load.

| Scenario | TIS Date columns | Our Date columns | Status | Visual |
|---|---|---|---|---|
| Project present, no changes | Updated (same values) | Unchanged | Unchanged | No change |
| Project present, TIS date changed | Updated with new TIS date | **Unchanged** | Unchanged | Orange fill on changed TIS cell + appended comment. Health recalculates. |
| Project present, CEID changed | CEID updated | Unchanged | Unchanged | Orange fill + comment on CEID cell. STD durations recalculate. Group unchanged. |
| Project present, other non-key field changed | Field updated | Unchanged | Unchanged | Orange fill + comment on changed field |
| New project (first time in TIS) | Populated from TIS | **Initialized from TIS** | **Active** | Blue border on auto-filled Our Date cells. Blue border on Entity Code cell. |
| Project removed from TIS (Active) | Retain last known TIS dates | Unchanged | **Cancelled** | Row stays; can sort/filter by Status. Only Active rows are auto-cancelled — Completed, On Hold, and Non IQ rows are protected from auto-cancellation. |
| Removed project returns to TIS | Updated with new TIS dates | Unchanged | **Active** | Orange fill where TIS dates differ from our last-known state |

**Key fields** (Site, Entity Code, Event Type) define project identity. They are never updated by TIS upload. If Event Type changes in TIS, it is treated as a different project: old key -> Status=Cancelled, new key -> added as new system.

**CEID** is not a key field. CEID changes in TIS update the CEID cell and trigger STD duration recalculation. Group is derived from Entity Type, not CEID — Group does not change when CEID changes.

---

## KPI Card Status Filtering

| Card | What it counts | Status filter |
|------|---------------|---------------|
| Total Systems | All tracked systems | Active + On Hold |
| New | New installations | Active + On Hold |
| Reused | Reinstalled systems | Active + On Hold |
| Demo | Removal systems | Active + On Hold |
| CT Miss | Cycle time misses | Active + On Hold |
| Escalated | Escalated systems | Active + On Hold |
| Watched | Watched systems | Active + On Hold |
| Conversions | Conversion systems | Active + On Hold |
| Completed | Finished systems | Status = "Completed" specifically |

Excluded from all counts (except Completed card): Cancelled, Completed, Non IQ.

---

## WhatIf Feature (Implemented)

**Status:** COMPLETE. Entry points: `WorkfileBuilder.ActivateWhatIfMode` / `WorkfileBuilder.DeactivateWhatIfMode`. Toggle: `TIS_Launcher.ToggleWhatIf`.

**Purpose:** Explore "what if we shift project X's start date?" without permanently changing any committed data. The Gantt and HC tables reflect the hypothetical scenario. Your Our Dates are restored when you exit WhatIf mode.

### How to Use

1. Type a new start date into the `WhatIf` column for any project(s)
2. Press **"WhatIf"** button (on Working Sheet near HC/Gantt toggle, or on Instructions sheet)
3. WorkfileBuilder:
   a. Backs up all Our Date values to a hidden `WhatIf_Backup` sheet
   b. Computes the delta: `WhatIf date - project start date` (where project start = first non-null Our Date in order: Set, SL1, SL2, SQ, Conv.S, Conv.F, MRCL.S, MRCL.F)
   c. Shifts Our Date columns in place by the delta (SDD is not an Our Date column — it is unaffected)
   d. Rebuilds Gantt from the shifted Our Dates
   e. Recomputes Health
4. **Mode indicator:** The `WhatIf_Backup` sheet's existence signals WhatIf mode is active.

### Restoring Normal Mode

- Press the **"WhatIf"** button again (toggle behavior) or **"Restore Normal Gantt"** on Instructions sheet
- Our Dates are restored from the backup sheet
- Gantt rebuilds normally
- `WhatIf_Backup` sheet is deleted
- WhatIf column values are **preserved** (not cleared) — user can re-run the scenario later

### What Shifts in WhatIf Mode

- All 8 Our Date columns shift by the delta
- SDD is a TIS column, not an Our Date column — it does not move
- Systems without a WhatIf date are unchanged
- Non-Active rows (Status = Cancelled, etc.) are skipped

### Guard Against Double-Activation

If `WhatIf_Backup` already exists when ActivateWhatIfMode is called, it first deactivates (restores) before re-activating. This prevents data corruption from double-shifting.

### WhatIf + HC Analyzer

The HC analyzer and HC graphs are derived from the Gantt. The Gantt reads Our Date columns. Therefore, when WhatIf is active and the Gantt is rebuilt from shifted Our Dates, the HC analyzer and all HC graphs automatically reflect the WhatIf scenario — no additional wiring is needed.

---

## TIS Change Tracking

When UpdateWorkingSheetFromTIS detects that a non-key TIS field has changed for an existing project:

1. The cell is updated with the new TIS value
2. **Orange background fill** is applied (`Interior.Color = CLR_CHANGE_FILL` — not CF — so Sort by Color works)
3. An **appended comment** is written: `"[YYYY-MM-DD] Changed from: [old value]"`. If a comment already exists, the new entry is appended with a line break. The full history accumulates in the comment.

---

## Working Sheet Sorting (SortWorkingSheet)

Called after TIS load (UpdateWorkingSheetFromTIS) AND after full rebuild (CreateWorkingSheet).

**Primary sort:** Status custom order — Active=1, Completed=2, On Hold=3, Non IQ=4, Cancelled=5.
**Secondary sort:** Project start date ascending — MIN of all Our Date columns for each row.

**Implementation:** Uses temporary helper columns (inserted, populated, sorted on, then deleted after sort). Cancelled projects always sort to the bottom. Within each status group, systems are ordered by their earliest committed milestone date.

---

## Visual Design — CommitmentTracker Light Theme

Rev14 implements zone-colored headers and a light theme for the Working Sheet.

### Title Bar and Subtitle Bar
- **Row 1:** Title bar ("TIS Commitment Tracker") — Navy background, white 16pt bold, 36px height
- **Row 2:** Subtitle bar with usage hints and version — Steel blue background, Slate gray 9pt text, 20px height

### Zone Category Bar (Row 13)
Row 13 contains merged, color-coded labels above the column headers identifying each zone:
- "IDENTITY" — Navy background
- "OUR COMMITMENT DATES (editable)" — Deep green background
- "TIS DATES (auto-updated)" — Deep blue background
- "ANALYSIS & USER FIELDS" — Deep amber background

### Zone-Colored Headers (Row 14)
Column headers are color-coded by functional zone:

| Zone | Background | Font Color | Columns |
|------|-----------|------------|---------|
| Identity | Navy (`ZONE_IDENTITY_BG`) | White | Site, Entity Code, Entity Type, CEID, Group, Event Type |
| Our Dates | Deep Green (`ZONE_USER_BG`) | Light Green (`ZONE_USER_FG`) | Set, SL1, SL2, SQ, Conv.S, Conv.F, MRCL.S, MRCL.F, Status, Lock?, Health, WhatIf |
| TIS Dates | Deep Blue (`ZONE_TIS_BG`) | Light Blue (`ZONE_TIS_FG`) | All TIS date columns |
| Milestone Analysis | Deep Amber (`ZONE_CALC_BG`) | Light Amber (`ZONE_CALC_FG`) | Actual Duration, STD Duration, Gap columns |
| User Fields | Deep Amber (`ZONE_CALC_BG`) | Light Amber (`ZONE_CALC_FG`) | New/Reused, Escalated, Ship Date, SOC, Comments, BOD |

### Column Grouping (Collapsible)
These column sections are grouped and can be expanded/collapsed via the +/- outline buttons:
- TIS date columns (collapsed by default)
- SOC through BOD2 user field columns
- STD Duration columns (via Definitions sheet grouping config)

### Light Theme
- **Background:** White (xlNone) for all data cells
- **Font:** Segoe UI 9pt for headers, 10pt for data
- **Table style:** TableStyleLight1 for the ListObject
- **Zebra striping:** Alternating white / near-white rows

### Health Status Conditional Formatting
| Status | Background | Font Color |
|--------|-----------|------------|
| Match | Green (`STATUS_ONTRACK_BG`) | Dark green (`STATUS_ONTRACK_FG`) |
| Minor | Amber (`STATUS_ATRISK_BG`) | Dark amber (`STATUS_ATRISK_FG`) |
| Gap | Red (`STATUS_BEHIND_BG`) | Dark red (`STATUS_BEHIND_FG`) |

---

## Module Details

### TISCommon.bas — Shared Foundation
- All `THEME_*` color constants (brand palette)
- All `TIS_SHEET_*` sheet name constants
- All `TIS_COL_OUR_*` Our Date column header constants (8 columns)
- All `TIS_SRC_*` TIS source column name constants (9 — including SDD)
- `TIS_COL_STATUS`, `TIS_COL_LOCK`, `TIS_COL_HEALTH`, `TIS_COL_WHATIF`
- `CLR_CHANGE_FILL` (orange), `CLR_NEW_DATE_BORDER` (blue)
- `STATUS_ONTRACK_*`, `STATUS_ATRISK_*`, `STATUS_BEHIND_*` Health CF colors
- `AppState` type + `SaveAppState` / `RestoreAppState` / `SetPerformanceMode`
- Utility functions: `ColLetter`, `SheetExists`, `FileExists`, `FindWorkingSheet`, `FindWorkingSheetTable`, `FindHeaderRow`, `FindHeaderCol`, `BuildProjectKey`, `FormatCardStyle`, `GetMilestoneStartHeaders`, sorting helpers, `DebugLog`

**Critical rule:** Never rename or remove any Public function or constant without auditing all callers.

---

### Step 1 — TISLoader
**What it does:** Loads the new TIS file via a 4-step flow that updates the Working Sheet TIS columns in-place.

**Architecture note:** TISLoader now includes Working Sheet update as Step 4 (UpdateWorkingSheetFromTIS). This is intentional — the old architecture where "Step 1 does NOT touch Working Sheet" is retired. TISLoader is the primary weekly workflow: load TIS, generate compare, update Working Sheet in-place. WorkfileBuilder's CreateWorkingSheet is only for first-time build or schema upgrades.

**Rev14 4-step flow:**
1. **Backup TIS -> TISold** — Copies the current TIS sheet values to a `TISold` sheet (value copy, not rename). Preserves the previous week's data for comparison.
2. **Load new TIS file into TIS sheet** — Opens the customer TIS file (user picks), copies it into the `TIS` sheet.
3. **Generate TIScompare** — Diffs TIS vs TISold to produce the `TIScompare` sheet. TIScompare is for **user review only** — it has no buttons or automation triggers.
4. **UpdateWorkingSheetFromTIS** — In-place update of Working Sheet TIS columns. Compares TIS data against existing Working Sheet rows, updates changed fields (orange fill + comment), cancels removed systems (`Status=Cancelled`), re-activates returning systems (`Status=Active`), appends new systems with Our Dates initialized from TIS (blue border).

**Removed:** The Apply button, `ApplyChangesToWorkingSheet`, and `ArchiveRemovedRows` are all removed. TIScompare is informational only.

**Reads:** TIS file (user picks), TISold, Definitions, CEIDs
**Writes:** TIS sheet, TISold sheet, TIScompare sheet, Working Sheet (TIS columns updated in-place)

---

### Step 2 — WorkfileBuilder
**What it does:** Core build module. Provides the full-build path for schema setup and the WhatIf scenario mode.

**Working Sheet architecture:**
- The Working Sheet is **PERMANENT** — it is never destroyed and never recreated during normal weekly operations.
- **TIS Load (Step 1)** is the standard weekly flow. `UpdateWorkingSheetFromTIS` (called by TISLoader) updates TIS columns in-place, non-destructive to user data (Our Dates, Status, Lock?, comments, NIF assignments, BOD, flags).
- **Build Working Sheet (`CreateWorkingSheet`)** is a **rare operation** — used only for schema changes, corruption recovery, or first-time setup. It clears and rebuilds on the **same sheet object** (preserves cross-sheet formula references from other sheets like `='Working Sheet'!A15`).

**WorkfileBuilder fast-update path (If True Then guard):** Intentionally disabled. The user decided that Build Working Sheet should always do a full rebuild. The in-place TIS update is handled by TISLoader.UpdateWorkingSheetFromTIS, not by WorkfileBuilder. The `If True Then` guard at ~line 224 forces every Build Working Sheet through CreateWorkingSheet. This is by design — Build Working Sheet is a rare operation (schema change, corruption, first-time setup). Weekly TIS updates use Load TIS -> UpdateWorkingSheetFromTIS.

**CreateWorkingSheet — full build from scratch:**
- Used when no Working Sheet exists, or existing sheet lacks Rev14 schema (detected by absence of `Conv.S` header)
- Backs up existing sheet as `Old YYYY-MM-DD`, clears, rebuilds with full column schema including Our Dates, Status, Lock?, Health, WhatIf
- Imports user data from old sheet (Our Dates, Status, NIF, BOD, flags)

**Rev14 features:**
- Our Date columns (8), Status, Lock?, Health, WhatIf added to column schema
- Status logic: Cancelled for removed projects, Active for active/returning. Five-value dropdown: Active, Completed, On Hold, Non IQ, Cancelled
- Completed column merged into Status (Completed=TRUE -> Status="Completed" during import from pre-Rev14)
- Change detection: orange fill + appended comment for changed non-key TIS fields
- New system initialization: Our Dates populated from TIS dates + blue border
- Health column: LIVE FORMULA (auto-recalculates on Our Date edit, no rebuild needed)
- WhatIf mode: backup-restore approach via ActivateWhatIfMode / DeactivateWhatIfMode
- Sheet protection in Cleanup: locks Our Date cells where Lock?=TRUE, protects with AllowFiltering + AllowSorting
- SortWorkingSheet called after build and after TIS load

**Full build critical sequence (ClearWorkingSheet):**
1. Delete slicer caches
2. Delete ListObjects
3. Clear conditional formatting
4. Delete shapes and chart objects
5. Unmerge cells, clear outline
6. Delete named ranges
7. Unprotect sheet
8. Remove freeze panes
9. Clear all content, formats, validations
10. Reset column widths and row heights

**Reads:** TIS, Definitions, CEIDs, Milestones, New-reused, SN sheets
**Writes:** Working Sheet (in-place), Old YYYY-MM-DD (backup, full build only)

---

### Step 3 — GanttBuilder
**What it does:** Adds a formula-based Gantt chart to the right of the Working Sheet data.

**Rev14 implementation:**
- Reads phase boundaries by **redirecting TIS date column lookups to Our Date columns** for install/reused milestones (Set, SL1, SL2, SQ, Conv.S, Conv.F, MRCL.S, MRCL.F)
- Demo and PreFab milestones still read TIS date columns directly (no Our Date equivalent)
- Redirect mechanism: `ourDateRedirect` dictionary maps TIS source header (e.g., `Set Start`) to Our Date header (e.g., `Set`). If the Our Date column exists in the header map, the phase column index is overridden. Otherwise falls back to TIS column.
- A "TIS Gantt" view (showing TIS dates instead of Our Dates) is planned for future.

**Reads:** Working Sheet (Our Date columns via redirect, TIS columns for Demo/PF), Definitions
**Writes:** Working Sheet (Gantt columns appended to right)

**Status-based Gantt filtering:** The Gantt wraps each cell's phase formula with `IF(Status="Active", [formula], "")`. Only Active systems show Gantt bars. Cancelled, Completed, On Hold, and Non IQ rows display blank cells in the Gantt area. The rows remain in the sheet structure — they are not deleted.

---

### Step 4 — NIF_Builder
**What it does:** Appends NIF assignment columns and builds 8 HC Analyzer tables below the data.

**Rev14 additions:**
- HC Need/Available formulas filter out Cancelled, Completed, and Non IQ systems (only Active + On Hold contribute to HC calculations)

**HC table titles (contracts — Dashboard finds by exact text):**
1. `New Systems - HC Need`
2. `Reused Systems - HC Need`
3. `Combined - HC Need`
4. `New Available HC (from NIF)`
5. `Reused Available HC (from NIF)`
6. `New HC Gap (Available - Need)`
7. `Reused HC Gap (Available - Need)`
8. `Total HC Gap`

**Reads:** Working Sheet
**Writes:** Working Sheet (NIF columns + HC tables below data)

---

### Step 5 — DashboardBuilder
**What it does:** Builds the management Dashboard sheet.

**Rev14 additions:**
- All KPI cards filter by Status: Active + On Hold only (excludes Cancelled, Completed, Non IQ). The Completed card counts Status="Completed" specifically.
- DashHelper table excludes Cancelled, Completed, and Non IQ rows
- Escalation Tracker skips Cancelled, Completed, and Non IQ rows
- PivotCharts reflect Active + On Hold systems only

**Contains:**
- KPI cards: Total, New, Reused, Demo, CT Miss, Escalated, Watched, Conversions, Completed
- Group dropdown filter (`DASH_KPI_GROUP`)
- System Counters with collapsible CEID drill-down
- HC Gap Analysis
- PivotCharts with start date selector (`DASH_CHART_START`)
- Escalation Tracker table
- **Install Base graph** — stacked column chart showing cumulative systems in production per quarter. Existing (AMAT Blue) + New (Emerald) + Reused (Teal). IB_BASELINE named range for user-entered starting tool count. Separate PivotCache with 5 independent slicers (NewReused, Group, CEID, EntityType, HasSetStart). Uses Our MRCL.F for install completion date.
- Injected `Worksheet_Change` handler

**Reads:** Working Sheet (Status-filtered), Definitions, DashHelper
**Writes:** Dashboard sheet, DashHelper sheet

---

### RampAlignment — Standalone
**Rev14:** Filters `Status = "Active"` rows.
**Reads:** Working Sheet (Status-filtered), Definitions, Milestones
**Writes:** `Ramp - {GroupName}` sheets

---

### HCHeatmap — Standalone
**Rev14:** Rev12 version (proven working) with WhatIf button addition. The WhatIf button (`btn_WhatIf`) calls `TIS_Launcher.ToggleWhatIf`. No changes to core heatmap logic. Reads from the Gantt area which is built from Our Dates.

---

### TIS_Launcher — Orchestrator
**Rev14 additions:**
- `ActivateWhatIf` / `DeactivateWhatIf` / `ToggleWhatIf` public subs
- `ToggleWhatIf` checks for `WhatIf_Backup` sheet existence to determine current mode
- Instructions sheet includes "Build WhatIf Gantt" and "Restore Normal Gantt" buttons under "WhatIf Scenario" section
- `LoadAllModules` updated: uses file dialog for folder selection (user picks the folder)
- No Migration button (Migration module removed)

---

## Sheet Reference

| Sheet | Created By | Who Edits It | Purpose |
|-------|-----------|-------------|---------|
| TIS Tracker | TIS_Launcher | Nobody | Instructions + workflow buttons |
| TIS | TISLoader | Nobody | Raw TIS data import (current week) |
| TISold | TISLoader | Nobody | Previous week TIS backup (value copy) |
| TIScompare | TISLoader | Nobody | Week-over-week diff (user review only, no action buttons) |
| Definitions | User setup | User | Config: filters, sort, milestones, Gantt columns |
| CEIDs | User setup | User | Entity Type -> Group mapping |
| Milestones | User setup | User | Standard durations per CEID/milestone |
| New-reused | User setup | User | Entity Code -> New/Reused classification |
| Working Sheet | WorkfileBuilder | User (Our Dates, NIF, BOD, flags, Status) | Central data hub — operational database |
| Old YYYY-MM-DD | WorkfileBuilder | Nobody | Previous week backup (full build only) |
| WhatIf_Backup | WorkfileBuilder | Nobody | Hidden backup of Our Dates during WhatIf mode |
| Dashboard | DashboardBuilder | Nobody | Management dashboard |
| DashHelper | DashboardBuilder | Nobody | Hidden PivotTable helper |
| HC CF Definitions | HCHeatmap | Nobody | Persistent CF rules for heatmap |
| Ramp - {Group} | RampAlignment | Nobody | Per-group customer reports |
| ~~Removed Systems~~ | -- | -- | **Retired in Rev14.** Removed projects stay in Working Sheet with `Status="Cancelled"`. |

---

## Technical Skills Reference

Skills and patterns for delivering production-quality Excel VBA applications. Reference this section when making code changes.

### Excel VBA Performance

| Pattern | Do | Don't |
|---------|-----|-------|
| **Data I/O** | `Range.Value` bulk read/write (O(1) COM call) | Cell-by-cell loops (`ws.Cells(r,c).Value` in a loop) |
| **Dictionary** | Create once outside loops | `CreateObject("Scripting.Dictionary")` inside loops |
| **UsedRange** | Cache in a variable, call once | Call `.UsedRange` repeatedly |
| **Screen updates** | `SaveAppState` / `SetPerformanceMode` / `RestoreAppState` | Leave ScreenUpdating on during builds |
| **Comments** | Acceptable with SetPerformanceMode (~1500 ops = 1-2s) | Thousands of AddComment calls without ScreenUpdating=False |
| **String concatenation** | Build strings with `&` or use arrays | Use Mid$/Replace in tight loops |
| **Line continuations** | Max 24 per statement. Use `s = s & "..."` pattern for long strings | Single statement with 25+ `_` continuations (compile error) |

### Chart & Graph Best Practices

| Rule | Why |
|------|-----|
| **Light text on dark backgrounds** | Chart areas use navy backgrounds — use RGB(226,232,240) for titles/legends, RGB(180,190,200) for axis labels |
| **Dark text on light backgrounds** | Data tables, KPI cards use white bg — use RGB(30,41,59) Slate-900 |
| **Gridlines subtler than data** | Use RGB(40,65,100) on dark plots. Never same color as axis labels |
| **Consistent series colors** | New=THEME_SUCCESS (Emerald), Reused=THEME_ACCENT2 (Teal), Demo=THEME_DANGER (Coral), Existing=THEME_ACCENT (Blue) |
| **Data labels on totals only** | Stacked charts: label the total, not each segment |
| **Chart area border** | Thin line RGB(50,75,110) — visible but not heavy |

### UX / UI Patterns

| Pattern | Implementation |
|---------|---------------|
| **Zone-colored headers** | Use `TISCommon.ApplyZoneHeader(rng, bgColor, fgColor)`. Each column zone gets a distinct dark bg with matching light fg. |
| **Zone category bar** | Use `TISCommon.ApplyZoneCategoryLabel(ws, row, startCol, endCol, label, bgColor)`. Merged cells, 8pt bold white text. |
| **Title bar** | Use `TISCommon.ApplyTitleBar(ws, lastCol, title)`. Row 1, navy, white 16pt bold, 36px height. |
| **Subtitle bar** | Use `TISCommon.ApplySubtitleBar(ws, lastCol, subtitle)`. Row 2, steel blue, slate 9pt, 20px height. |
| **KPI cards** | Use `TISCommon.FormatCardStyle(rng, bgColor, accentColor)`. White bg, colored left border, slate label text. |
| **Toggle buttons** | GanttBuilder `CreateSegmentedToggle` pattern. Two-segment button with active/inactive states. |
| **Mode indicators** | WhatIf mode: sheet tab color change + banner text. Multiple simultaneous signals for mode clarity. |
| **Change indicators** | `Interior.Color` (not CF) for orange change fills — enables Sort by Cell Color. |
| **Status indicators** | CF rules with `StopIfTrue` for Health: Match (green), Minor (amber), Gap (red). |
| **Smooth toggles** | Save/restore column widths around Gantt rebuilds. Use `Application.ScreenUpdating = False` throughout. Activate Working Sheet at end. |

### Sheet Protection — REMOVED

Sheet protection has been **removed entirely** from all modules. `FormatConditions.Add` with `xlExpression` type fails silently on protected sheets in some Excel 365 builds, which was the root cause of Gantt CF not showing after TIS load.

Lock? enforcement uses **Data Validation** (primary) — no sheet protection needed. Named ranges `OUR_DATE_START`, `OUR_DATE_END`, `LOCK_COL` are still created for the optional Worksheet_Change handler.

### Conditional Formatting Patterns

| Purpose | Method | Why |
|---------|--------|-----|
| **Health status** (Match/Minor/Gap) | CF rules with `StopIfTrue = True` | Persists, auto-updates, color-coded |
| **TIS change indicators** (orange fill) | `Interior.Color` (direct format) | Enables "Sort by Cell Color" in Excel |
| **New system markers** (blue border) | `Borders` (direct format) | Visible signal, survives rebuild |
| **Gantt phase colors** | CF rules per phase | CF persists when data changes, survives Undo |
| **Today marker** (green column) | CF rule with `TODAY()` formula | Auto-updates daily without rebuild |

### Error Handling Pattern

```vba
Public Sub DoSomething()
    On Error GoTo ErrorHandler
    Dim appSt As AppState
    appSt = SaveAppState()
    SetPerformanceMode
    ' ... work ...
    GoTo Cleanup
ErrorHandler:
    MsgBox "Error in DoSomething: " & Err.Description, vbCritical
    DebugLog "MODULE ERROR: " & Err.Description
Cleanup:
    RestoreAppState appSt
    Set m_ws = Nothing
End Sub
```

### Cross-Module Call Safety

```vba
' Always check errors after cross-module calls:
On Error Resume Next
GanttBuilder.BuildGantt silent:=True, targetSheet:=ws
If Err.Number <> 0 Then
    DebugLog "GanttBuilder failed: " & Err.Description
    Err.Clear
End If
On Error GoTo ErrorHandler
```

### VBA Gotchas (Quick Reference)

| Gotcha | Rule |
|--------|------|
| `Dim` inside loops | VBA hoists to procedure level — variable doesn't reset per iteration |
| `On Error Resume Next` scope | Persists until explicitly reset. Always follow with `On Error GoTo 0` |
| `Range.Value` single cell | Returns scalar, not array. Guard with `If rowCount = 1 Then` |
| PivotItem hide sequence | Two-phase: all `Visible=True` first, then hide targets |
| Slicer cleanup order | Delete slicer caches BEFORE ListObjects BEFORE cells |
| `FormatConditions.Add` on protected sheets | Fails silently with `xlExpression` in some Excel 365 builds even with `UserInterfaceOnly:=True`. Sheet protection removed entirely. |
| Line continuations | Max 24 `_` per statement. Use `s = s & "line"` pattern for long injected code strings. |
| `SlicerCaches.Add2` | Required for Excel 365. `Add` may fail silently. |
| Multi-line column headers | Strip `vbLf`/`vbCr` when comparing: `LCase(Trim(Replace(Replace(val, vbLf, ""), vbCr, "")))` |

---

## Glossary

| Term | Meaning |
|------|---------|
| **TIS** | Tool Install Schedule — weekly customer-provided schedule file |
| **Ramp** | Fab Ramp — large-scale fab construction or upgrade (hundreds of systems) |
| **CEID** | Chamber Equipment ID — sub-configuration of an equipment model. Does NOT determine Group. |
| **Group** | Operational team (e.g., "Litho", "Etch"). Derived from Entity Type via CEIDs sheet, not from CEID. |
| **NIF** | Name In Frame — field engineer assigned to install/qualify a system |
| **HC** | Headcount — number of field engineers available for upcoming work |
| **BOD** | Blackout Date — date range when certain install phases are blocked |
| **CT** | Cycle Time — total install duration (Set -> MRCL) |
| **SDD** | System Dock Date — when the system physically arrives at the fab. TIS column only — not an Our Date column. |
| **SOC** | Statement of Compliance |
| **Our Dates** | The 8 committed milestone dates owned by the operations team. Values, never formulas. Never overwritten by automation (except WhatIf temporary shift with backup-restore). |
| **TIS Dates** | The corresponding dates from the customer TIS file. Updated on every TIS upload. Reference only. |
| **WhatIf** | Scenario tool. Enter a hypothetical new project start date; Gantt/HC rebuild to show the impact. Our Dates are backed up, shifted temporarily, then restored. |
| **Health** | Live formula from max deviation between Our Dates and TIS Dates: Match (all <= 0), Minor (any 1-3), Gap (any > 3). Green/amber/red CF. Auto-recalculates on Our Date edit. |
| **Status** | Dropdown: Active, Completed, On Hold, Non IQ, Cancelled. Replaces both the boolean Active? column and the Completed TRUE/FALSE column from pre-Rev14. Cancelled = excluded from all live counts. |
| **Non IQ** | Status value for systems not in scope for IQ (Installation Qualification) tracking. |
| **Project Key** | `Site \| Entity Code \| Event Type`. All three together identify a project. Key fields never change from TIS upload. |
| **Schema Detection** | Rev14 detects its own schema by checking for the `Conv.S` header in the Working Sheet. Present = schema current (weekly updates via TISLoader). Absent = full build needed (CreateWorkingSheet). |

### Milestone Phases (lifecycle order)

| Phase | Full Name | Our Date column | Notes |
|-------|-----------|-----------------|-------|
| PF | Pre-Fab | -- | Facility prep. Not in Our Dates. Reads TIS dates. |
| SET | Tool Placement | `Set` | First install milestone. Our Set = Set Start. |
| SL1 | Signoff Level 1 | `SL1` | Power-on + non-toxic utilities. Our SL1 = SL1 Signoff Finish. |
| SL2 | Signoff Level 2 | `SL2` | Toxic gases. Our SL2 = SL2 Signoff Finish. |
| CV | Convert | `Conv.S` / `Conv.F` | Conversion Start and Finish. May overlap other phases. |
| SQ | Supplier Qual | `SQ` | Supplier verification. Our SQ = Supplier Qual Finish. |
| MRCL | Material Release Checklist | `MRCL.S` / `MRCL.F` | Production-ready sign-off. Our MRCL.S = MRCL Start, MRCL.F = MRCL Finish. |
| DC | Decon | *(no Our Date — uses TIS dates)* | Demo/removal flow. |
| DM | Demo | *(no Our Date — uses TIS dates)* | Removal of existing system. |
| SDD | System Dock Date | *(no Our Date — uses TIS dates)* | Arrival date (single marker, not a duration phase). TIS column only. |

### System Types

| Type | Meaning |
|------|---------|
| New | New install — full milestone lifecycle |
| Reused | System being moved/reinstalled — may have reduced durations |
| Demo | Being removed — Decon/Demo/Move Out milestones |

---

## Open Design Questions

| # | Status | Question | Decision |
|---|---|---|---|
| **OD1** | Resolved | Demo project Our Dates | Demo systems use TIS dates directly. No Our Date columns for DC/DM. |
| **OD2** | Resolved | MRCL exact TIS column names | `MRCL Start` and `MRCL Finish`. Confirmed from TIS.xlsx reference file. |
| **OD3** | Resolved | WhatIf state persistence | WhatIf uses backup sheet approach. `WhatIf_Backup` existence = active mode. Toggle button for easy on/off. |
| **OD4** | Open | Sync Our Dates from TIS | Not yet implemented in Rev14. Planned enhancement. |
| **OD5** | Resolved | Lock? manual edit behavior | Data Validation only. Sheet protection removed — caused CF failures in Excel 365. |
| **OD6** | Resolved | Demo WhatIf | Non-Active rows skipped in WhatIf. Demo milestones read TIS dates directly. |
| **OD7** | Open | TIS Gantt view | Show Gantt using TIS dates instead of Our Dates. Planned future feature. |
| **OD8** | Resolved | Full zone header styling | Zone header colors fully implemented. Title bar, subtitle bar, zone category bar, and zone-colored headers all applied. |

---

## Users & Workflow

- **Users:** Ramp managers, ramp specialists, operations managers (Intel fab, primary)
- **Cadence:** TIS file loaded weekly; tracker updated after each load (fast path for weekly updates)
- **Completed systems:** Stay in tracker with `Status="Completed"` for historical context
- **Escalated / Watched:** Manual flags preserved across rebuilds
- **NIF roster:** 50+ engineers, currently free-form text. Master roster planned.
- **Reports:** Ramp Alignment reports presented at Intel meetings — PowerPoint export is critical
- **Web app:** `C:\Users\razzl\Desktop\SA\TISGANTTAPP\`. May replace Excel in the future.

---

## Known Bugs & Issues

### Fixed in Rev12

| # | File | Fix |
|---|------|-----|
| C1 | `TIS_Launcher` | Step subs use stable module names — no version suffix, no future version mismatch. |
| C2 | `TIS_Launcher` | `StripAllModules` uses pattern-based removal via `GetModuleBaseName`. |
| H1 | `GanttBuilder` | Removed silent `On Error Resume Next` before `NIF_Builder.BuildNIF` call. |
| H2 | `GanttBuilder` | `UsedRange` in `ClearExistingGantt` captured once into variable. |
| H3 | `DashboardBuilder` | PivotItem two-phase visibility fix. |

### Fixed in Rev14

| # | File | Fix |
|---|------|-----|
| C1 | `WorkfileBuilder` | ArchiveRemovedProjects replaced with Status=Cancelled (in-place). |
| C2 | `WorkfileBuilder` | WhatIf mode implemented (ActivateWhatIfMode / DeactivateWhatIfMode with backup-restore). |
| C3 | `WorkfileBuilder` | TIS column mappings corrected (MRCL Start/Finish confirmed via TIS_SRC_* constants). |
| C4 | `WorkfileBuilder` | Update-in-place path eliminates unnecessary full rebuilds for weekly TIS updates. |
| C5 | `WorkfileBuilder` | Completed column merged into Status dropdown (5 values including Non IQ). |
| C6 | `DashboardBuilder` | Status filter added to all KPI cards. |
| C7 | `DashboardBuilder` | DashHelper table now excludes Cancelled rows. |
| H4 | `GanttBuilder` | Our Date redirect mechanism: Gantt reads Our Dates for install/reused milestones. |
| H5 | `HCHeatmap` | WhatIf toggle button added to Working Sheet near HC/Gantt toggle. |
| H6 | `TIS_Launcher` | LoadAllModules uses file dialog for folder selection. |
| H7 | `TISLoader` | 4-step flow: Backup TIS->TISold, Load TIS, Generate TIScompare, UpdateWorkingSheetFromTIS. Apply button removed. |
| H11 | `TISLoader` | `UpdateWorkingSheetFromTIS` fixed: (1) ListObject now extended before bulk-writing new rows (were landing outside the table); (2) Gantt + NIF rebuild called at end of update; (3) `ws.Unprotect` uses empty password (consistent with WorkfileBuilder); (4) On Hold rows now cancelled when removed from TIS (previously only Active rows); (5) Blue border applied to auto-filled Our Date cells on new systems; (6) Eliminated duplicate bulk-reads (`wsDataArr`/`tisDataArr` → reuse existing `wsAllData`/`tisAllData`); (7) `Application.StatusBar` cleared in cleanup. |
| H12 | `TISLoader` | `CompareTISWorkflow` fixed: (1) `CreateChangeTrackingLog` removed — orange fills + comments in Working Sheet replace it; (2) TIScompare sheet now replaces existing sheet instead of accumulating "TIScompare1/2/3..."; (3) `ShowSummaryReport` call removed entirely — it compared TIS vs TISold and showed counts that didn't match the "TIS Update Summary" dialog (which compares TIS vs Working Sheet). Two mismatched dialogs were confusing. The TIScompare sheet provides the full TIS-vs-TISold diff; the single "TIS Update Summary" dialog from `UpdateWorkingSheetFromTIS` is the only notification now. |
| H13 | `TISLoader` | Duplicate key detection updated: duplicate rows are now ALL kept (not skipped). First occurrence goes into `tisKeyMap` normally; extra occurrences are appended via `tisDupExtraRows` collection. All duplicate rows get red `vbRed` medium border on Entity Code cell — both newly appended rows and any pre-existing WS rows sharing the key. Warning dialog updated: "All rows will be kept. Duplicate rows will be appended with red borders on Entity Code." |
| H14 | `GanttBuilder` | Gantt CF comprehensive fix: (1) Root cause: `SortWorkingSheet` re-protected the sheet before `BuildGantt` ran; `FormatConditions.Add` fails silently on protected sheets in some Excel 365 builds. Fixed by removing all sheet protection entirely. (2) `lastDataRow` used `End(xlUp)` which picked up HC table content — replaced with `GetDataLastRow()` using ListObject boundary. (3) `ws.Calculate` forced after formulas written so CF can match. (4) `.Formula2` fallback added for LET() formulas. (5) `ApplyTodayMarker` moved to after `ApplyGanttConditionalFormatting` so it isn't deleted by `FormatConditions.Delete`. (6) `DebugLog` added to every CF error guard for diagnostics. |
| H15 | `TISLoader` | `SortWorkingSheet` MIN formula changed from individual cell references to contiguous range reference (`MIN(G15:N15)`). Individual refs treat blank cells as 0. Range refs ignore blanks. Protection removed from sort. |
| H16 | `TISLoader` | Both `CompareTISWorkflow` and `UpdateWorkingSheetFromTIS` now use the same Definitions filters. Filters loaded once in `LoadNewTIS` and passed to both. TIScompare and Working Sheet update now show consistent counts. |
| H17 | Multiple | Sheet protection removed entirely from all modules. Was causing `FormatConditions.Add` failures on protected sheets (Excel 365 bug). Lock? enforcement uses Data Validation (primary) — no protection needed. Removed from: WorkfileBuilder (Cleanup, ClearWorkingSheet, WhatIf), TISLoader (UpdateWorkingSheetFromTIS, SortWorkingSheet), HCHeatmap (PaintHCHeatmap, RestoreGanttView), DashboardBuilder (ProtectDashboardSheet), RampAlignment, TIS_Launcher. |
| H18 | `TISLoader` | `UpdateWorkingSheetFromTIS` now has proper `On Error GoTo UpdateErrorHandler` with cleanup. Previously had no error handler — unhandled errors left Application in performance mode with events disabled. Early exits now route through `UpdateCleanup`. |
| H19 | `NIF_Builder` | Cleanup now restores `prevCalc` instead of forcing `xlCalculationAutomatic`. Previously destroyed caller's manual-calc state when called from within `SetPerformanceMode` blocks. |
| H20 | Multiple | `DATA_START_ROW` local constants now reference `TIS_DATA_START_ROW` from TISCommon instead of hardcoding 15. Changed in: WorkfileBuilder, TISLoader, GanttBuilder, DashboardBuilder, RampAlignment, HCHeatmap. |
| H21 | Multiple | Module header comments updated from "Rev11" to "Rev14" in RampAlignment and HCHeatmap. |
| H10 | `WorkfileBuilder` | WhatIf scroll-jump fixed: `ActiveWindow.ScrollRow/Col` saved after `SetPerformanceMode` and restored in Cleanup before `RestoreAppState` in both `ActivateWhatIfMode` and `DeactivateWhatIfMode`. Re-anchored after `Sheets.Add` which shifts ActiveSheet even with `ScreenUpdating=False`. |
| H8 | `WorkfileBuilder` | Health column changed from VBA-computed to live formula. Values: Match/Minor/Gap (thresholds 0/3). |
| H9 | `WorkfileBuilder` | SortWorkingSheet: Status custom order + project start date sort after TIS load and rebuild. |
| C8 | `DashboardBuilder` | KPI cards now exclude Completed+Cancelled+Non IQ (only count Active+On Hold). Completed card counts Status="Completed" specifically. |
| C9 | `DashboardBuilder` | Escalation Tracker skips Cancelled/Completed/Non IQ rows. |
| C10 | `RampAlignment` | RampAlignment filters Active-only systems in customer reports. |
| C11 | `NIF_Builder` | NIF HC Need/Available formulas filter out Cancelled/Completed/Non IQ. |
| C12 | `TISLoader` | TIS_SHEET_TIS -> SHEET_TIS crash fix in TISLoader. |
| C13 | `TISLoader` | TISLoader now owns Working Sheet update (UpdateWorkingSheetFromTIS) as Step 4 — intentional architecture change. |
| C14 | `TISLoader` | UpdateWorkingSheetFromTIS cleanup restores Application state (Calculation, Events, ScreenUpdating). |
| C15 | `TISLoader` | Only Active rows auto-cancelled when removed from TIS (Completed/On Hold/Non IQ protected from auto-cancellation). |
| C16 | `TISLoader` | Schema check (Conv.S sentinel) before UpdateWorkingSheetFromTIS — prevents running on pre-Rev14 sheets. |
| C17 | `RampAlignment` | RampAlignment now discovers and filters by Status column dynamically. |
| C18 | `TISCommon` | ZONE_OUR_BG and ZONE_TIS_BG now have distinct colors (green vs blue). |
| C19 | Multiple | All version strings updated to Rev14. |
| H22 | `NIF_Builder` | Error 1004 crash fix: `ApplyNIFOverlapCF`, `ApplyStaffedCF`, `ApplyHeatCS`, `ApplyGapCS` had NO error handling on `FormatConditions.Add`/`AddColorScale` calls. Any failure (invalid range, formula too long, Excel quirk) threw unhandled Error 1004, crashing BuildNIF before HC tables were built. Fixed: all four functions now have `On Error Resume Next` guards with `DebugLog` on failure, plus `If ld < fd Then Exit Sub` range validation. HC tables now always build even if CF application fails. |

### High — Fix Before Release

*(No outstanding high-priority issues.)*

### Medium — Fix When Convenient

| # | File | Issue |
|---|------|-------|
| M1 | `RampAlignment` | `m_msStartCols` used without null check after `DiscoverMilestoneColumns()`. |
| M2 | `NIF_Builder` | Header scan capped at column 50. Make cap dynamic. |
| M3 | `DashboardBuilder` | Chart start date default drifts. Define as constant or read from Definitions. |
| M4 | `HCHeatmap` | Cell tracking address string capped at 200 chars. Use helper column. |
| M5 | Multiple | `THEME_*` values duplicated as Private constants in some modules. Use TISCommon Public constants only. |
| M6 | `HCHeatmap` | OnAction string must use stable VB_Name: `"HCHeatmap.ToggleHCHeatmap"`. |
| M7 | `WorkfileBuilder` | Zone header colors not yet applied to Working Sheet formatting — uses THEME_ACCENT uniformly for Our Date/Status/Lock/Health/WhatIf headers. |
| M8 | -- | ~~Sheet protection uses empty password.~~ REMOVED — sheet protection no longer used. |
| M9 | `WorkfileBuilder` | Per-milestone deviation columns not yet implemented as separate sheet columns (Health is a live formula from raw Our Date vs TIS Date comparison). |

### Low / Observations

| # | File | Issue |
|---|------|-------|
| L1 | General | No rollback if a build step fails partway. |
| L2 | `DashboardBuilder` | KPI card rows written cell-by-cell. Acceptable now; bulk array writes would future-proof. |
| L3 | `TISLoader` | UpdateWorkingSheetFromTIS appends new rows cell-by-cell (not bulk). Acceptable for typical weekly additions (1-10 new systems). |

---

## Suggested Skills & Enhancements

### High Value
- **PowerPoint export** — Auto-generate `.pptx` from Ramp sheets for Intel meetings.
- **"Sync Our Dates from TIS" action** — Per-row or bulk accept of current TIS dates into Our Date columns. Reduces re-typing when user agrees with TIS.
- **Change summary notification** — After TIS load: formatted summary of added/removed/changed systems by group. Copy to clipboard for email.
- **Master NIF roster** — Managed engineer list. Enables typo prevention and HC math.
- **Holiday calendar** — Non-working days in cycle time and Gantt calculations.
- **TIS Gantt view** — Show Gantt from TIS dates for comparison with Our Dates Gantt.

### Medium Value
- **Per-milestone deviation columns** — Visible formula columns (Our Date - TIS Date) per milestone, complementing the Health live formula.
- **Per-project Our Date history** — Track every manual change to Our Dates with timestamp. See drift over time.
- **Parallel activity view** — Systems per group simultaneously active per milestone phase per week.
- **Config validation on startup** — Check required sheets on open.
- **Progress indicator** — `Application.StatusBar` messages during long builds.
- **Full zone header styling** — Apply ZONE_*_BG/FG constants to all header zones.

### Lower Priority
- **Automatic backup rotation** — Keep only last N `Old YYYY-MM-DD` sheets.
- **Dark/light theme toggle** — One-click for printing.
- **Export to .xlsx for web app** — Strip VBA, save for React companion app.
- **WhatIf multi-scenario** — Save and name multiple WhatIf scenarios.

---

## Planned Enhancements (Roadmap)

- TIS Gantt view (show Gantt from TIS dates)
- "Sync Our Dates from TIS" bulk accept action
- Per-project Our Date change history with timestamp audit trail
- Per-milestone deviation columns (visible in sheet)
- Full zone header styling
- WhatIf named scenarios
- Parallel activity view
- Cost tracking
- Holiday calendar integration
- Master NIF roster
- PowerPoint export for RampAlignment

---

<!-- ============================================================
  TECHNICAL REFERENCE
  (Dense implementation details for Claude — safe to skip when
   reading for editing. Do not edit casually.)
  ============================================================ -->

## [TECH] Naming Conventions

| Scope | Convention | Examples |
|-------|-----------|----------|
| Module-level variables | `m_camelCase` | `m_groupCol`, `m_workSheet`, `m_ptMonthly` |
| Private constants | `UPPER_SNAKE_CASE` | `DATA_START_ROW`, `TABLE_HEADER_BG` |
| Public constants | `UPPER_SNAKE_CASE` | `THEME_ACCENT`, `TIS_COL_OUR_SET` |
| Functions / Subs | `PascalCase` | `BuildHelperTable`, `PopulateHealthColumn` |
| Local variables | `camelCase` | `startRow`, `grpVal`, `wifDelta` |
| Loop counters | Short lowercase | `r`, `c`, `ri` |
| Parameters | `camelCase` | `silent As Boolean`, `grpName As String` |
| GoTo labels | `PascalCase` | `ErrorHandler:`, `Cleanup:` |

---

## [TECH] Error Handling Pattern

```vba
Public Sub DoSomething()
    On Error GoTo ErrorHandler
    Dim appSt As AppState
    appSt = SaveAppState()
    SetPerformanceMode
    ' ... work ...
    GoTo Cleanup
ErrorHandler:
    MsgBox "Error in DoSomething: " & Err.Description, vbCritical
    DebugLog "MODULE ERROR: " & Err.Description
Cleanup:
    RestoreAppState appSt
    Set m_ws = Nothing
End Sub
```

---

## [TECH] Rev14 Working Sheet Column Layout

```
[Identity]      Site | Entity Code | Entity Type | CEID | Group | Event Type
[TIS Dates]     TIS SDD | Set Start | SL1 Signoff Finish | ... (from Definitions)
[Our Dates]     Set | SL1 | SL2 | SQ | Conv.S | Conv.F | MRCL.S | MRCL.F
[Scenario]      Status | Lock? | Health | WhatIf
[Analysis]      Actual Duration × N | STD Duration × N | Gap × N
[User Fields]   New/Reused | Escalated | Tool S/N | Ship Date | Pre-Install Meeting |
                Est CAR Date | Est Cycle Time | SOC Available | SOC Uploaded? |
                Staffed? | Comments | BOD1 | BOD2
[Gantt]         [weekly columns, formula-based, CF-colored]
[NIF]           [NIF assignment columns, appended by NIF_Builder]
[HC Tables]     [below data rows, appended by NIF_Builder]
```

Header row: Row 14 (TIS_DATA_START_ROW - 1). Data start row: Row 15 (TIS_DATA_START_ROW = 15).
Zone category bar: Row 13.
Title bar: Row 1. Subtitle bar: Row 2.

---

## [TECH] Our Date + Operational Column Constants (TISCommon)

```vba
' Our Date column headers (exact strings -- never literals in module code)
' NOTE: No TIS_COL_OUR_SDD in Rev14 -- SDD is TIS-only
Public Const TIS_COL_OUR_SET    As String = "Set"
Public Const TIS_COL_OUR_SL1    As String = "SL1"
Public Const TIS_COL_OUR_SL2    As String = "SL2"
Public Const TIS_COL_OUR_SQ     As String = "SQ"
Public Const TIS_COL_OUR_CONVS  As String = "Conv.S"
Public Const TIS_COL_OUR_CONVF  As String = "Conv.F"
Public Const TIS_COL_OUR_MRCLS  As String = "MRCL.S"
Public Const TIS_COL_OUR_MRCLF  As String = "MRCL.F"

' TIS source column names (exact headers in TIS.xlsx)
Public Const TIS_SRC_SDD        As String = "SDD"
Public Const TIS_SRC_SET        As String = "Set Start"
Public Const TIS_SRC_SL1        As String = "SL1 Signoff Finish"
Public Const TIS_SRC_SL2        As String = "SL2 Signoff Finish"
Public Const TIS_SRC_SQ         As String = "Supplier Qual Finish"
Public Const TIS_SRC_CONVS      As String = "Convert Start"
Public Const TIS_SRC_CONVF      As String = "Convert Finish"
Public Const TIS_SRC_MRCLS      As String = "MRCL Start"
Public Const TIS_SRC_MRCLF      As String = "MRCL Finish"

' Operational columns
Public Const TIS_COL_STATUS     As String = "Status"
Public Const TIS_COL_LOCK       As String = "Lock?"
Public Const TIS_COL_HEALTH     As String = "Health"
Public Const TIS_COL_WHATIF     As String = "WhatIf"

' Fill/border colors for change tracking
Public Const CLR_CHANGE_FILL    As Long = 42495     ' RGB(255, 165, 0) Orange
Public Const CLR_NEW_DATE_BORDER As Long = 16711680 ' RGB(0, 0, 255) Blue

' Health status CF colors
Public Const STATUS_ONTRACK_FG  As Long = 1409045   ' RGB(21, 128, 61) dark green
Public Const STATUS_ONTRACK_BG  As Long = 15204060  ' RGB(220, 252, 231) light green
Public Const STATUS_ATRISK_FG   As Long = 520097    ' RGB(161, 98, 7) dark amber
Public Const STATUS_ATRISK_BG   As Long = 13107198  ' RGB(254, 243, 199) light amber
Public Const STATUS_BEHIND_FG   As Long = 1842617   ' RGB(185, 28, 28) dark red
Public Const STATUS_BEHIND_BG   As Long = 14869246  ' RGB(254, 226, 226) light red

' Zone header colors (defined but not yet fully applied)
' ZONE_IDENTITY_BG/FG, ZONE_TIS_BG/FG, ZONE_USER_BG/FG, ZONE_CALC_BG/FG
' (placeholder constants — to be added when zone styling is implemented)
```

---

## [TECH] Status Filter — Mandatory Locations

Every one of these must filter by `Status`. Cancelled, Completed, and Non IQ systems must never inflate live counts (except the Completed KPI card which specifically counts Completed).

| Location | Module | Filter |
|---|---|---|
| KPI card formulas (Total, New, Reused, Demo, CT Miss, Escalated, Watched, Conversions) | DashboardBuilder | Active + On Hold only |
| Completed KPI card | DashboardBuilder | Status = "Completed" specifically |
| System Counters | DashboardBuilder | Active + On Hold only |
| DashHelper table row generation | DashboardBuilder | Excludes Cancelled, Completed, Non IQ |
| Escalation Tracker table | DashboardBuilder | Excludes Cancelled, Completed, Non IQ |
| HC Need/Available SUMPRODUCT | NIF_Builder | Excludes Cancelled, Completed, Non IQ |
| RampAlignment row collection | RampAlignment | Active only |
| Gantt formula | GanttBuilder | `IF(Status="Active", formula, "")` |
| UpdateWorkingSheetFromTIS cancel logic | TISLoader | Only cancels Active rows |

---

## [TECH] WhatIf Mode — Implementation Contract

**Backup sheet `WhatIf_Backup`** (xlSheetVeryHidden):
- Created by `ActivateWhatIfMode` — stores original Our Date values
- Deleted by `DeactivateWhatIfMode` — after restoring values
- Its existence signals WhatIf mode is active (checked by `ToggleWhatIf`)

**ActivateWhatIfMode algorithm:**
```
1. Guard: if WhatIf_Backup exists, call DeactivateWhatIfMode first
2. Build column map from header row
3. Find WhatIf column, Our Date columns (8), Status column
4. Create hidden WhatIf_Backup sheet
5. Backup: copy Our Date columns (header + data) to backup sheet
6. For each data row:
   a. Skip non-Active rows
   b. Read WhatIf date
   c. If WhatIf date is valid:
      - Find project start (first non-null Our Date: Set, SL1, ..., MRCL.F)
      - Compute delta = WhatIf date - project start date
      - Shift all 8 Our Date cells by delta (directly overwrite)
   d. If no WhatIf date: leave row unchanged
7. Rebuild Gantt (reads shifted Our Dates)
8. Recompute Health
9. Re-protect sheet
```

**DeactivateWhatIfMode algorithm:**
```
1. Find WhatIf_Backup sheet
2. Build column map for both sheets
3. Restore Our Date values from backup -> Working Sheet
4. Delete WhatIf_Backup sheet
5. Rebuild Gantt (reads restored Our Dates)
6. Recompute Health
7. Re-protect sheet
```

**Project start detection (WhatIf):**
- Check Our Date columns in order: Set, SL1, SL2, SQ, Conv.S, Conv.F, MRCL.S, MRCL.F
- Return the first non-null date value
- If no dates populated, skip that row (no shift)
- Non-Active rows are skipped entirely

---

## [TECH] GanttBuilder Our Date Redirect

GanttBuilder uses a redirect mechanism to read Our Dates instead of TIS dates for install/reused milestones:

```vba
' Map TIS source headers to Our Date headers
ourDateRedirect(LCase(TIS_SRC_SET)) = LCase(TIS_COL_OUR_SET)
ourDateRedirect(LCase(TIS_SRC_SL1)) = LCase(TIS_COL_OUR_SL1)
' ... (all 8 pairs)

' For each redirect: if Our Date column found in header map,
' override the phase column index to use Our Date position
For Each rdKey In ourDateRedirect.Keys
    ourKey = ourDateRedirect(rdKey)
    If headerMap.exists(ourKey) Then
        ' Override TIS column -> Our Date column position
    End If
Next
```

Milestones without Our Date equivalents (PF, DC, DM, SDD) fall through and use TIS date columns.

---

## [TECH] Health Computation Contract (Live Formula)

Health is now a **live Excel formula** (not VBA-computed). It auto-recalculates whenever an Our Date cell is edited — no rebuild needed.

**Formula logic per row:**
```
If Event Type = "Demo": Health = "" (blank)
Else:
    maxDev = MAX(Our Date - TIS Date) across all 8 milestone pairs (days)
    (Only pairs where both dates are present are considered)
    If no date pairs found: Health = ""
    ElseIf maxDev > 3: Health = "Gap"
    ElseIf maxDev > 0: Health = "Minor"
    Else: Health = "Match"
```

**Thresholds:** Match (all <= 0), Minor (1-3 days), Gap (> 3 days).

Note: Health only considers positive deviations (Our Date later than TIS). Negative deviations (ahead of TIS) are treated as matching.

---

## [TECH] TIS Change Comment Format

```
New comment (no existing):     "[YYYY-MM-DD] Changed from: [old value]"
Appended to existing comment:  existing_text & Chr(10) & "[YYYY-MM-DD] Changed from: [old value]"
```

Check before writing: `If cell.Comment Is Nothing Then cell.AddComment text Else cell.Comment.Text cell.Comment.Text & vbLf & text`

---

## [TECH] Schema Detection

WorkfileBuilder uses a sentinel header to determine whether to offer CreateWorkingSheet (full build) or report that the schema is already current:

```vba
' Check for "Conv.S" header in existing Working Sheet
If FindHeaderCol(oldSheet, headerRow, TIS_COL_OUR_CONVS, lastCol) > 0 Then
    hasRev14Schema = True  ' Schema is current — weekly updates via TISLoader
Else
    hasRev14Schema = False ' -> CreateWorkingSheet (full build needed)
End If
```

This means any Working Sheet with a `Conv.S` column will be treated as Rev14 schema. Pre-Rev14 sheets (which have `Conversion Start` but not `Conv.S`) will trigger a full rebuild. Weekly TIS updates go through `TISLoader.UpdateWorkingSheetFromTIS` and do not use schema detection — they assume the Working Sheet already has Rev14 schema.

---

## [TECH] Sheet Protection — REMOVED

Sheet protection has been removed from all modules. `FormatConditions.Add` with `xlExpression` fails silently on protected sheets in some Excel 365 builds, causing Gantt CF to not display.

Lock? enforcement uses **Data Validation** as the primary mechanism. WorkfileBuilder Cleanup still creates named ranges (`OUR_DATE_START`, `OUR_DATE_END`, `LOCK_COL`) for the optional injected `Worksheet_Change` handler.

---

## [TECH] VBA Gotchas

### Language
- **`Dim` inside loops** — VBA hoists all Dims to procedure level. Variables declared inside a loop do not reset each iteration.
- **Static arrays in `Collection.Add`** — Use `Array(a, b, c)` not a reused `Dim arr()`.
- **`On Error Resume Next` scope** — Persists until explicitly reset.
- **`Range.Value` bulk read** — 1-based 2D Variant array. Single-cell reads return scalar.

### Excel Object Model
- **PivotTable date grouping** — Use text `YYYY-MM` for month fields.
- **No spill / dynamic arrays** — Use `IFERROR(VLOOKUP(...), "")`.
- **`Formula` vs `Formula2`** — Try `Formula` first, fall back to `Formula2`.
- **Slicer cleanup order** — Delete slicer caches BEFORE ListObjects BEFORE cells.
- **`SlicerCaches.Add2`** — Required for Excel 365.
- **Multi-line column headers** — Strip `vbLf`/`vbCr` when comparing headers.
- **Sort by color requires `Interior.Color`** — CF backgrounds are not sortable. Use hardcoded fill for orange change indicators and blue new-system indicators.
- **`AddComment` on existing comment** — Check `cell.Comment Is Nothing` first.
- **Named ranges** — Use `Application.Range("name")` for workbook-scoped names.
- **PivotItem hide sequence** — Two-phase: all Visible=True first, then hide targets.
- **Adding a sheet changes ActiveSheet** — Even with `ScreenUpdating=False`. Re-activate the intended sheet after `Sheets.Add`.

### Performance
- **Bulk I/O always** — `Range.Value` is O(1). Cell-by-cell is O(n).
- **Cache `UsedRange`** — Call once, store in variable.
- **Create Dictionary once** — Outside loops.
- **Comments are fast** — ~1500 operations with SetPerformanceMode = ~1-2 seconds. Acceptable.

---

## [TECH] Architecture Decisions

| Decision | Why |
|----------|-----|
| Our Dates are values not formulas | User owns these. Formulas could be overwritten. Values survive any rebuild. |
| No SDD in Our Dates | SDD is a customer delivery date — it does not change with installation planning. TIS SDD column is sufficient. |
| No Effective Date columns (unlike Rev13 design) | Simplicity. Gantt reads Our Dates via redirect. WhatIf uses backup-restore instead of parallel columns. Fewer columns = less schema complexity. |
| WhatIf backup-restore approach | Simpler than hidden Effective Date columns. Our Dates are temporarily shifted, then fully restored. Trade-off: WhatIf mode is not visible in the sheet data during activation (just in the Gantt output). |
| Update-in-place (UpdateWorkingSheetFromTIS via TISLoader) | Weekly TIS updates typically change a handful of cells. Full rebuild is wasteful. In-place update touches only changed cells, preserving all formatting, comments, and user data. |
| Schema detection via Conv.S sentinel | Conv.S is a Rev14-specific Our Date header that does not exist in pre-Rev14 Working Sheets. Simple, reliable detection. |
| Status dropdown (5 values) instead of Active? boolean | Richer lifecycle tracking. "Non IQ" added for systems not in IQ scope. Completed column merged in — one dropdown replaces two columns. |
| Orange `Interior.Color` (not CF) for TIS changes | CF backgrounds not exposed to Excel Sort-by-Color. Hardcoded fill enables user to sort all changed rows to the top. |
| Blue border for auto-populated Our Dates | Visual distinction: "this was auto-filled from TIS, not hand-entered — please verify." |
| Appended comments for field history | Lightweight audit trail per cell. No extra sheet, no schema change. Full history accumulates in the comment. |
| Group from Entity Type (not CEID) | Entity Type maps to the operational group. CEID is a sub-variant. CEID changes -> STD durations change, Group stays stable. |
| Removed Systems sheet retired | Status=Cancelled keeps all history on the Working Sheet. One less sheet to maintain. Simpler schema. |
| Gantt reads Our Dates via redirect (not TIS dates) | Users' committed schedule drives the Gantt. TIS dates are reference only. Redirect mechanism in GanttBuilder swaps TIS header lookups to Our Date column positions. |
| Hidden DashHelper sheet | PivotChart source data derived from Working Sheet — avoids modifying Working Sheet structure. |
| CF over VBA painting (Gantt) | CF persists when data changes and survives Undo. |
| In-place Working Sheet rebuild | Preserves worksheet object identity for cross-sheet formula survival. |
| Backup to Old YYYY-MM-DD | User can review/recover. Appends time if same-day. Full build path only. |

---

## [TECH] Named Range Contracts

| Name | Scope | Created By | Used By | Purpose |
|------|-------|------------|---------|---------|
| `DASH_KPI_GROUP` | Workbook | DashboardBuilder | Worksheet_Change handler | Group filter on Dashboard |
| `DASH_CHART_START` | Workbook | DashboardBuilder | Worksheet_Change handler | Chart date filter on Dashboard |
| `IB_BASELINE` | Workbook | DashboardBuilder | GETPIVOTDATA formulas | User-entered starting tool count for Install Base |
| `IB_PT_ANCHOR` | Workbook | DashboardBuilder | GETPIVOTDATA formulas | PT_InstallBase PivotTable anchor for formulas |
| `OUR_DATE_START` | Workbook | WorkfileBuilder Cleanup | Worksheet_Change Lock? handler | Column number of first Our Date (Set) |
| `OUR_DATE_END` | Workbook | WorkfileBuilder Cleanup | Worksheet_Change Lock? handler | Column number of last Our Date (MRCL.F) |
| `LOCK_COL` | Workbook | WorkfileBuilder Cleanup | Worksheet_Change Lock? handler | Column number of Lock? |

Note: Rev14 does not use a `WHATIF_MODE_ACTIVE` named range. WhatIf state is determined by the existence of the `WhatIf_Backup` sheet.

---

## [TECH] Cross-Module String Contracts

**Stable (never change with version bumps):**
- HC table titles (NIF_Builder section above). Dashboard finds by exact text.
- Stable VB_Names: `WorkfileBuilder`, `GanttBuilder`, `NIF_Builder`, `DashboardBuilder`, `RampAlignment`, `HCHeatmap`, `TISLoader`
- `"TISLoader.UpdateWorkingSheetFromTIS"` — called as Step 4 of TIS load flow
- `"HCHeatmap.ToggleHCHeatmap"` — OnAction string. No Rev suffix.
- `"TIS_Launcher.ToggleWhatIf"` — OnAction for WhatIf button on Working Sheet.
- `"DashboardBuilder.RefreshChartStartDate"` — injected into Dashboard sheet code.

**Rev14 additions — all defined in TISCommon:**
- `TIS_COL_OUR_*` constants (8 Our Date headers)
- `TIS_SRC_*` constants (9 TIS source column names)
- `TIS_COL_STATUS`, `TIS_COL_LOCK`, `TIS_COL_HEALTH`, `TIS_COL_WHATIF`
- `CLR_CHANGE_FILL`, `CLR_NEW_DATE_BORDER`
- `STATUS_ONTRACK_*`, `STATUS_ATRISK_*`, `STATUS_BEHIND_*`
- Never use string literals for these in module code.

**TIS_VERSION:**
- `Public Const TIS_VERSION As String = "Rev14"` — informational only. Never use for runtime dispatch.

---

## [TECH] VBA Environment

| Property | Value |
|----------|-------|
| Excel version | Microsoft 365 / Excel 2021+ |
| VBA references | Standard Excel Object Model only |
| Binding | Late binding for Dictionary; early binding for Excel objects |
| `Option Explicit` | Required in all modules |
| `DEBUG_MODE` | `#Const` in TISCommon — `False` production, `True` dev |
| Data start row | Row 15 (`TIS_DATA_START_ROW = 15`) |
| Header row | Row 14 on Working Sheet (found by `FindHeaderRow` scan) |
| Table structure | Working Sheet uses `ListObject` |
| Max tested scale | ~500 systems |
