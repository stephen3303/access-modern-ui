# Adoption plan: every Access app on this kit

**Goal.** SABER and Submarine Program Manager (SPM) both build on `access-modern-ui`.
Nothing in either app redefines what the kit already provides. SPM's best table, card
and pill work moves into the kit first, so both apps can use it.

This plan is written for an agent to complete one task at a time. Each task names what
to read, what to change and when it is done. Work the phases in order: phase 1 changes
the kit, and phases 2 and 3 copy the result into the apps.

Source: design system audit of 12 Oct 2026 (kit 75fda1d, SABER 95a3549, SPM 7785cf6).

---

## Rules for every task

- Follow `docs/DESIGN_GUIDE.md`. In particular, colors come from `CLR_*` roles or `FIXED`
  constants, never `RGB()` or a number; runtime colors are wrapped in `Th()`; sizes are in
  `px()`; and every entry point uses the standard error handler.
- Kit files are copied into an app **unchanged**. An app-specific need goes in an app
  module or a hook the kit calls, never in an edited copy of a kit file. To verify, diff
  each app's copy against `kit/` (line endings aside). The only allowed difference is the
  header comment line.
- Source files stay ASCII. Write any other character as `ChrW(n)`.
- Access can't run in a cloud session. Each task that changes a form ends with a build
  and click-through on Stephen's machine through Remote Control. Check for modal Visual
  Basic error dialogs as well as the error log, and click through in one light and one
  dark palette (`docs/DESIGN_GUIDE.md` §8.9).
- SPM's `AGENTS.md` forbids opening the SABER folder on Stephen's machine. Work on SPM
  reads from this kit repo and the SPM repo only.
- One task per PR unless the task says otherwise. Tick the box here in the same PR.

---

## Phase 1: Harvest SPM features into the kit

SPM built these in PowerShell (`current/source/tools/form-helpers.ps1`,
`ui-current.ps1`) and VBA (`modPills.bas`, `responsive.bas`, `presentation.bas`,
`demo-core.bas`). Its colors are hard-coded teal/navy, so each item is re-expressed
with kit tokens, not copied.

### K1. Rounded cards
- [x] **Source.** SPM `ui-current.ps1` draws `card-light.png`/`card-dark.png` (a 22px
  radius rounded rectangle with a 2px border). `PaintForm` swaps them for controls tagged
  `StaticCard`. `form-helpers.ps1 Box` uses a disabled `Shape = 1` button as a rounded
  surface when it is 650 twips or taller.
- **Change.**
  - Add `CardPng(fill, border, wPx, hPx, radiusPx)` to `kit/modules/modPng.bas`. Draw it
    anti-aliased at 2x and cache it per color and size, like `PillPng`.
  - Add a `Card nm, title, x, y, w, h, [parentCtl]` builder to `modUIKit`. It draws an
    image whose `ControlSource` is `=CardPng(CLR_SURFACE, CLR_BORDER, ...)`, so
    `ThemeForm` repaints it like a pill. On top go a 10pt 600 title and the dark-only
    accent bar. The guide lists `Card` as app-specific today; it becomes a kit builder.
  - Keep the square `Box` card as a fallback for subform hosts, where images behind
    subforms flicker.
- **Done when** a test page shows a rounded card in every palette with no repaint flicker
  on hover. SPM's todo list reports flicker on cards, so check it specifically. The
  README catalog and guide §6.6 describe `Card` and `CardPng`.
- **As built.** One image of the whole card would distort its corners when the card
  stretches with the window, and would be megabytes per size. So the card is a square
  `Box` with four small corner images, `CardPng(fill, border, back, rPx, corner)`, each
  anchored to its corner. `back` is the color the card sits on. Shown on the demo's
  Cards page (`examples/demo`).

### K2. Progress capsule
- [x] **Source.** SPM `modPills.ProgressPng` and `WriteCapsule(..., fraction)` draw a
  capsule filled to a percentage over a gray track.
- **Change.** Add `ProgressPng(percent, fill, track, wPx, hPx)` to `modPng`, plus a
  `Progress nm, expr, x, y, w, h` builder (bound image,
  `=ProgressPng(expr, CLR_PRIMARY, CLR_BORDER, w, h)`). Remove SPM's hard-coded
  `RGB(0,145,165)` and `225,232,239` track. Clamp to 0–100 and treat Null as 0. Guide
  §6.6 "Progress bar" changes from two square boxes to this.
- **Done when** 0, 37 and 100 percent render correctly in light and dark.

### K3. Table upgrades in `BuildGrid` / `GridHeaders`
- [x] **Source.** SPM `form-helpers.ps1 ModernList` and `TableHeader`,
  `responsive.TableSize`.
- **Change.** Extend the column spec and builder. Existing callers must work unchanged.
  - **Header band:** `GridHeaders ..., [band:=True]` draws a filled `CLR_INSET` strip
    (SPM: 460 twips, soft fill) behind the captions.
  - **Zebra rows:** optional `zebra` argument. It sets `AlternateBackColor` to `CLR_ROW2`
    and makes `txtBg` transparent except when the row is selected.
  - **Pill column:** column option `p`, as in `"status|Status|110|p"`. It builds a bound
    pill (image `_bg` plus label) and calls the app hook `StatusVariant(value)` to map
    the value to a badge variant, then `BadgeCF`/`RoundCF`. This replaces SPM's
    `StatusFill` RGB table. Document `StatusVariant` in the kit/app contract.
  - **Row height:** optional `rowPx`. SPM uses 38px (570 twips) and 56px (840 twips) rows;
    the kit derives height from `rowLines`. The default stays the current formula.
  - **Proportional fit:** add a runtime helper `GridFit gridSubformCtl`, called from
    the host's `Form_Resize`, that scales every column in proportion to the new width.
    Today only the last column stretches. Keep `last column stretches` as the default
    and use proportional fit when the host asks for it.
  - **Row click opens:** no arrow buttons. SPM's todo list asks for row click to replace
    the `>` button, and `subGrid.RowClick` already does this, so document it as the rule.
- **Done when** an example page shows each option. `subGrid.vba` stays shared by every
  grid, and the README catalog and guide §6.2 describe every option.
- **As built.** Proportional fit is chosen at build time with `BuildGrid ..., fit:=True`
  (the last column then keeps its own width), and the host calls `GridFit sf, prefix`
  from `Form_Resize`. Columns never shrink below their design width. Shown on the
  demo's Tables page.

### K4. List with detail panel
- [x] **Source.** SPM `presentation.ShowDetail`, `SelectFirstDetail` and the
  `tableCard`/`detailCard`/`subDetail`/`emptyDetail` layout in `responsive.V5Resize`.
- **Change.** Add a page pattern to the guide (§6.5 or a new §6.8) and a runtime helper
  `ShowDetail hostForm, detailCtl, formName, where`. It saves a dirty detail first, swaps
  the `SourceObject` through the deferred-timer rule, filters it and hides the empty
  state. Add `SelectFirstRow grid` to open the first record when the list changes. The
  layout is a grid card on the left that fills the width, plus a fixed 360px detail card
  on the right, anchored.
- **Done when** the example app has one list/detail page that works through palette
  switching and window resize.
- **As built.** `ShowDetail host, detailCtl, formName, where, [emptyCtl]`,
  `ClearDetail` and `SelectFirstRow` are in the new kit module `modKitUI`; guide §6.8
  describes the page. `ShowDetail` only swaps the panel's form when it changes, and must
  not be called from inside the detail form. Shown on the demo's List and detail page.

### K5. Print and export the current grid
- [x] **Source.** SPM `presentation.BuildPrintPreview` and `demo-core.ExportCurrentList`.
- **Change.** Add a new kit module `modGridOut` with:
  - `PrintGrid gridForm, title, subtitle`: builds a temporary report from the grid's
    visible columns, captions and current filter with `NewReport` and `PRINT_*` colors,
    in landscape when the columns are wider than portrait.
  - `ExportGrid gridForm, path`: writes the same rows with the current filter as CSV
    through `modCsv`, or as xlsx through `TransferSpreadsheet` when asked.
  - Both raise through `RaiseAgain`, and neither leaves temp queries behind.
- **Done when** printing and exporting a filtered grid gives exactly the filtered rows,
  with the on-screen captions.
- **As built.** `ExportGrid` picks Excel when the path ends in `.xlsx`. The report is
  kept as `zz_rptGrid` and replaced each time; the Excel export's temporary query
  `zz_qryGridExport` is deleted afterwards.

### K6. Filter chips and tabs with counts
- [x] **Source.** SPM `program-workflow.V7Counts` and Stephen's SPM todo list ("Filter
  buttons should include counts", "move the filters and search bars to the right of the
  filter buttons").
- **Change.** `TabStrip` and `StyleChip` take an optional count, shown as `Name (n)`.
  Add `SetCount ctl, n`, which keeps the base caption in `Tag`. Guide §6.6 shows the
  layout: chips on the left, then search and filters to their right, on one row.
- **Done when** it is documented and shown in the example.
- **As built.** `TabStrip` takes a count as a third part of a tab, `"key|Caption|0"`.
  `StyleChip lbl, isOn, [n]` and `SetCount` are kit code now (modKitUI).

### K7. Kit housekeeping (from the audit)
- [x] `tools/make_themes.py` creates `modules/` and `themes/light/` when they are
  missing. Today it crashes on a fresh app.
- [x] Guide references to `kit/README.md` point at the root `README.md`, or the catalog
  moves to `kit/README.md`. Pick one. **Picked:** the catalog stays in the root
  `README.md`, and the guide points there.
- [x] Promote `SetVisible c, v` from SABER (`access/src/modules/modUI.bas`) into
  `kit/modules/modKitUtil.bas`. It moves focus to `txtFocusSink` before hiding the
  focused control, which avoids error 2165. The guide (§3 module table, §6.7) already
  relies on it, but the kit doesn't ship it. In the same PR, move it from the `modUI` row
  to `modKitUtil` in the guide's §3 module table and add it to the README catalog. Remove
  SABER's copy in S1 and use the kit one in SPM.
- [x] Guide §6.4 documents a **compact** size (28px button, 26px input) for toolbars and
  grid rows. SABER uses it heavily without it being documented.
- [x] Add a teal/navy palette matching SPM's current look (sidebar 10,42,73, primary
  0,132,151) to `PALETTES`, so SPM users can keep the look they know. It must pass the
  contrast check. **As built:** "Teal Navy" (`tealnavy`). SPM's teal gives white text
  4.4:1, so the palette uses 0,128,147, which passes.

---

## Phase 2: SABER on the current kit

SABER copied an older kit (light/dark only). Do these after phase 1 so it picks up the
new features in one pass.

### S1. Refresh kit copies
- [ ] Copy every file in `kit/modules`, `kit/forms` and `tools/make_themes.py` over
  SABER's copies. Move anything SABER-specific out first:
  - `Banner()` from `modUIKit` goes into an app module (`modBuildUI`). The kit removed
    the classification banner in 75fda1d. Ask Stephen whether SABER keeps it. **Default:
    keep it as app code.**
  - Hard-coded `"SABER"` paths in `modPng` and `modDarkMode` are replaced by `APP_NAME`.
  - `SetVisible` and `StyleChip` are kit code now (modKitUtil, modKitUI): delete them
    from `modUI`. `StyleChip`'s new optional count argument keeps old calls working.
  - `StatusVariant` takes a Variant (pill columns pass Null for empty values).
  - Copy the new kit modules too: `modKitUI`, `modGridOut`.
- [ ] Re-run `make_themes.py`, which now generates 7 palettes plus the contrast check.
- **Done when** the diffs against `kit/` show header lines only, and the build passes.

### S2. Palette picker
- [ ] Replace the sidebar dark-mode toggle (`ToggleDarkMode`) with Appearance →
  `dlgAppearance` (`ShowAppearance`). Copy `kit/forms/dlgAppearance.vba` and
  `examples/modules/modBuildAppearance.bas`. Saved `light`/`dark` choices must still
  load.

### S3. No color literals
- [ ] `modUI.BandBack` uses `CLR_BAND_*`.
- [ ] Classification colors become `FIXED` entries in `make_themes.py`
  (`CLR_CLASS_SECRET`, and so on) and are used from `modGlobals.ClassificationColor`.
- [ ] `modBuildPages2.bas:220` uses `CLR_BAND_NE`, and `:780` uses `CLR_WHITE` or
  `CLR_SURFACE`.
- [ ] Remove unused `CLR_LINK` and `CLR_KPI6`, or use them.
- **Done when** `grep -rn "RGB(" access/src --include=*.bas --include=*.vba` only finds
  generated files.

### S4. Tables and cards to the new style
- [ ] Rebuild SABER's grids and list subforms with the K3 options, and cards with K1. This
  covers Stephen's SABER task "Adjust remaining tables in UI to match the style" and the
  Find Past Comments dialog changes in `Claude Task List.txt`.
- [ ] Button heights become 32px standard or 28px compact (K7). Remove the 24/26/30 one-offs.
- [ ] Sidebar subtitle and login name go from 7pt to 8pt. `lblCompletion` goes from 14pt
  to 15pt (page title) or 13pt.

---

## Phase 3: Rebuild SPM on the kit

SPM has no kit code today. Its forms are made by PowerShell scripts through COM and kept
as exported `.access.txt` objects. Its theming is a hand-written `PaintForm` with RGB
literals. The rebuild moves it to the kit's model: VBA `Build_*` procedures, a
generated palette and `ThemeForm`. Keep SPM's behavior and data. Change only how it is
built and styled.

### P1. Scaffolding
- [ ] Add `access/` in the kit layout: `kit/` (copied), `tools/make_themes.py`,
  `build.ps1` (from `examples/build.ps1`), `src/modules`, `src/forms`,
  `src/themes/light`, `src/help/help.txt`, `src/seed`. Set `APP_NAME = "Submarine
  Program Manager"`. Convert `user-seed-data.json` into the seed step. Seed data
  Stephen asked to keep (departments, divisions, program roles) must survive.
- [ ] `modErrors` from `examples/modules/modErrors.bas` replaces `HandleUIError` /
  `HandleDataError`. Keep the `ErrorLog` table's data.
- [ ] `modGlobals`: `APP_NAME`, `g_Silent`, `Notify` (replaces `V6Message`), `Q`/`SqlVal`.
- [ ] `BuildAll` builds an empty shell and the build passes on Stephen's machine.

### P2. Shell
- [ ] `frmMain` replaces `frmHome`: the kit sidebar (collapsible, deferred `NavTo`) holds
  the command logo with command text below it, at full window height and with no top
  title bar, as Stephen's SPM todo list asks. Add the Appearance palette picker, and add
  Help and the error log under Admin. This removes `V6ToggleNav`, `V6ToggleTheme`,
  `V6PaintTheme`, `PaintForm` and `V5Resize`.

### P3. Pages, one task each
Each becomes `pg<Name>` built with `PageHeader`, chips with counts (K6), `BuildGrid` (K3)
and the list/detail pattern (K4), with print and export through K5. Remove the "# items
in this view" text, the extra white space under titles and the demonstration notes, as
the SPM todo list asks.
- [ ] P3.1 Overview / home (KPI tiles; replaces `statCard*`)
- [ ] P3.2 Gains (`pgGains`, `frmGainWorkspace`, `v3GainList`, `v4Gains`)
- [ ] P3.3 Active (`pgActive`, `frmActiveWorkspace`, `v3ActiveList`, `v4Active`)
- [ ] P3.4 Losses (`pgLosses`, `frmLossWorkspace`, `v3LossList`)
- [ ] P3.5 Programs and coverage (`pgPrograms`, `v4Programs`, `v4Coverage`, `frmCoverage`)
- [ ] P3.6 Documents and verification (`frmVerification`, `v4Documents`, `v4DocHistory`)
- [ ] P3.7 Archive (`frmArchive*`, `v3ArchiveList`, `v4Archive`)
- [ ] P3.8 Personnel (`pgPersonnel`, `frmPersonEditor`)
- [ ] P3.9 Reports (`pgReports`, `rptDemoPrint` replaced by K5 `PrintGrid`)
- [ ] P3.10 Admin (`pgAdmin`, `v6AuditLog`, `v6DepartmentChoices`, `v6DivisionChoices`,
  `frmErrorLog`)

### P4. Editors become dialogs
- [ ] Each `frm*Editor` (Appointment, Assessment, Document, Event, Milestone, POAM,
  POAMMilestone, Program, Person) becomes `dlg<Name>` built with `NewDialog`. Validation
  returns `False` + `errMsg` into `lblMsg` (guide §3), replacing `V7Validate*` message
  boxes. Cancel goes left of the primary action, both bottom-right.

### P5. Naming and cleanup
- [ ] No form keeps a `v3`–`v7` or `frm*` (non-window) prefix. Subforms are `sub*`.
  Procedure prefixes `V4`–`V7` are dropped as each page moves, and procedures are grouped
  into `mod<Area>` modules (guide §3).
- [ ] The old PowerShell form scripts and exported `.access.txt` forms move to `archive/`
  once their replacements build.
- [ ] `current-tests.bas` runs against the new build in silent mode.
- **Done when** SPM has no `RGB(` outside generated files, no Aptos font, and only type
  scale sizes (guide §6.3).

---

## Progress

| Phase | Tasks | Done |
|---|---|---|
| 1 Kit harvest | K1–K7 | 7 (demo tested in Access 2026-10-12; card corners, hover flicker and the zebra selected row still need a look on screen) |
| 2 SABER | S1–S4 | 0 |
| 3 SPM | P1–P5 | 0 |
