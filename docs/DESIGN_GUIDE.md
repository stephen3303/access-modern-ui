# Design guide: modern UI in Microsoft Access

Read this before designing or changing any Microsoft Access front end in this style. It
covers module layout, documentation, error handling, form design and color palettes. It
covers no specific app's tables or business rules, so you can apply it to any use case.

The kit files it describes are in this repository (`kit/`, `tools/`, `examples/`). Paths
below such as `access/src` describe the layout of an app built with the kit. The kit's API
catalog and "What the app must provide" table are in this repository's root
[README.md](../README.md); apps link to it rather than copying it. `examples/demo` is a small
app that shows every builder working.

---

## 1. Design philosophy

- **The Access file is a build output.** Every form, report, table, query, macro and module
  is generated from plain-text source by a build script. Nothing is edited in the Access
  designer, because the next build would overwrite it. Source lives in git, and the
  `.accdb` files can be thrown away.
- **Port a web design system, don't accept Access defaults.** The look comes from a
  Tailwind/shadcn-style web app: white surfaces, slate grays, one blue accent, rounded
  pills, 600-weight headings, muted uppercase captions, generous padding. Access chrome
  (record selectors, navigation buttons, dividing lines, datasheet view, the nav pane,
  document tabs) is turned off everywhere.
- **Measure in CSS pixels.** Layouts are written in px and converted with
  `px(n) = n * 15` twips (96 dpi). That lets you copy a web mockup's numbers directly.
- **Use tokens, never literals.** Every color is a named role (`CLR_*`). No `RGB()` or
  numeric color appears in VBA. Palettes are generated from one Python file.
- **Split front end and back end.** Each user runs their own front-end copy (forms, code,
  per-user scratch tables) linked to one shared back-end data file. Version checks keep the
  two in step.
- **Keep the kit separate from the app.** Reusable modules go in `kit/`. App modules go in
  `src/` and provide a few named hooks the kit calls.

---

## 2. Source layout and build

```
access/
  build.ps1                  drives Access via COM; loads modules, calls BuildAll
  kit/
    modules/*.bas            reusable: modUIKit, modKitUI, modDarkMode, modPng, modGridOut,
                             modCsv, modCrypto, modKitUtil (copied unchanged from the kit)
    forms/*.vba              kit code-behind: subGrid, dlgAppearance, dlgDatePick
  src/
    modules/*.bas            the app's own standard modules (incl. generated modPalette)
    themes/light/modTheme.bas   build palette constants (generated)
    forms/<FormName>.vba     code-behind, attached when that form is built
    help/help.txt            user help topics (loaded into a local table at build)
    assets/*                 files stored in a local table at build (e.g. helper scripts)
    seed/*.csv               test data
  tools/make_themes.py       generates modTheme.bas and modPalette.bas, checks contrast
  DEVELOPING.md              maintainer guide
```

**build.ps1**
- Uses the `Access.Application` COM object with `Visible = $false` and
  `AutomationSecurity = 1`.
- Modes:
  - Full build: recreates the front end and the back end and loads seed data, which
    replaces all data.
  - `-FrontEndOnly`: a new front end linked to the existing back end. Use this for
    everyday UI and code work.
  - `-UiOnly`: rebuilds forms and reports inside the existing front end.
- Load order: `kit/modules`, then `src/modules`, then `modTheme.bas`. Each file is
  normalized to CRLF and **ASCII**, then loaded with `LoadFromText(5, ...)`.
- It compiles everything first (`RunCommand 126`), so syntax errors stop the build before
  any work runs.
- It calls `BuildAll(srcRoot, outDir, uiOnly, frontEndOnly)` through `Application.Run`.
  `BuildAll` returns `"OK"` or `"ERROR during <step>: <num> <desc> [<last Trace>]"`, and
  the script exits 1 on failure.

**BuildAll**
- Sets `g_Silent = True` and calls `UseBuildPalette`.
- Runs named steps, setting `m_step = "..."` before each so a failure names the step:
  create back end, link, local tables, seed, assets, user interface, help, startup.
- Startup options it sets:
  - StartUpForm = the login form.
  - Hidden: the DB window, document tabs and datasheet schema.
  - Name AutoCorrect off.
  - AppTitle.
  - A custom DB property recording the build stamp (`yyyy-mm-dd hh:nn`), which error-log
    entries and version checks use.

**Source files must be plain ASCII.** Write non-ASCII characters in code as `ChrW(n)`, for
example `ChrW(8230)` for an ellipsis, `ChrW(183)` for a middle dot and `ChrW(8592)` for a
left arrow.

**Build order inside BuildUserInterface:** build subforms before the parents that embed
them, and the shell (`frmMain`) last.

---

## 3. Module conventions

**Naming**
- `mod<Area>` for standard modules.
- Forms by role:
  - `frm*`: top-level windows (login, main shell).
  - `pg*`: pages hosted inside the shell.
  - `dlg*`: modal dialogs.
  - `sub*`: subforms, list rows and trees.
  - `rpt*`: reports.
- Local per-user tables use the prefix `zz_`.

**Typical module split for an app in this style**

| Module | Role |
|---|---|
| modGlobals | Constants (`APP_NAME`), session state accessors (TempVars), SQL helpers (`Q`, `SqlVal`, `NzStr`, `NewRow`), `Notify` |
| modErrors | Error log and handlers (section 5) |
| modVersion | FE/BE schema version and minimum front-end build checks at sign-in |
| modHelp | In-app help: loads `help.txt`, picks the topic for the current page, F1 key |
| modAuth, modCrypto | Sign-in, permissions, salted SHA-256 via Windows CNG |
| modSchemaDDL | Table specs and back-end creation/upgrade (raise `SCHEMA_VERSION` + migrate) |
| modBuild | `BuildAll` entry point and startup options |
| modBuildUI, modBuildPages, modBuildPages2 | `Build_<form>` procedures, one per form, grouped by area |
| modUI | Run-time UI helpers: `NavTo`, badges/pills painting, `ShowTip`, the app's hooks (`BadgeColors`, `StatusVariant`) |
| modUIKit (kit) | Build-time control builders (section 6) |
| modKitUI (kit) | Run-time page helpers: `GridFit`, `SelectFirstRow`, `ShowDetail`, `ClearDetail`, `PickDate`, `SetCount`, `StyleChip` |
| modGridOut (kit) | `PrintGrid`, `ExportGrid`: a table as the user sees it, to paper, CSV or Excel |
| modKitUtil (kit) | `FileExists`, `FolderExists`, `EnsureFolder`, `IsFormOpen`, `RunDialog`, `SetVisible`, `SetEnabled` |
| modDarkMode (kit), modPalette (gen), modTheme (gen) | Palette engine (section 7) |
| modPng (kit) | Draws anti-aliased PNGs for rounded pills, card corners and progress capsules |
| modWrap | GDI text measurement (`MeasureText`) and word wrap for fixed-height rows |
| modCsv, modZip (kit / app) | CSV I/O, ZIP |
| Domain modules | Business logic, one module per area; they enforce rules so they hold from any caller |

**Kit/app contract.** Kit modules call names that the app must provide. The kit's
README.md lists them:
- `CLR_*`, `FONT_UI`, `THEME_DARK`
- `PALETTE_COUNT`, `PaletteKey`, `PaletteName`, `PaletteIsDark`, `PaletteSwatch`,
  `PaletteOf`
- `APP_NAME`
- `HandleError`, `LogError`, `RaiseAgain`
- `MeasureText`, `ShowTip`
- `frmMain` with `PageKey` and `Reopen pg`
- `dlgAppearance`
- Optional, only for the builders that use them: `BandLabels`,
  `BandBack`, `BadgeColors`, `StatusVariant(value)` (pill columns), `subGrid.vba`, and
  `GridClick(grid, dbl)` on every form that hosts a table

**Rules live in modules, not forms.** Forms hide buttons the user can't use, but every
rule is checked again in the module function the button calls. A validation failure the
user can fix is not an error: the function returns `False` with a message in a `ByRef
errMsg`, and the form shows it in an inline red label (`lblMsg`, `CLR_ERROR`).

**Session state** lives in `TempVars`. Nothing about the session is written to tables, and
saved queries can read `[TempVars]![X]`. Expose it through accessor functions such as
`CurUserId()` and `IsAdmin()`.

**Silent mode.** `g_Silent` is True during the build and tests. In that mode, `Notify`
prints to the debug window, raises on warnings and answers Yes to questions, and
`HandleError` raises the error again, so nothing can block the invisible Access.

---

## 4. Documentation conventions

Write documentation in three layers, all in plain, direct prose:

1. **Module header comment** at the top of every `.bas`:
   - Line 1: `' <App> (Access edition) - <what this module does>.`
     Kit modules use `' Access UI kit - ...`.
   - Then a short paragraph on the concept and any non-obvious mechanism: why it exists,
     what calls it, and where its state lives.
   - Generated modules say `Generated by access/tools/make_themes.py - edit the palette
     there.`
2. **Procedure comments**: one or two lines above each non-trivial procedure, saying what
   it returns or does and any argument conventions, for example
   `' kind: primary | outline | ghost | destructive | nav` or `' back / fore = -1 keep the
   control's own color.`
   - Explain *why* for Access workarounds, for example "A focused text box paints a square
     over its pill" or "swapping the hosted form while its event is still running is
     unsafe".
   - Don't narrate obvious code.
3. **Markdown docs**:
   - `README.md` (root): what the files are, getting started, using the app, and how to
     rebuild.
   - `access/DEVELOPING.md` (maintainer guide), with these sections:
     - The one rule: edit the source, then rebuild.
     - Build modes table.
     - Source layout.
     - How the build works.
     - Module map table.
     - Key concepts (session state, navigation, dialogs, list subforms, local tables, wrap,
       pills, palettes).
     - Error handling.
     - Where rules are enforced (rule → function table).
     - Common tasks as numbered recipes (change a layout, add a page, change help, add a
       column).
     - Versioning.
     - Access quirks.
     - Testing.
   - The kit's own `README.md` (this repository's root, not copied into apps): module
     table, "Starting a new app with the kit" steps, a "What the app must provide" table,
     "Rules that keep palettes working", and a **builder catalog** (signature → what it
     draws) grouped as Forms, Controls, Page furniture, Tables and conditional colors,
     Layout helpers and Run time.
4. **In-app help** (`src/help/help.txt`), loaded into a local table at build time:
   - The file format is described at the top of the file:
     - `=== key | Title` starts a topic. A dotted key (`page.tab`) is indented under its
       parent.
     - `# Heading`, `- bullet`, `**bold**`.
     - A blank line ends a paragraph.
   - Topic keys match page keys, so Help, the F1 key (an `AutoKeys` macro running
     `=ShowHelp()`) and the sign-in Help link open the topic for whatever is on screen.
   - Name buttons and tabs exactly as they are captioned on screen.
   - The help window is non-modal so it can stay open beside the page. It opens as a
     dialog when another modal form has the focus.
5. **"?" title tips**: each page title and many section titles get a `TitleTip`, a
   Segoe MDL2 `ChrW(&HE9CE)` glyph that shows the explanation on hover and in a message
   on click.

Writing style for all docs: second person, short sentences, active voice, no marketing.
Name UI elements exactly as captioned.

---

## 5. Error handling

All of this lives in `modErrors`, with three patterns.

**Entry points** are event procedures and any function a control calls (`=RowClick()`).
Each ends with the standard handler:

```vba
Private Sub btnSave_Click()
    On Error GoTo Err_Handler
    ...
    Exit Sub
Err_Handler:
    HandleError "frmName.btnSave_Click"
End Sub
```

- `HandleError(source, [context], [intro])` reads `Err`, calls `LogError`, turns off the
  hourglass, and shows a friendly `MsgBox`: "Something went wrong, so that action didn't
  finish." plus the description, the error number and "saved in the error log as entry
  N…".
- In `BeforeUpdate` handlers, also set `Cancel = True`.

**Library procedures** that know useful context add it and pass the error up:

```vba
Fail:
    RaiseAgain "modX.DoThing", "record " & id & ", file " & path
```

`RaiseAgain` builds a call trail, innermost first. The trail is keyed to the current
error's number and description, so a stale trail is never attached to an unrelated error.

**Small helpers** have no handler. Errors bubble to the nearest handler, so one failure
produces one log entry.

**`LogError(source, context)`** never raises:
- It writes a row to an `error_log` table **in the back end**, so admins see every user's
  errors in an Admin › Error Log tab with detail view, CSV export and clear.
- Each row records:
  - Time, error number, description and `Err.Source`.
  - The catching procedure, the call trail and the context.
  - The current page key, active form and active control.
  - App user, Windows user and computer.
  - The current context record id.
  - The front-end build stamp and the Access version.
- If the back end can't be written, it falls back to
  `%LOCALAPPDATA%\<APP>\logs\errors.log`.
- An `m_busy` guard prevents recursion.
- Use `LogError` alone (no message) where a failure must not interrupt the user. For
  example, `ThemeForm` and `ThemeControl` log and carry on, so one bad control can't stop
  a form opening.

**Rules**
- Keep `On Error Resume Next` only for expected, harmless failures (deleting something
  that may not exist, reading a property not every control has, moving focus). Keep the
  block tiny and end it with `On Error GoTo <handler>` or `On Error GoTo 0`.
- Any `On Error` statement resets `Err`, so save `Err.Number` and `Err.Description` before
  calling clean-up code from a handler.
- Expected, user-fixable problems return `False` + `errMsg`. They are never raised.
- The build reports failures with `Trace "<form name>"` (called by `SaveForm`) and
  `m_step`.

---

## 6. Form design system

### 6.1 How forms are built

Every form is drawn in code by a `Build_<name>` procedure:

```vba
Private Sub Build_dlgExample()
    Dim c As Control, x As Long
    NewDialog "Window Title", px(520), px(330)       ' or NewForm recSrc, view, w, h
    Ev F, "OnLoad"                                    ' wire form events to [Event Procedure]
    x = px(24)
    Lbl "lblHeader", "Dialog Heading", x, px(14), px(470), px(26), 13, True
    FieldLbl "lblNameCap", "NAME", x, px(54)          ' 8pt bold muted uppercase caption
    Txt "txtName", "", x, px(72), px(472), px(30)
    Lbl "lblMsg", "", x, px(278), px(250), px(30), 9, False, CLR_ERROR   ' inline validation
    Set c = Btn("btnCancel", "Cancel", px(292), px(278), px(96), px(32), "outline"): c.Cancel = True
    Set c = Btn("btnOK", "Save", px(400), px(278), px(96), px(32)): c.Default = True
    SaveForm "dlgExample"                             ' attaches src\forms\dlgExample.vba
End Sub
```

- `F` is the global form or report under construction.
- `SaveForm` attaches the code-behind, saves, and renames the form to its final name.
- Event wiring:
  - `Btn` sets `OnClick = [Event Procedure]` itself.
  - Other events use `Ev c, "AfterUpdate", ...`.
  - Controls in continuous subforms call public functions instead, for example
    `.OnClick = "=RowClick()"`.
  - `Form_Resize` on a subform only fires if the build sets `Ev F, "OnResize"`.

### 6.2 Builders (modUIKit)

- **Forms**
  - `NewForm recSrc, view(0 single/1 continuous), w, h`
  - `NewDialog caption, w, h`: pop-up, modal, auto-center, dialog border, `CLR_DIALOG`
    background.
  - `NewReport recSrc, w`: Letter portrait, 0.5 in margins, `REPORT_W = 10800`, pop-up
    preview.
  - `SaveForm`, `Trace`
- **Controls**
  - `Lbl`: transparent label.
  - `Txt`: input with `CLR_INPUT_BG`, a 1px `CLR_INPUT_BORDER` border and inner margins.
  - `DateBox`: a date input, 120px wide, with its date left-aligned like text, and a
    30px outline calendar button 4px to its right that opens the kit's date picker
    (`PickDate`, `dlgDatePick`, built by `BuildDatePicker`): a month of day cells with
    previous and next arrows, the box's date highlighted, today in bold, and Today,
    Clear and Cancel. Access's own picker is off: it shows only on focus, and
    `acCmdShowDatePicker` is refused once it is off. The pair is 154px
    (`DATEFIELD_W`), so it fits a half-width field column.
  - `RO`: read-only text, borderless, transparent, no tab stop. Show read-only dates
    left-aligned (`TextAlign = 1`) so they line up with text values.
  - `Cbo`, `Lst`, `Chk`
  - `Btn ... kind`: `primary`, `outline`, `ghost`, `destructive` or `nav`, each with
    themed hover and pressed colors and `Shape = 1` for rounded corners.
  - `Link`: clickable label.
  - `Box`: rectangle; pass -1 for no fill or no border.
  - `Rule`: a 1px hairline (a line control), flat (h = 0) or upright (w = 0). Use it
    for every divider and edge: Access paints a filled 1px `Box` 2 device pixels thick.
  - `RowBack`: row background that paints but never takes focus.
  - `Img`, `SubF`, `TabCtl`
  - `TabStrip x, y, w, handler, "key|Caption", ...`: a tab given as `"key|Caption|n"`
    shows a count, "Caption (n)", that `SetCount` changes at run time.
  - `Card nm, title, x, y, w, h, [parentCtl], [hAnchor], [vAnchor], [back], [square]`:
    a rounded card (section 6.6).
  - `Progress nm, expr, x, y, w, h`: a rounded progress capsule (section 6.6).
- **Page furniture**
  - `PageHeader title, tip`: 15pt bold title at y = 22px, 9pt muted subtitle at
    y = 52px, a "?" tip and a dark-only accent bar the height of the title's text
    (it does not reach the subtitle).
  - `TitleTip`, `AccentBar`/`DarkOnly`, `TitleBar` (an accent bar sized to a title's
    font: cap tops to baseline)
  - `ContextBanner`: inset box with an icon, name, subline and optional pill.
- **Tables**
  - `BuildGrid name, cols, w, [rowLines], [back]`: a styled read-only table. It is a
    continuous form, not a list box.
    - Each column is `"field|caption|widthPx|opts"`. Options: `r` right-aligned,
      `c` centered, `b` bold, `m` muted, `p` a status pill, `w` wraps onto `rowLines`
      lines (the last column always wraps when `rowLines` > 1); text longer than that
      stops at the end of the last line. Never let a cell cut text off mid-word: widen
      the column or make it wrap. A pill column draws the
      value as a rounded badge whose variant comes from the app's
      `StatusVariant(value)` (it gets a Variant, which may be Null, and returns a
      `BadgeColors` variant or "" for no pill). Map statuses there, never as colors.
    - The last column takes the remaining width and stretches with the window. A pill
      in the last column stops 8px further in, so it clears the table's scrollbar.
    - Rows are separated by a 1px line, and the selected row is highlighted in
      `CLR_PRIMARY_SOFT`.
    - Optional arguments, all off by default so existing tables don't change:
      - `zebra:=True` fills every other row with `CLR_ROW2`. The selected row still
        turns `CLR_PRIMARY_SOFT`.
      - `rowPx:=n` sets the row height in px. Use 38 for a roomy one-line table and
        56 for two lines. By default the height comes from `rowLines`.
      - `fit:=True` keeps the last column at its own width so `GridFit` can scale
        every column in proportion. Call `GridFit Me.sfList, "hdrList"` from the
        host's `Form_Resize`; it also moves the captions `GridHeaders` drew with that
        prefix. Columns never get narrower than designed.
    - **Clicking a row opens it.** There are no `>` or Open buttons in rows: the row
      click selects the record and the host's `GridClick` shows it (in a detail panel,
      section 6.8, or a dialog).
    - All grids share one code-behind, `subGrid.vba`, which calls the parent's
      `GridClick(grid, dbl)`.
    - Each column's `Tag` holds its caption and design position, which `GridFit` and
      modGridOut read.
  - `GridHeaders prefix, cols, x, y, [parentCtl], [band], [w]` puts the captions above
    the grid in 8pt bold muted uppercase. `band:=True` draws them on a 22px `CLR_INSET`
    strip, `w` wide (default: the columns' widths), that stretches with the window.
  - `PrintGrid gridForm, title, [subtitle]` and `ExportGrid gridForm, path` (modGridOut)
    print or export a table as the user sees it: its visible columns with their
    captions, and the rows its filter and sort show. `PrintGrid` builds a report
    (`zz_rptGrid`) in `PRINT_*` colors, landscape when the columns are wider than
    portrait. `ExportGrid` writes CSV, or Excel when the path ends in `.xlsx`. A filter
    must name fields without a table prefix.
- **Conditional colors**
  - `CF c, expr, back, fore, [bold]`
  - `BadgeCF c, field, value, variant, ...`
  - `BandCF`
  - `RoundCF`/`PillLabel`: rounded pill images.
- **Layout**
  - `Anchor c, h, v`: 0 = left/top, 1 = right/bottom, 2 = both.
  - `VCenter`: pads the top margin, because Access has no vertical alignment.

Write app-specific helpers next to the builds that use them. Examples:
- `FieldLbl`: an 8pt bold muted uppercase caption above a field.
- `ColHdr`

### 6.3 Type scale (font `FONT_UI = "Segoe UI"`)

| Use | Size / weight | Color |
|---|---|---|
| Page title | 15pt, 600 | `CLR_FG` |
| Dialog heading | 13pt, 600 | `CLR_FG` |
| Section heading | 11pt, 600 (with accent bar) | `CLR_FG` |
| Card title, emphasized row text | 10pt, 600 | `CLR_FG` |
| Body, inputs | 10pt, 400 | `CLR_FG` |
| Secondary text, subtitles, buttons (600) | 9pt | `CLR_MUTED_FG` for secondary |
| Field captions, column headers, badges | 8pt, 600, captions UPPERCASE | `CLR_MUTED_FG` |
| KPI numbers | 24pt, 600 | `CLR_FG` |

- Bold is weight 600, and 700 only for banners.
- Icons are **Segoe MDL2 Assets** glyphs in labels (`c.FontName = "Segoe MDL2 Assets"`,
  `ChrW(&Hxxxx)`). Examples: E700 menu, E80F home, E713 settings, E897 help, E790 color,
  E7E8 sign out, E73E check, E9CE info/"?", E81D context, E8BD comments.

### 6.4 Spacing and sizing (CSS px)

- Page padding: 32px left and right. Dialog padding: 24px.
- Inputs are 30px tall. Buttons are 32px (34px for a full-width primary such as Sign in).
  Captions sit about 18px above their input.
- **Compact** size, for toolbars, filter rows, and actions inside cards, panels and
  table rows: buttons 28px and inputs (search boxes, filter combos) 26px to 28px, text
  still 9pt. Use one size per row. No other heights: no 24, 26 or 30px buttons. The one
  exception is the `DateBox` calendar button, which takes its input's 30px height.
- Vertical rhythm for stacked fields: 38px.
- Cards and KPI tiles are separated by 12px gaps.
- A KPI tile is 78px tall: a surface box with a 1px border, a 4px colored accent strip
  along the bottom, a 10pt caption and a 24pt number. A transparent `Link` over the whole
  tile makes it clickable.
- Dialog buttons sit bottom-right in the order **Cancel (outline)** then **primary or
  destructive** action. Cancel sets `.Cancel = True` and the primary action sets
  `.Default = True`. The inline error label sits bottom-left.
- Header actions sit top-right, anchored right: the primary action rightmost, then
  outline, then destructive.

### 6.5 Application shell

- `frmMain` is a single form containing:
  - A **sidebar** on the left: 200px, collapsible to 60px icons-only, with the collapsed
    state saved in the registry. It contains:
    - The brand icon, name and subtitle.
    - A 1px divider in `CLR_SIDEBAR_LINE`.
    - Nav items. Each is three clickable labels: a highlight background `nb_<key>`, a
      glyph `ni_<key>` and text `nt_<key>`, 34px tall on a 38px step, with a
      `ControlTipText` so collapsed icons still have names.
    - At the bottom, anchored: admin, an avatar circle with the user's name, Appearance,
      and Help + Sign out side by side.
  - A **page host** subform `sfPage` that fills the rest and is anchored both ways. Pages
    are designed at 1056×560px and grow by anchoring.
- **Navigation is deferred.**
  - Pages call `NavTo "key"`, which calls `frmMain.QueueNav`. That sets a pending key and
    `TimerInterval = 10`, and `Form_Timer` swaps `sfPage.SourceObject`. Swapping a
    subform while its own event is still running is unsafe.
  - `ShowPage` redirects keys that aren't allowed in the current state.
  - `RefreshNav` restacks the visible items without gaps and highlights the current one.
- **Reopen**: `frmMain.Reopen pageKey` closes and reopens the shell on the same page,
  which is how a palette switch repaints everything. An `m_reopen` flag keeps `Form_Close`
  from signing the user out.
- **Dialogs**:
  - `RunDialog(name, openArgs)` opens the form with `acDialog` and returns True when the
    dialog set `TempVars("DialogResult") = "ok"`.
  - Other results come back in named TempVars.
  - Dialog input comes through `OpenArgs`.
- **List subforms** (continuous forms of cards or rows) call back to their host with
  `Me.Parent.<PublicProc>`.

### 6.6 Visual components

- **Rounded pills and badges.** Access only draws square corners.
  - Each pill is a transparent text or label control with an **image control
    `<name>_bg` behind it**. The image shows a capsule PNG drawn by `PillPng(color,
    wPx, hPx)` (modPng: CRC and zlib-stored PNG, drawn at 2x, cached per color and size
    in `%LOCALAPPDATA%\<APP>\pills`).
  - Bound pills: `RoundCF` gives the image a
    `ControlSource = "=PillPng(Switch(cond1,color1,...),w,h)"` built from the control's
    conditional formats. Call it **last**, because `CF` makes the control opaque again.
  - Pills painted in code: `PaintBadge` or `PillBack lbl, color`.
  - Badge variants (shadcn): `success`, `warning`, `info`, `destructive`, `default`,
    `outline`, `secondary`. `BadgeColors(variant)` returns `Array(back, fore)`.
- **Cards**: `Card` draws a `CLR_SURFACE` box with a 1px `CLR_BORDER` border, a 10pt 600
  title at (x + 14px, y + 10px) and a dark-only 3px accent bar beside it, the height of the title's text. Its corners
  are rounded (12px): each corner is a small image control showing `CardPng(fill,
  border, back, r, corner)`, laid over the box's square corner. `back` is the color the
  card sits on (default `CLR_BG`), which fills outside the arc. The box itself has no
  border: Access paints a rectangle's right and bottom border a pixel past the corner
  images, which shows as a thin square line. The four straight edges are 1px
  `CLR_BORDER` rules (`box<nm>_t`, `_b`, `_l`, `_r`) that stop at the corner images.
  Anchor the card with `hAnchor`/`vAnchor`; the corners and edges follow, so a card
  that stretches with the window keeps round corners.
  - Keep content at least 12px in from the card's edge so nothing covers a corner. An
    image under a subform flickers when the subform repaints, so a card whose corner a
    subform must cover uses `square:=True`.
  - Images never take focus or hover, so cards don't flicker when the mouse crosses
    them. Don't build cards from disabled buttons (`Shape = 1`): they flicker on hover.
  - A selected or in-context card turns `CLR_PRIMARY_SOFT` through a conditional
    format on a square `Box` card.
- **Tabs**: either a native `TabCtl` themed through `TabCtl()`, or a `TabStrip` of label
  links, each with a 2px `CLR_PRIMARY` underline (shown for the active tab) over a 1px
  `CLR_BORDER` baseline.
- **Filter chips**: labels toggled between filled `CLR_PRIMARY` with white text and
  transparent with muted text (`StyleChip lbl, isOn, [n]`). Chips and tabs show how many
  rows each would list, "Blocked (3)": pass the count to `StyleChip`, or call
  `SetCount ctl, n`, which keeps the caption without the count in the control's `Tag`.
  Lay a filter row out on one line: chips on the left, then the search box and any filter
  combos to their right, all compact size (section 6.4).

  ```
  [All (24)] [In progress (8)] [Blocked (4)] [Complete (4)]   [Search items...    ] [Category v]
  ```
- **Progress bar**: `Progress nm, expr, x, y, w, h` draws a rounded capsule, 8 to 10px
  tall: a `CLR_BORDER` track filled with `CLR_PRIMARY` to `expr` percent (`ProgressPng`).
  `expr` is a field such as `"[percent_done]"` or a number; values are clamped to 0-100
  and Null shows an empty track. It is bound, so it follows the record in a continuous
  form. Give it a fixed width: the image keeps its shape, so don't anchor it to stretch.
- **Empty states**: a centered muted 10pt label shown in place of an empty list.
- **Search box placeholder**: `Format = "@;""Search by ..."""`.
- **Hierarchical trees and wrapped rows**:
  - A continuous form gives every row the same height, so build rows tall enough for
    2–3 lines.
  - Use GDI `MeasureText` to insert line breaks that fit the label's *current* width,
    ending in an ellipsis.
  - Re-wrap on resize.
  - Indent with level-based weight and row tints (`CLR_ROW1`, `CLR_ROW2`).

### 6.7 Access quirks the design depends on

- **Focus sink**: `NewForm` adds a 1×1 `txtFocusSink` with tab index 0. It takes the
  initial focus so Access doesn't highlight the first field. Move focus to it before
  hiding a focused control (`SetVisible`, which avoids error 2165) and in pill click
  handlers (a focused pill paints square). Change `Visible` and `Enabled` at run time
  only through `SetVisible` and `SetEnabled` (error 2164 for a focused control being
  disabled). Both also cover a subform, tab control or option group that holds the
  focus, and fall back to the first free tab stop on a form with no sink.
- Read-only row text with no click handler is set to `Enabled = False` +
  `Locked = True`. That keeps its normal look while it can never take focus.
- **Disabled controls ignore theme colors.** Lock controls instead of disabling them when
  they must stay themed.
- **Anchored widths**: an anchored control's `Width` still returns the design value. Use
  `form.InsideWidth - form.Width` to get the extra width.
- Set late-bound properties such as an image's `ControlSource` through
  `c.Properties("Name")`. Setting them directly can fail silently.
- An unset conditional ForeColor falls back to black, so always carry the control's own
  fore color (`CF` does this).
- A conditional BackColor only paints on an opaque control (`CF` sets `BackStyle = 1`).

### 6.8 List and detail pages

A page that lists records and shows the selected one beside the list:

```
+--------------------------------------------------------------+------------------+
| Title  (?)                                 [Export] [Print]  |                  |
| [All (24)] [Open (8)] ...  [Search...         ]              |                  |
| +-- Items ------------------------------------------+ +-- Details ----------+  |
| | ITEM          OWNER       STATUS       DUE        | | ITEM                |  |
| | Pump overhaul Rivera     (In progress) 01 Oct     | | [Pump overhaul   ]  |  |
| | ...                                               | | ...          [Save] |  |
| +---------------------------------------------------+ +---------------------+  |
```

- Two cards. The list card on the left holds `GridHeaders` and a `BuildGrid` table and
  fills the width (anchored both ways). The detail card on the right is 360px wide, 12px
  from the list card, anchored right and stretching down. It holds a subform
  (`sfDetail`) inset 16px, and an empty-state label (`lblDetailEmpty`) that says what to
  do, such as "Select an item to see its details."
- The detail form is a single form bound to the table, built `CLR_SURFACE` to match the
  card, with the dialog field layout (captions over inputs) and a Save button
  bottom-right. Validation follows section 3.
- `GridClick` shows the clicked row: `ShowDetail Me, "sfDetail", "subItem", "id=" & id`.
  `ShowDetail` saves a changed record in the panel first, swaps the panel's form only
  when it changes, filters it, and hides the empty state. Call it from the list or the
  page, never from inside the detail form: swapping a subform while one of its own events
  runs is unsafe (the same reason navigation is deferred, section 6.5).
- Whenever the list changes (load, a chip, the search box), call `SelectFirstRow
  Me.sfList`. It selects the first row and calls `GridClick` as a click would; when the
  list is empty it returns Null, and the page calls `ClearDetail Me, "sfDetail"` to show
  the empty state.
- When the detail form saves, it tells the page (`Me.Parent.ItemSaved id`), which
  requeries the list and keeps the row selected.
- Header actions for the list (Print, Export) go top-right and use `PrintGrid` and
  `ExportGrid` on the list's form, so they print exactly the rows shown.

`examples/demo` builds this page as `pgList`.

---

## 7. Color palettes

### 7.1 Mechanism

1. **`tools/make_themes.py`** is the only place colors are defined:
   - `P`: about 65 roles, `name -> (light, dark)`.
   - `FIXED`: colors that never change with the palette.
   - `PALETTES`: named palettes. Each has a base (light or dark) plus overrides.
     `accent_light(sidebar, primary)` derives the sidebar, nav, button and hover shades by
     mixing with white or black.

   Running the script generates two files and prints a WCAG contrast check: 4.5:1 for
   normal palettes, 7:1 for High Contrast.
2. **`modTheme.bas`** holds the *build* palette (light) as `Public Const CLR_X As Long`.
   Every UI role gets a **unique** value: duplicates are nudged by 1–2 units per channel,
   and nothing may equal a `FIXED` color. That lets a color found on a built form identify
   its role.
3. **`modPalette.bas`** holds `PALETTE_COUNT`, `PaletteKey`, `PaletteName`,
   `PaletteIsDark`, `PaletteSwatch(i, k)` and `PaletteOf(c, i)`. `PaletteOf` is a
   `Select Case` that maps each build color to palette *i*'s color and returns any other
   color unchanged.
4. **Run time** (`modDarkMode`):
   - Every form's `OnOpen` is `=ThemeForm([Form])` (set by `NewForm`). For any palette
     other than 1, it repaints:
     - Section colors.
     - Every color property: BackColor, ForeColor, BorderColor, Hover, Pressed and their
       fore colors.
     - Conditional formats.
     - Colors embedded as numbers in pill-image `ControlSource` expressions.
   - Controls tagged `darkonly` become visible in dark palettes.
   - Code that sets a color at run time **must** wrap it: `lbl.ForeColor = Th(CLR_MUTED_FG)`.
   - The choice is saved per Windows user with `SaveSetting APP, "UI", "Theme",
     paletteKey`. With no saved choice, it follows the Windows app mode (light or dark).
   - Switching palettes: Appearance in the sidebar opens `dlgAppearance` (one row per
     palette with three swatches: sidebar, accent, page). Picking a row calls
     `SetPalette`, which saves it and calls `frmMain.Reopen`.
5. **Reports** use `PRINT_*` colors (black, grays, white paper) and are never repainted.

### 7.2 Rules

- Never write `RGB()` or a numeric color in VBA. Use a `CLR_` role, or a `FIXED` constant
  for semantic colors that must not change, such as status colors.
- Build forms in palette 1 only (`UseBuildPalette`).
- To change a color: edit `make_themes.py`, run it, then rebuild with `-FrontEndOnly`.
- Palette 1 must be the build palette with no overrides. Keep the keys `light` and
  `dark`, so saved choices keep working.

### 7.3 Role list (light base → dark base)

Surfaces:

| Role | Light | Dark | Use |
|---|---|---|---|
| CLR_BG | 255,255,255 | 10,14,30 | page background |
| CLR_SURFACE | 255,255,255 | 14,24,44 | cards, panels, rows |
| CLR_INSET | 248,250,252 | 6,11,24 | inset tiles, context banners |
| CLR_DIALOG | 255,255,255 | 14,24,44 | dialog background |
| CLR_ROW1 / ROW2 | 241,245,249 / 248,250,252 | 24,37,62 / 18,30,52 | tree level tints |
| CLR_INPUT_BG | 255,255,255 | 6,11,24 | inputs |
| CLR_CELL_BG | 226,232,240 | 30,44,72 | value cells |
| CLR_BORDER | 226,232,240 | 33,45,72 | card/row borders, lines |
| CLR_INPUT_BORDER | 203,213,225 | 48,62,94 | input borders |

Text:

| Role | Light | Dark |
|---|---|---|
| CLR_FG | 15,23,42 (slate-900) | 241,245,249 |
| CLR_MUTED_FG | 100,116,139 (slate-500) | 148,160,184 |
| CLR_ATTR_FG | 51,65,85 | 203,213,225 |
| CLR_WHITE | 255,255,255 | 255,255,255 (text on colored fills) |
| CLR_ERROR / DANGER_TEXT / WARN_TEXT | 220,38,38 / 185,28,28 / 180,83,9 | 248,113,113 / 248,113,113 / 251,146,60 |
| CLR_GREEN_600 / AMBER_500 | 22,163,74 / 245,158,11 | 74,222,128 / 251,191,36 |

Accent and buttons:

| Role | Light | Dark |
|---|---|---|
| CLR_PRIMARY | 46,109,180 | 59,130,246 |
| CLR_PRIMARY_SOFT (selected row/card) | 236,243,252 | 23,40,78 |
| CLR_PRIMARY_LINE | 191,211,236 | 37,64,120 |
| CLR_LINK | 46,109,180 | 96,165,250 |
| CLR_DESTRUCTIVE | 220,38,38 | 220,38,38 |
| CLR_BTN_BG / FG | 46,109,180 / white | **white / 10,14,30** (dark mode uses white CTAs) |
| CLR_BTN_HOVER / PRESSED | 37,91,153 / 30,75,128 | 226,232,240 / 203,213,225 |
| CLR_OUT_BG / FG / HOVER / BORDER | white / 15,23,42 / 241,245,249 / 203,213,225 | 14,24,44 / 241,245,249 / 30,44,72 / 226,232,240 |

Sidebar:

| Role | Light | Dark |
|---|---|---|
| CLR_SIDEBAR | 27,58,92 (navy) | 20,23,43 |
| CLR_SIDEBAR_LINE | 52,79,108 | 38,42,66 |
| CLR_NAV_FG | 187,196,206 | 205,210,224 |
| CLR_NAV_ACTIVE_BG / FG | 61,88,121 / white | white / 10,14,30 |
| CLR_NAV_MUTED | 141,156,173 | 130,138,160 |
| CLR_BRAND | 27,58,92 | white |
| CLR_BLUE_300, CLR_AVATAR | 147,197,253 / 59,130,246 | same |

Badges (background / foreground):

| Variant | Light | Dark |
|---|---|---|
| success | 220,252,231 / 22,101,52 | 16,58,38 / 74,222,128 |
| warning | 254,243,199 / 146,64,14 | 66,45,12 / 251,191,36 |
| info | 219,234,254 / 30,64,175 | 23,40,84 / 96,165,250 |
| secondary | 226,232,240 / 15,23,42 | 30,41,64 / 203,213,225 |
| severity A, B (red/orange tints) | 254,226,226 / 153,27,27; 255,237,213 / 154,52,18 | 69,16,16 / 252,165,165; 67,30,8 / 253,186,116 |

- KPI accents (same in both bases): blue 59,130,246; indigo 129,140,248; green 74,222,128;
  red 248,113,113; orange 251,146,60; pink 244,114,182.
- Print colors: black text, 64-gray muted text, 230-gray row tint, white paper.

Fixed colors (never repainted), used for semantic scales:
- purple 147,51,234
- blue 37,99,235
- green 22,163,74
- amber 245,158,11
- red 185,28,28
- gray 107,114,128 (not evaluated or none)
- destructive hover 185,28,28 and pressed 153,27,27

### 7.4 Palettes shipped

| Key | Name | Base | Sidebar | Primary | Notes |
|---|---|---|---|---|---|
| light | Navy | light | 27,58,92 | 46,109,180 | default, the build palette |
| dark | Midnight | dark | 20,23,43 | 59,130,246 | navy surfaces, white primary buttons, accent bars shown |
| slate | Slate | light | 51,65,85 | 71,85,105 | neutral, low color |
| forest | Forest | light | 31,61,43 | 47,125,79 | green accent |
| harbor | Harbor | light | 15,59,76 | 14,116,144 | teal accent |
| tealnavy | Teal Navy | light | 10,42,73 | 0,128,147 | Submarine Program Manager's look; its teal is a shade darker than SPM's 0,132,151 so white text passes 4.5:1 |
| graphite | Graphite | dark | 18,18,20 | 130,130,140 | near-monochrome zinc |
| contrast | High Contrast | light | 0,0,0 | 0,64,192 | black text and borders, 7:1 target |

To add a palette:
1. Append an entry to `PALETTES`. For a light variant, use
   `over=accent_light(sidebar, primary)`. For a dark one, list explicit overrides.
2. Run the script and check that the contrast output says `ok`.
3. Rebuild.

The Appearance dialog grows by itself, because it is built from `PALETTE_COUNT`.

---

## 8. Starting a new app in this style (checklist)

1. Copy `kit/`, `examples/build.ps1` and `tools/make_themes.py` from this repository.
   Set `APP_NAME` (the palette choice is saved under it in the registry).
2. In `make_themes.py`, set the roles, `FIXED` and `PALETTES`, then run it.
3. Write modErrors (start from `examples/modules/modErrors.bas`; only the context columns
   are app-specific),
   modGlobals (`APP_NAME`, `g_Silent`, `Notify`, SQL helpers, session accessors) and
   modWrap (`MeasureText`).
4. Write `BuildAll`:
   - Set `g_Silent = True` and call `UseBuildPalette`.
   - Set `g_SrcRoot`.
   - Run named steps.
   - Build subforms, then pages, then dialogs, then the shell.
   - Set startup options.
   - Return `"OK"` or the step that failed.
5. Build the shell: collapsible sidebar with `NavItem`s, `sfPage` host, deferred
   `QueueNav`/`ShowPage`/`RefreshNav`, and `PageKey`/`Reopen`.
6. For each page, write a `Build_pg<Name>` using `PageHeader`, cards, grids and buttons,
   and a `src/forms/pg<Name>.vba` where every handler uses the standard error handler.
7. Copy the dialogs: `dlgAppearance`, `BuildDatePicker` (for date fields), and a non-modal `dlgHelp` with `help.txt` and the F1
   `AutoKeys` macro.
8. Write `README.md` and `DEVELOPING.md` using the structure in section 4, linking to the
   kit's README for the builder catalog.
9. Check every page in a light palette and a dark palette. There is no automated UI test,
   so rebuild with `-FrontEndOnly` and click through.

## 9. Shared table refinements

The table policy and API in the root README extend the original grid builders:
categorical pills use the widest measured category; centered rows and their headers
share the actual runtime column geometry; long text receives additional column width.
A short table scrolls without a footer unless its host is marked Pagination:on.
Full-height rosters can opt into counted footer paging. Reports and exports always use
TableFullFilter rather than the current page's ID filter. Resized rounded cards regenerate
their image to avoid clipping the border. Keep progress, dates and free text out of pills.
