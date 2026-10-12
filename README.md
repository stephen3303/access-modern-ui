# Access Modern UI

A kit for building Microsoft Access front ends that look like a modern web app: forms drawn in code with a consistent look, color palettes users switch at run time (light, dark and five more), rounded pills and badges, styled tables, and a standard error handler and log.

It holds no app-specific tables, data or business rules, so you can start any Access app from it.

## What's here

```
docs/DESIGN_GUIDE.md          how apps in this style are designed and built (read first)
docs/ADOPTION_PLAN.md         moving SABER and Submarine Program Manager onto the kit (agent task list)
kit/modules/*.bas             the reusable modules, loaded before the app's own
kit/forms/dlgAppearance.vba   code-behind for the palette picker
kit/forms/subGrid.vba         code-behind for BuildGrid tables
tools/make_themes.py          generates the build palette and the palette modules, checks contrast
examples/build.ps1            build script that loads the kit, then the app, then calls BuildAll
examples/modules/             modErrors (error handler and log), Build_dlgAppearance
examples/palette/             make_themes.py output for the 7 shipped palettes
```

Requires Microsoft Access 2016 or Microsoft 365 (32- or 64-bit) and Python 3 for the palette generator.

## Modules

| Module | What it does |
|---|---|
| modUIKit | Build-time form and control builders (catalog below) |
| modDarkMode | Palette engine: `ThemeForm` repaints each form on open, `Th()` gives runtime colors, and `SetPalette`/`ShowAppearance` switch palettes |
| modPng | `PillPng`: anti-aliased capsule PNGs for rounded pills and badges, cached in `%LOCALAPPDATA%\<APP_NAME>\pills` |
| modCsv | UTF-8 text files, RFC 4180 CSV read/write, query-to-CSV export |
| modCrypto | SHA-256 through Windows CNG, salted password hashing |
| modKitUtil | `FileExists`, `FolderExists`, `EnsureFolder`, `IsFormOpen`, `RunDialog` |

## Starting a new app with the kit

1. Copy `kit\` into the app's `access\` folder, and `examples\build.ps1` next to it. Copy `tools\make_themes.py` to `access\tools\`.
2. In `make_themes.py`:
   - Set the color roles (`P`) and the palettes (`PALETTES`).
   - Set `FIXED` to the colors that must never change with the palette.
   - Run the script. It writes `src\themes\light\modTheme.bas` and `src\modules\modPalette.bas`, and prints a contrast check.
3. In the app's own modules, provide the names the kit calls (listed below).
4. For the palette picker, copy `examples\modules\modBuildAppearance.bas` to `src\modules` and `kit\forms\dlgAppearance.vba` to `src\forms`.
5. Write a `BuildAll` entry point that calls `UseBuildPalette`, sets `g_SrcRoot`, and builds the forms. The design guide (section 2) describes what it does.

### What the app must provide

| Name | Used by | Where it usually lives |
|---|---|---|
| `CLR_*` roles, `FONT_UI`, `THEME_DARK` | modUIKit | modTheme (generated) |
| `PALETTE_COUNT`, `PaletteKey`, `PaletteName`, `PaletteIsDark`, `PaletteSwatch`, `PaletteOf` | modDarkMode, dlgAppearance | modPalette (generated) |
| `APP_NAME` | modPng | modGlobals |
| `HandleError`, `LogError`, `RaiseAgain` | modDarkMode, modCsv | modErrors (`examples/modules`) |
| `MeasureText(s, size, weight)` (text width in twips) | `TitleTip`/`PageHeader` | modWrap |
| `ShowTip(frm, nm)` | the "?" marks from `TitleTip` | modUI |
| `frmMain` with `PageKey` and `Reopen pg` | `SetPalette` (reopens the window) | frmMain |
| `dlgAppearance` form | `ShowAppearance` | `Build_dlgAppearance` + dlgAppearance.vba |
| `BandLabels()`, `BandBack(label)` | `BandCF` | modUI |
| `BadgeColors(variant)` (returns `Array(back, fore)`) | `BadgeCF` | modUI |
| `subGrid.vba` code-behind | `BuildGrid` tables | copy `kit\forms\subGrid.vba` to src\forms |

An app that doesn't use `BandCF`, `BadgeCF` or `BuildGrid` doesn't need the names those rows list.

## Rules that keep palettes working

- Build forms in the light palette only. `UseBuildPalette` makes sure of that.
- Never write a color literal (`RGB(...)` or a number). Use a `CLR_` role, or a `FIXED` constant for colors that must not change.
- Wrap colors that code sets at run time in `Th()`, for example `lbl.ForeColor = Th(CLR_MUTED_FG)`.
- Every form's `OnOpen` is `=ThemeForm([Form])`. `NewForm` and `NewDialog` set this for you. Controls given `DarkOnly` appear only in dark palettes.

## modUIKit catalog

Sizes are in twips. `px(n)` converts CSS pixels at 15 twips per pixel.

**Forms**

| Builder | What it does |
|---|---|
| `NewForm recSrc, view, w, h` | Start a form. `view` is 0 for single, 1 for continuous. Sets `ThemeForm` as its Open hook. |
| `NewDialog caption, w, h` | Modal pop-up dialog that sizes to its design |
| `NewReport recSrc, w` | Start a report. Reports use the `PRINT_` colors and are never repainted. |
| `SaveForm name [, codeName]` | Save under the final name and attach `src\forms\<name>.vba` as the code-behind |
| `Trace s` / `BuildTrace` | Record build steps, so a failed build reports where it stopped |

**Controls**

| Builder | What it does |
|---|---|
| `Lbl nm, cap, x, y, w, h [, size, bold, fore]` | Label with a transparent background |
| `Txt nm, src, x, y, w, h` | Text box (a field name, or an expression starting with `=`) |
| `RO nm, src, x, y, w, h` | Read-only text: no border, not a tab stop |
| `Cbo` / `Lst` / `Chk` | Combo box, list box, check box with its label |
| `Btn nm, cap, x, y, w, h [, kind]` | Button. `kind` is `primary`, `outline`, `ghost`, `destructive` or `nav`. |
| `Link nm, cap, x, y, w, h, onClick` | Clickable text (links and tab labels) |
| `Box nm, x, y, w, h, back [, border]` | Rectangle. Pass -1 for no fill or no border. |
| `RowBack` | Row background for a continuous form that paints but never takes focus |
| `Img` | Unbound image whose picture is set at run time |
| `SubF nm, srcObj, x, y, w, h` | Subform control |
| `TabCtl nm, x, y, w, h, "name\|caption", ...` | Tab control |
| `TabStrip x, y, w, handler, "key\|Caption", ...` | Tab strip of labels with underlines |

**Page furniture**

| Builder | What it does |
|---|---|
| `PageHeader title, tip` | Page title with a "?" explanation mark |
| `TitleTip` | "?" mark placed after any title |
| `AccentBar` | Left accent bar beside section titles (dark palettes only) |
| `ContextBanner` | Inset banner showing the current record, with an optional pill |

**Tables and conditional colors**

| Builder | What it does |
|---|---|
| `BuildGrid name, cols, w` | Styled read-only table subform (columns as `"field\|caption\|width px\|options"`) |
| `GridHeaders` | Column captions above a `BuildGrid` table |
| `CF c, expr, back, fore [, bold]` | Expression-based conditional format |
| `BadgeCF c, field, value, variant, ...` | Rounded badge colors keyed on a status field |
| `BandCF c, field` | Status band pill colors |
| `RoundCF` / `PillLabel` | Rounded pill image behind a control (bound, or painted in code) |

**Layout helpers**

| Builder | What it does |
|---|---|
| `Anchor c, h, v` | Anchoring. 0 is left/top, 1 is right/bottom, 2 is both. |
| `VCenter` | Vertically center single-line text |
| `Ev c, "OnClick", ...` | Point event properties at the form module's handlers |
