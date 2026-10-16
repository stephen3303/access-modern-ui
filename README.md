# Access Modern UI

A kit for building Microsoft Access front ends that look like a modern web app: forms drawn in code with a consistent look, color palettes users switch at run time (light, dark and six more), rounded cards, pills, badges and progress bars, styled tables you can print and export, and a standard error handler and log.

It holds no app-specific tables, data or business rules, so you can start any Access app from it.

## What's here

```
docs/DESIGN_GUIDE.md          how apps in this style are designed and built (read first)
docs/ADOPTION_PLAN.md         moving SABER and Submarine Program Manager onto the kit (agent task list)
kit/modules/*.bas             the reusable modules, loaded before the app's own
kit/forms/dlgAppearance.vba   code-behind for the palette picker
kit/forms/dlgColumns.vba      code-behind for the column chooser (BuildColumnsDialog, ChooseGridColumns)
kit/forms/dlgDatePick.vba     code-behind for the date picker (BuildDatePicker)
kit/forms/dlgFilter.vba       code-behind for the header filter (BuildFilterDialog, GridFilter)
kit/forms/dlgMessage.vba      code-behind for the message dialog (BuildMessageDialog, KitMsgBox)
kit/forms/dlgToast.vba        code-behind for the toast (BuildToast, KitToast)
kit/forms/dlgInput.vba        code-behind for the input dialog (BuildInputDialog, KitInputBox)
kit/forms/frmSplash.vba       code-behind for the splash screen (BuildSplash)
kit/forms/subGrid.vba         code-behind for BuildGrid tables
tools/make_themes.py          generates the build palette and the palette modules, checks contrast
tools/AccessWatch.ps1         makes a build script report Visual Basic and Access dialogs instead of hanging
tools/VbaLineNumbers.ps1      numbers the build copy of the code so the error log names the failing line
tools/review-mcp/             MCP server for visual review (also a Claude Desktop extension): opens, sizes and captures forms without the screen or mouse
examples/build.ps1            build script that loads the kit, then the app, then calls BuildAll
examples/modules/             modErrors (error handler and log), Build_dlgAppearance
examples/palette/             make_themes.py output for the 8 shipped palettes
examples/demo/                a small app that shows the kit's cards, tables and list/detail pages
```

To see the kit working, run `examples\demo\build.ps1` and open `examples\demo\KitDemo.accdb` (see [examples/demo/README.md](examples/demo/README.md)).

Requires Microsoft Access 2016 or Microsoft 365 (32- or 64-bit) and Python 3 for the palette generator.

## Modules

| Module | What it does |
|---|---|
| modTableUI | Responsive columns, matching headers, counted filters, uniform tag widths and opt-in pagination |
| modTextUI | Standalone measured single-line text fitting and header captions with full-text tooltips |
| modUIKit | Build-time form and control builders (catalog below) |
| modDarkMode | Palette engine: `ThemeForm` repaints each form on open (and, on Windows 11, colors pop-up title bars and borders through `ThemeTitleBar`), `Th()` gives runtime colors, and `SetPalette`/`ShowAppearance` switch palettes |
| modPng | Anti-aliased PNGs for rounded shapes: `PillPng` (pills and badges), `CardPng` (card corners), `ProgressPng` (progress capsules), `ColumnPillPng` (uniform-width pills in a table column), `RoundedCardPng` (whole resizable card backgrounds), `TogglePng` and `RingPng` (toggle switch track and focus ring), cached in `%LOCALAPPDATA%\<APP_NAME>\pills` |
| modKitUI | Run-time page helpers: `GridFit`, `GridHighlight`, `GridSort`, `SelectFirstRow`, `ShowDetail`, `ClearDetail`, `PickDate`, `MarkDate`, `FederalHoliday`, `ToggleSwitch`, `ToggleFocus`, `SegmentClick`, `SetSegment`, `SegmentValue`, `ShowAlert`, `HideAlert`, `SetBreadcrumb`, `CrumbClick`, `ToggleSecret`, `HideSecrets`, `KitMsgBox`, `KitToast`, `KitInputBox`, `CheckRequired`, `SplashStatus`, `SetCount`, `StyleChip` (catalog below) |
| modGridOut | `PrintGrid` and `ExportGrid`: print or export a table as the user sees it |
| modGridFilter | Click-to-filter table headers: `GridFilter` (a caption funnel's click), `SetGridFilter`, `SetColumnFilter`, `ClearGridFilters`, `GridFilterWhere`, `GridColumnValues`, and `BuildFilterDialog` for `dlgFilter` |
| modCsv | UTF-8 text files, RFC 4180 CSV read/write, query-to-CSV export |
| modCrypto | SHA-256 through Windows CNG, salted password hashing |
| modTrust | Startup check for an untrusted folder: `BuildTrustCheck mainForm` (build step, in place of `StartUpForm`) adds an `AutoExec` macro and `dlgTrust`, a page explaining how to add the folder as a Trusted Location. `TrustNeeded` and `InTrustedLocation` read the Trust Center settings (never change them), `ShowTrustHelp` opens the page. It replaces any `AutoExec` the app has |
| modKitUtil | `FileExists`, `FolderExists`, `EnsureFolder`, `IsFormOpen`, `RunDialog`, `CenterOnApp`, `SetVisible`, `SetEnabled` |
| modKitUpdate | `CheckForUpdate current, folder, [advice]`: at startup, tells the user when the release folder's `version.txt` names a newer version (`ReleasedVersion`, `CompareVersions`) |
| modBackdrop | `DimApp`/`UndimApp`: dims the app behind a modal dialog, like a web modal (`RunDialog`, `KitMsgBox` and `KitInputBox` use it) |

## Starting a new app with the kit

1. Copy `kit\` into the app's `access\` folder, and `examples\build.ps1` and `tools\VbaLineNumbers.ps1` next to it. Copy `tools\make_themes.py` to `access\tools\`.
2. In `make_themes.py`:
   - Set the color roles (`P`) and the palettes (`PALETTES`).
   - Set `FIXED` to the colors that must never change with the palette.
   - Run the script. It writes `src\themes\light\modTheme.bas` and `src\modules\modPalette.bas`, and prints a contrast check.
3. In the app's own modules, provide the names the kit calls (listed below).
4. For the palette picker, copy `examples\modules\modBuildAppearance.bas` to `src\modules` and `kit\forms\dlgAppearance.vba` to `src\forms`. For date fields, copy `kit\forms\dlgDatePick.vba` to `src\forms` too, for messages copy `kit\forms\dlgMessage.vba` and `kit\forms\dlgToast.vba`, for questions copy `kit\forms\dlgInput.vba`, for the column chooser copy `kit\forms\dlgColumns.vba`, for header filters copy `kit\forms\dlgFilter.vba`, and for the splash screen copy `kit\forms\frmSplash.vba`: `SaveForm` attaches code-behind from `src\forms` only, so without them `dlgDatePick`, `dlgMessage`, `dlgToast`, `dlgInput`, `dlgColumns`, `dlgFilter` and `frmSplash` build with no code.
5. Write a `BuildAll` entry point that calls `UseBuildPalette`, sets `g_SrcRoot`, and builds the forms. The design guide (section 2) describes what it does.
6. For startup, call `BuildTrustCheck "frmMain"` (or the login form, or `frmSplash` when the app has a splash screen) last in the startup step, instead of setting `StartUpForm`. When Access would ask to enable content, users then see how to trust the folder.

### What the app must provide

| Name | Used by | Where it usually lives |
|---|---|---|
| `CLR_*` roles, `FONT_UI`, `THEME_DARK` | modUIKit | modTheme (generated) |
| `PALETTE_COUNT`, `PaletteKey`, `PaletteName`, `PaletteIsDark`, `PaletteSwatch`, `PaletteOf` | modDarkMode, dlgAppearance | modPalette (generated) |
| `APP_NAME` | modPng, modKitUI, modKitUpdate, `BuildSplash`, modDarkMode, modGridCols (registry key for saved settings) | modGlobals |
| `HandleError`, `LogError`, `RaiseAgain` | modDarkMode, modCsv, modKitUI, modKitUpdate, modGridCols, modGridFilter | modErrors (`examples/modules`) |
| `MeasureText(s, size, weight)` (Segoe UI text width in twips; size in points, weight 400 normal / 700 bold) | modTextUI, modPng, `TitleTip`/`PageHeader` | modWrap |
| `ShowTip(frm, nm)` | the "?" marks from `TitleTip` | modUI |
| `frmMain` with `PageKey` and `Reopen pg` | `SetPalette` (reopens the window) | frmMain |
| `dlgAppearance` form | `ShowAppearance` | `Build_dlgAppearance` + dlgAppearance.vba |
| `dlgDatePick` form | `PickDate` (`DateBox` buttons) | call the kit's `BuildDatePicker` with the app's dialogs + copy `kit\forms\dlgDatePick.vba` to src\forms |
| `dlgMessage` form | `KitMsgBox` (the app's `Notify` and `HandleError`) | call the kit's `BuildMessageDialog` with the app's dialogs + copy `kit\forms\dlgMessage.vba` to src\forms |
| `dlgToast` form | `KitToast` | call the kit's `BuildToast` with the app's dialogs + copy `kit\forms\dlgToast.vba` to src\forms |
| `dlgColumns` form | `ChooseGridColumns` (a table's Columns button) | call the kit's `BuildColumnsDialog` with the app's dialogs + copy `kit\forms\dlgColumns.vba` to src\forms |
| `dlgFilter` form | `GridFilter` (a caption's funnel, `GridHeaders ..., filterSf`) | call the kit's `BuildFilterDialog` with the app's dialogs + copy `kit\forms\dlgFilter.vba` to src\forms |
| `dlgInput` form | `KitInputBox` | call the kit's `BuildInputDialog` with the app's dialogs + copy `kit\forms\dlgInput.vba` to src\forms |
| `frmSplash` form (optional) | the startup form, `SplashStatus` | call the kit's `BuildSplash` + copy `kit\forms\frmSplash.vba` to src\forms; pass `frmSplash` to `BuildTrustCheck` |
| `BandLabels()`, `BandBack(label)` | `BandCF` | modUI |
| `BadgeColors(variant)` (returns `Array(back, fore)`) | `BadgeCF`, `BuildGrid` | modUI |
| `StatusVariant(value)` (a Variant, may be Null; returns a badge variant or "") | `BuildGrid` pill columns (`p`) | modUI |
| `subGrid.vba` code-behind | `BuildGrid` tables | copy `kit\forms\subGrid.vba` to src\forms |
| `GridClick(grid, dbl)` on the form hosting a table | `BuildGrid` rows, `SelectFirstRow` | that form's code-behind |

An app that doesn't use `BandCF`, `BadgeCF` or `BuildGrid` doesn't need the names those rows list. `BandLabels`, `BandBack` and `BadgeColors` are needed to compile modUIKit, so an app without bands or badges provides small stand-ins (the demo's `modDemo` shows them).

## Rules that keep palettes working

- Build forms in the light palette only. `UseBuildPalette` makes sure of that.
- Never write a color literal (`RGB(...)` or a number). Use a `CLR_` role, or a `FIXED` constant for colors that must not change.
- Wrap colors that code sets at run time in `Th()`, for example `lbl.ForeColor = Th(CLR_MUTED_FG)`.
- Every form's `OnOpen` is `=ThemeForm([Form])`. `NewForm` and `NewDialog` set this for you. Controls given `DarkOnly` appear only in dark palettes.

## modUIKit catalog

Sizes are in twips. `px(n)` converts CSS pixels at 15 twips per pixel.

**Spacing**

The designer can choose how roomy layouts are. Call `SetSpacing` once at the start of the build, before any form is created. Without it the scale is 1 and every layout builds exactly as before.

| Helper | What it does |
|---|---|
| `SetSpacing "compact"` / `"default"` / `"comfortable"` | Scale gaps and padding by 0.75, 1 or 1.25. A number from 0.5 to 2 also works. The choice is saved in the database (property `KitSpacing`), so run-time layouts match the build. |
| `Sp(n)` | A gap or padding of `n` CSS px at the chosen spacing, in twips, rounded to whole px. Use `Sp` for space and `px` for control sizes. |
| `SP_PAGE` 32, `SP_DIALOG` 24, `SP_CARD_INSET` 16, `SP_CARD_GAP` 12, `SP_ROW_GAP` 8, `SP_ITEM_GAP` 4 | The guide's spacing tokens, for `Sp(SP_CARD_GAP)` and so on. |
| `SpacingScale()` | The scale in use. |

Control heights and font sizes never scale. `TabStrip` and `LayoutFilterRow` already space their items with `Sp(SP_ITEM_GAP)`.

**Forms**

| Builder | What it does |
|---|---|
| `NewForm recSrc, view, w, h` | Start a form. `view` is 0 for single, 1 for continuous. Sets `ThemeForm` as its Open hook. |
| `NewDialog caption, w, h` | Modal pop-up dialog that sizes to its design. Its code-behind's `Form_Load` should end with `CenterOnApp Me` (the builder can't add it). |
| `NewReport recSrc, w` | Start a report. Reports use the `PRINT_` colors and are never repainted. |
| `SaveForm name [, codeName]` | Save under the final name and attach `src\forms\<name>.vba` as the code-behind |
| `BuildSplash [nextForm, subtitle, glyph]` | Builds `frmSplash`, a 480 x 280px splash screen (taller at a roomier `SetSpacing`) in the palette's sidebar colors: an MDL2 icon glyph, `APP_NAME`, a subtitle, the version (database property `AppVersion`, else `AppBuild`) and a status line. Pass it to `BuildTrustCheck` as the startup form. It opens `nextForm` (default `frmMain`), whose Load does the start-up work under it, then closes, after at least 1.2 seconds. |
| `HideAccessSplash` | Writes a 1 x 1 pixel `<database name>.bmp` next to the front end, which Access shows in place of its own splash. Call it last in BuildAll's startup step; ignore the .bmp in git and ship it with the front end. |
| `Trace s` / `BuildTrace` | Record build steps, so a failed build reports where it stopped |

**Controls**

| Builder | What it does |
|---|---|
| `Lbl nm, cap, x, y, w, h [, size, bold, fore]` | Label with a transparent background |
| `Txt nm, src, x, y, w, h` | Text box (a field name, or an expression starting with `=`). A text box that isn't an `=` expression shows the focus ring. |
| `BuildDatePicker` | Builds `dlgDatePick`, the month calendar `DateBox` buttons open. Call it once with the app's dialogs. |
| `BuildMessageDialog` | Builds `dlgMessage`, the dialog `KitMsgBox` opens. Call it once with the app's dialogs. |
| `BuildToast` | Builds `dlgToast`, the toast `KitToast` shows. Call it once with the app's dialogs. |
| `BuildInputDialog` | Builds `dlgInput`, the dialog `KitInputBox` opens. Call it once with the app's dialogs. |
| `BuildColumnsDialog` (modGridCols) | Builds `dlgColumns`, the column chooser `ChooseGridColumns` opens. Call it once with the app's dialogs. |
| `BuildFilterDialog` (modGridFilter) | Builds `dlgFilter`, the value checklist a caption's funnel opens (`GridFilter`). Call it once with the app's dialogs. |
| `DateBox nm, src, x, y, h` | Date entry: a 120px text box (dd mmm yyyy, left-aligned) and a calendar button `cal_<nm>` 4px to its right that opens the date picker. The pair is `DATEFIELD_W` (154px) wide. |
| `SecretBox nm, src, x, y, w, h` | Password or other sensitive entry: a text box that shows asterisks (input mask `Password`) and an eye button `eye_<nm>` 4px to its right, 30px wide, that shows or hides the text. The pair is `w` wide. Masking is display only, so hash passwords (`HashPassword`) before saving. Single-record forms only: on a continuous form it raises error 5. |
| `RO nm, src, x, y, w, h` | Read-only text: no border, not a tab stop |
| `Cbo` / `Lst` / `Chk` | Combo box (with the focus ring), list box, check box with its label |
| `Toggle nm, src, x, y [, w, h, onChange]` | On/off switch for a yes/no field `src` ("" for unbound): a 44 x 24px track, `CLR_PRIMARY` with the knob right when on, `CLR_MUTED_FG` with it left when off, the knob `CLR_SURFACE`. Click it, or Tab to it and press Space. The value is the hidden check box `<nm>_val`; `onChange` names a Public Sub in the form's code-behind called after each change. Hide it with `SetVisible` on `<nm>`; `SetEnabled` on `<nm>` stops clicks but doesn't dim it. Single forms only (the focus ring would show on every row of a continuous form). |
| `Segmented nm, src, x, y, w, items [, h, onChange]` | Segmented control (a web toggle group): 2 to 6 options side by side on a rounded `CLR_SECONDARY_BG` track, the picked one bold `CLR_FG` on a raised `CLR_SURFACE` thumb. `items` is `"week|Week;month|Month"`; `src` is the text or number field it shows and changes ("" for unbound, starting on the first item). The options share `w` equally; `h` is 32px by default. Click an option, or Tab to it and press Space. The value is the hidden text box `<nm>_val` (`SegmentValue`); a bound control's thumb follows the record with no code. `onChange` names a Public Sub in the form's code-behind called after the user picks another option. Single forms only (error 5 on a continuous form, a bad or repeated key, or options under 40px). |
| `Alert nm, x, y, w, h [, variant, title, msg, dismiss]` | Inline alert (a web callout) that stays on the page: a panel tinted `CLR_INFO_BG`, `CLR_SUCCESS_BG`, `CLR_WARNING_BG` or `CLR_SEVA_BG` for `info` (default), `success`, `warning` or `destructive`, with a 4px accent bar and icon in `KitToast`'s colors, a bold title and the wrapped message. 64px fits a title and two lines; an untitled alert under 48px is compact, one line centered (a page's header row). `dismiss:=True` adds a close mark that hides it. Error 5 for an unknown variant. |
| `Btn nm, cap, x, y, w, h [, kind]` | Button. `kind` is `primary`, `outline`, `ghost`, `destructive` or `nav`. |
| `Link nm, cap, x, y, w, h, onClick` | Clickable text (links and tab labels) |
| `Box nm, x, y, w, h, back [, border]` | Rectangle. Pass -1 for no fill or no border. |
| `Rule nm, x, y, w, h, color` | 1px hairline divider, flat (h = 0) or upright (w = 0). A filled 1px `Box` paints 2px, so dividers use this. |
| `Card nm, title, x, y, w, h [, parentCtl, hAnchor, vAnchor, back, square]` | Rounded card: surface, 1px border, 10pt title and a dark-only accent bar. `square:=True` keeps square corners where a subform covers a corner. |
| `Progress nm, expr, x, y, w, h` | Rounded progress capsule filled to `expr` percent (a field or a number; Null is 0) |
| `RowBack` | Row background for a continuous form that paints but never takes focus |
| `Img` | Unbound image whose picture is set at run time |
| `SubF nm, srcObj, x, y, w, h` | Subform control. Its Enter and Exit keep the focus ring right for the fields inside it. |
| `TabCtl nm, x, y, w, h, "name\|caption", ...` | Tab control |
| `Breadcrumb prefix, x, y, w, slots, handler, "key|Caption", ...` | Breadcrumb trail of links separated by chevrons, root first. Every crumb but the last calls `handler` (a Public Sub or Function on the form taking the key); the last is the current page in `CLR_FG`. Crumbs are as wide as their text (`MeasureText`); a path wider than `w` folds its middle into an ellipsis whose tooltip names the folded crumbs. Build `slots` (2 to 12) for the deepest path and change it with `SetBreadcrumb`. |
| `TabStrip x, y, w, handler, "key\|Caption", ...` | Tab strip of labels with underlines. `"key\|Caption\|n"` shows a count, "Caption (n)". |

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
| `BuildGrid name, cols, w [, rowLines, back, zebra, rowPx, fit]` | Styled read-only table subform. Columns are `"field\|caption\|width px\|options"`; options `r` right, `c` center, `b` bold, `m` muted, `p` status pill (`StatusVariant`), `w` wrap onto `rowLines` lines instead of cutting text off, `s` mark the text matching the page's search (`GridHighlight`). `zebra` fills every other row, `rowPx` sets the row height, `fit` lets `GridFit` scale the columns. Clicking a row selects and opens it: no open buttons. |
| `GridHeaders prefix, cols, x, y [, parentCtl, band, w, sortSf, filterSf, onFilter, fit]` | Column captions above a `BuildGrid` table, each as wide as its cells. Pass the table's width as `w` so the last caption reaches the table's edge, and `fit:=True` for a `fit` table. `band:=True` puts them on a filled strip. `sortSf` (the table's subform control) makes them sort the table on click, again to reverse, with an up/down arrow; a fifth column part names a different sort field. `filterSf` (the same control) adds a funnel to each field column's caption that filters the table by the column's values (`GridFilter`); `onFilter` names a Public Sub on the form called with the subform control's name after the user changes a filter. |
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

**Run time** (modKitUI, modGridOut, modGridCols, modGridFilter, modKitUtil, modBackdrop)

| Helper | What it does |
|---|---|
| `GridFit sf [, prefix]` | Scale a `fit` table's columns, and its `GridHeaders` captions, to the subform's width. Call it from the host's `Form_Resize`. |
| `ChooseGridColumns sf [, prefix]` (modGridCols) | A Columns button's click: let the user hide, show and reorder a `fit` table's columns (and its captions) in `dlgColumns`, and save the choice per Windows user and table |
| `RestoreGridColumns sf [, prefix]` (modGridCols) | Put back the user's saved column layout. Call it once the table's record source is set, and after anything that shows its captions again. |
| `GridHighlight sf, term` | Mark `term` in a table's `s` columns, in `CLR_WARNING_FG` on `CLR_WARNING_BG`, ignoring case. Call it with the search box text whenever the search changes; `""` clears the marks. Print and export show the plain text. |
| `GridScroll g` | A `BuildGrid` table's OnCurrent (set by `BuildGrid`): show its vertical scroll bar only when its rows overflow |
| `GridScrollAll frm` | `GridScroll` for every table on a window and the forms in it. Call it from the window's `Form_Resize`. |
| `GridFilter host, sfName, prefix, col` (modGridFilter) | A caption funnel's click (set by `GridHeaders ..., filterSf`): list the column's values in `dlgFilter` (those the other filters leave, plus those the column's own filter keeps, in the column's sort order, blanks last) and show only the rows whose value the user keeps ticked. Ticking every value removes the filter. While the column is filtered the funnel shows white on a primary-color chip. |
| `SetGridFilter sf, where` (modGridFilter) | The page's own filter (search, chips) for a table with funnels: the table shows rows matching `where` and every header filter. Use it instead of setting the table's `Filter`, which would replace the header filters. |
| `SetColumnFilter sf, prefix, col, keep` / `ClearGridFilters sf` (modGridFilter) | Filter a column to the values in the array `keep` in code (`Empty` removes it), or remove every header filter. `GridFilterWhere sf` returns the header filters for counting rows. |
| `GridSort host, sfName, prefix, col` | A sortable `GridHeaders` caption's click (set by `GridHeaders ..., sortSf`): sort the table by that column (ties in id order), or reverse it, and move the arrow; a paginated table goes back to page 1. Call it to sort in code so the arrow matches. |
| `SelectFirstRow sf` | Select a table's first row and call the host's `GridClick`, as a click would. Returns the id, or Null for an empty table. |
| `ShowDetail host, detailCtl, formName, where [, emptyCtl]` | List/detail pages: show the matching record in the detail subform, saving a changed one first, and hide the empty state. Returns False, changing nothing, when the detail form refuses the save (its `BeforeUpdate` cancels, e.g. `CheckRequired`). |
| `ClearDetail host, detailCtl [, emptyCtl]` | Empty the detail subform and show the empty state. Returns False like `ShowDetail`. |
| `SaveDetail host, detailCtl` | Save the detail subform's changed record before a page change that doesn't go through `ShowDetail` (a filter chip). Returns False when the detail form refuses. |
| `PickDate frm, nm` | A `DateBox` calendar button's click: open the kit's date picker (`dlgDatePick`) for box `nm` and write the chosen date into it (not while the box is disabled or locked, or its form is read-only) |
| `SplashStatus msg` | Show what start-up is doing on the splash screen's status line, "Opening the data file...". Does nothing once the splash has closed. |
| `CloseSplash` | Close the splash screen now. `KitMsgBox` calls it, so a start-up message never opens behind the splash; call it yourself before a plain `MsgBox` or a file dialog at start-up. |
| `ToggleSwitch frm, nm` / `ToggleFocus frm, nm, hasFocus` | A `Toggle`'s click and focus handlers (set by the builder): flip the value and call `onChange` (not while its value is locked or its form is read-only); show the switch's focus ring and take an input's `FocusRing` off |
| `SegmentClick frm, nm, key` / `SetSegment frm, nm, key` / `SegmentValue(frm, nm)` | A `Segmented` option's click (set by the builder): pick `key`, move the thumb and call `onChange` (not when `key` is already picked, while `<nm>_val` is disabled or locked, or on a read-only form). `SetSegment` picks in code without calling `onChange` (error 5 for an unknown key); `SegmentValue` returns the picked key as text. The focus ring is `ToggleFocus`, as for a `Toggle`. |
| `ShowAlert frm, nm, msg [, variant, title]` / `HideAlert frm, nm` / `AlertShown(frm, nm)` | Show an `Alert` with a new message, variant and title ("" for none; the message moves to fit) in the current palette, hide it (its close mark calls `HideAlert`), or ask whether it shows. Error 5 for an unknown variant. |
| `SetBreadcrumb frm, prefix, "key|Caption", ...` / `CrumbClick frm, prefix, i` | Show a new path on a `Breadcrumb` (error 5 for more items than slots), and a crumb's click (set by the builder): call the handler with the crumb's key, except on the current page. |
| `MarkDate d, label` / `ClearMarkedDates` | Mark an app date (a deadline, a drill) on the date picker with a tooltip label; a second label on the same day is added after "; ". Set marks before the picker opens, for example at startup from the app's deadlines table. |
| `FederalHoliday d` | The US federal holiday on day `d`, on its observed day (a Saturday holiday on Friday, a Sunday one on Monday, "(observed)" added), or "". The date picker marks these on its own. |
| `ToggleSecret frm, nm` | A `SecretBox` eye button's click: show box `nm`'s text if hidden, hide it if shown, and put the focus back in the box |
| `HideSecrets frm` | Hide every `SecretBox` on a form again. Call it from the form's `Form_Current` so the next record opens hidden. |
| `FocusRing frm, nm, isOn` | The GotFocus (`True`) and LostFocus (`False`) that `Txt` and `Cbo` set: a `CLR_PRIMARY` border on the field being typed in, and the border it had before (a required field's red outline too) when it loses the focus. Only the border color changes, so nothing moves. Not on locked fields, read-only forms or continuous-form rows. One field is ringed at a time. A form that handles a kit input's GotFocus or LostFocus itself calls it from that handler. |
| `FocusRingEnter frm, nm` / `FocusRingClear` | The Enter and Exit that `SubF` sets, so a field inside a subform loses its ring when the focus leaves the subform and gets it back on return (Access fires no LostFocus or GotFocus for it then). `FocusRingClear` is also the GotFocus of `Btn`, `Chk`, `Lst` and the focus sink. |
| `KitMsgBox msg [, style, title]` | `MsgBox` in the kit's look, centred over Access (a `MsgBox` is centred on the monitor, so on a large or second screen it opens outside Access). Takes OK, OK/Cancel, Yes/No and Yes/No/Cancel, the four icons and `vbDefaultButton2`/`3`, and returns what `MsgBox` would. Anything else, or a dialog that isn't built yet, falls back to `MsgBox`. Use it in the app's `Notify` and `HandleError`. |
| `KitInputBox prompt [, title, defaultText, masked]` | `InputBox` in the kit's look, centred over Access like `KitMsgBox`. Returns the text typed on OK ("" when left empty) or Null on Cancel or Esc. `masked:=True` shows each character as * for a PIN or password. A dialog that isn't built yet falls back to `InputBox`, except a masked question, which returns Null and is logged rather than shown in plain text. |
| `CheckRequired frm, fields` | Save check: outline every empty required input in `CLR_DESTRUCTIVE`, list them in one `KitMsgBox` and focus the first. `fields` is `"txtName=Item;txtOwner=Owner"` in tab order. Returns True when all are filled; a filled input gets its normal border back. `RequiredMissing` does the same without the message and returns the list. |
| `KitToast msg [, variant, ms]` | A message that doesn't interrupt: a toast in the bottom-right corner of `frmMain` that fades after `ms` milliseconds (default 3000) without taking the focus. `variant` is `success`, `info` (default), `warning` or `destructive` and sets the icon and accent bar. Up to 3 lines; a new toast replaces the one showing; a click closes it. Use it to confirm what the user just did ("Saved"); anything they must read or answer goes in `KitMsgBox`. If `dlgToast` isn't built, the message goes to the status bar. |
| `RunDialog(name [, openArgs, dimBehind])` | Open a modal dialog and return True when it set `TempVars("DialogResult") = "ok"`. The app behind it is dimmed while it is open (pass `dimBehind:=False` to skip), and the dim comes off when it closes or fails to open. |
| `DimApp` / `UndimApp` | Dim the Access window with a see-through layer of `CLR_SIDEBAR`, and remove it. Pair every `DimApp` with one `UndimApp`, on the error path too; nested dialogs share one layer. Only needed for a modal form opened without `RunDialog` or `KitMsgBox`. Dialogs built with `NewDialog` call `BackdropOnClose` as they close, which removes a dim a VBA reset left behind. |
| `SetCount ctl, n` | Show a count after a tab or chip caption, "Name (n)" |
| `StyleChip lbl, isOn [, n]` | Filter chip on or off, with an optional count |
| `SetVisible c, v` | Show or hide a control, moving the focus away when it or a control inside it has the focus (avoids error 2165) |
| `SetEnabled c, v` | Enable or disable a control, moving the focus away when it has the focus (avoids error 2164) |
| `PrintGrid gridForm, title [, subtitle]` | Print preview of a table's rows as shown, with its captions; landscape when wide |
| `ExportGrid gridForm, path` | Write a table's rows as shown to CSV, or to Excel when `path` ends in `.xlsx`. Returns the row count. |

## Shared table policy

### Text fitting without table dependencies

`modTextUI` can be imported alone. Its only application hook is
`MeasureText(text, size, weight)`, which must measure Segoe UI text in twips.
Access's normal form/control types and `Nz` are also required.

Apps without an existing measurement hook can import the optional Windows adapter
`kit/adapters/modTextMeasure.bas`. It measures Segoe UI through GDI, supports
32-bit and 64-bit Access, and returns twips using a canonical 96 DPI font.
Import it **instead of** an existing `MeasureText` implementation to avoid a
duplicate public name. Adapters are outside `kit/modules` and are not imported
automatically. The native test's `-UseGdiMeasurement` switch imports this adapter,
removes its test stand-in, and checks proportional character widths before
running the complete fixture.

- `FitSingleLineText(value, width [, size = 9, weight = 400])` replaces line breaks
  with spaces and returns measured text with an ellipsis when it does not fit.
  Width is the available text area in twips, excluding control margins. A null
  value or a width too small for the ellipsis returns an empty string. Keep the
  original bound value intact and available in the record editor.
- `FitHeaderText(heading, fullCaption [, preferredSize = 9, minimumSize = 8])`
  fits the caption to the control's inner width, reducing the font to the supplied
  minimum before ellipsizing. Pass the original caption on every resize so a
  wider column restores its complete text and preferred font. The native tooltip
  retains up to 255 characters of the full caption. Use a Segoe UI label or button
  with margin properties.

Native `ControlTipText` is shared across repeated instances of a continuous form
control. It cannot reliably reveal each hovered row's bound value. Use header
tooltips for captions and the record editor for complete row values.

Use pills for finite categories, statuses and filter tags. Leave names, free text,
dates and numeric progress as plain text. Use TagWidth(values) or TableTagWidth(form,
field) to measure the widest category, then use ColumnPillPng with that width for every
pill in the same row or column. Text measurement comes from the app's MeasureText hook.

BuildGrid keeps read-only cells out of the tab order. Non-clickable cells are disabled
and locked; clickable cells move focus to txtFocusSink before the parent callback.
subGrid resolves its hosting control by form instance, even when reused twice.

Call ResizeGrid on resize, then AlignGridHeaders with the corresponding txtCol names.
The column option t gives Action/Description fields three shares of extra width.
CenterTableCell gives native bound values one centered line; open the record editor to
read or edit longer values. Keep multiline editors outside the continuous table.

RoundedCard / ResizeRoundedCard regenerate the background (RoundedCardPng) at the actual control size,
so corners remain rounded when a panel expands. Native image/subform borders stay off.

### Pagination is optional for each table

GridFooter subName, prefix, keyField, enabled builds the parent footer. enabled defaults
to False. A hosting subform Tag of Pagination:on enables it; unmarked tables retain
their full height and normal scrolling. SetTablePagination changes this token without
affecting other tags. Mark mainly full-height roster tables; omit it on short detail tables.

Call LayoutTableFooter from the host load/resize handler after sizing the table. It
reserves footer space only when enabled, preserves the scoped filter, and restores both
height and all matching rows when disabled. The source must have a stable ORDER BY and
a unique non-null key. Use TableFullFilter for exports and reports so every matching
record is included, regardless of the visible page.

Example build:
    SubF "sfRoster", "subRoster", px(24), px(100), px(700), px(400)
    GridFooter "sfRoster", "roster", "id", True

Example host resize:
    ResizeGrid Me.sfRoster.Form, Me.sfRoster.Width, columns
    AlignGridHeaders Me, "sfRoster", "head", Array("txtCol0", "txtCol1")
    LayoutTableFooter Me, "sfRoster", "roster", "id"

Use CountedCaption("Transfer", 5) for one-line filter captions. Give a filter group
the width of its widest measured caption and use one row when space permits.

LayoutFilterRow measures the widest counted caption and gives each filter the same width. LayoutNumberedTimeline keeps circle and number dimensions fixed while distributing labels and connector lengths across the available width. See tests/README.md for the native Access checks.

## Licence

MIT. See [LICENSE](LICENSE).
