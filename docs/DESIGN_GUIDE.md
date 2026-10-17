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
  two in step (`KitVersionProblem`), and `KitEnsureBackend` at startup reconnects a copy whose
  data file has moved, with one Open window instead of a broken app.
- **One Personnel directory for every app.** People live once, in a shared, encrypted
  `Personnel_be.accdb` in its own restricted folder (`modPersonnel`), never in an app's data
  file. Apps link to it, sign the Windows user in with no questions, and keep only their own
  roles, linked by `person_id`. No SSN or DoD ID.
  `KitPersonnelStart` at startup finds or sets up the file and signs the user in;
  `KitPersonnelSignOut` behind the Sign out button shows the sign-in screen for the next person,
  and `KitPersonnelManage` opens the Personnel page for admins. Each site's Personnel file is
  locked with its own site key, made on the setup screen and kept in `Personnel.key` beside the
  file, where every app reads it. Nobody types it and it is never in the code, so the folder's
  permissions are what protect the file and its key.
- **Keep the kit separate from the app.** Reusable modules go in `kit/`. App modules go in
  `src/` and provide a few named hooks the kit calls.

---

## 2. Source layout and build

```
access/
  build.ps1                  drives Access via COM; loads modules, calls BuildAll
  kit/
    modules/*.bas            reusable: modUIKit, modKitUI, modDarkMode, modPng, modGridOut,
                             modGridCols, modGridFilter, modCsv, modCrypto, modKitUtil, modBackdrop,
                             modTrust (copied unchanged from the kit)
    forms/*.vba              kit code-behind: subGrid, dlgAppearance, dlgDatePick,
                             dlgMessage, dlgColumns, dlgFilter
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
- Line numbers: the build numbers its copy of every module and form's code
  (`VbaLineNumbers.ps1`, from the kit's `tools/`) so `Erl` names the failing line. Each
  line gets its own line number in the source file. The files in git are never numbered:
  edit the source, not the built front end. `-NoLineNumbers` builds without them.
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
  - The first form (the login form or `frmMain`, or the kit's optional splash screen
    `frmSplash` from `BuildSplash`, which opens it) through the kit's
    `BuildTrustCheck "<form>"`, not StartUpForm. Call it after anything that sets
    StartUpForm, since it removes it. It adds an `AutoExec` macro (replacing any the app
    had) that opens the form when the code is on, and opens `dlgTrust` instead when Access
    has turned the code off. When Access would still ask next time (not a Trusted
    Location, not a remembered trusted document), `dlgTrust` also opens over the form.
    `dlgTrust` has no code module or Open hook, so it works with the code off; it explains
    how to add the folder in the Trust Center. Code never changes Trust Center settings.
    The check reads the user's (HKCU) settings and policies only, not machine-wide (HKLM)
    policy locations, so on such a PC the page can show when it isn't needed.
    With the code off, Access first shows its own "active content is blocked" notice;
    `dlgTrust` follows it. Its Close macro keeps Save at Prompt, the only setting Access
    allows then.
  - `HideAccessSplash`, called just before `BuildTrustCheck`, hides Access's own splash.
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
| modAuth, modCrypto | Sign-in, permissions, salted SHA-256 via Windows CNG |
| modSchemaDDL | Table specs and back-end creation/upgrade (raise `SCHEMA_VERSION` + migrate) |
| modBuild | `BuildAll` entry point and startup options |
| modBuildUI, modBuildPages, modBuildPages2 | `Build_<form>` procedures, one per form, grouped by area |
| modUI | Run-time UI helpers: `NavTo`, badges/pills painting, `ShowTip`, the app's hooks (`BadgeColors`, `StatusVariant`) |
| modUIKit (kit) | Build-time control builders (section 6) |
| modKitUI (kit) | Run-time page helpers: `GridFit`, `GridSort`, `SelectFirstRow`, `ShowDetail`, `ClearDetail`, `PickDate`, `MarkDate`, `FederalHoliday`, `ToggleSwitch`, `ToggleFocus`, `SegmentClick`, `SetSegment`, `SegmentValue`, `ShowAlert`, `HideAlert`, `SetBreadcrumb`, `CrumbClick`, `SetStep`, `CurrentStep`, `OpenSection`, `SectionOpen`, `ShowEmpty`, `SearchText`, `SetSearch`, `SetRating`, `RatingValue`, `SetSlider`, `SliderValue`, `KitPopover`, `ToggleSecret`, `HideSecrets`, `KitMsgBox`, `KitToast`, `KitInputBox`, `CheckRequired`, `SetCount`, `StyleChip` |
| modGridOut (kit) | `PrintGrid`, `ExportGrid`: a table as the user sees it, to paper, CSV or Excel |
| modGridCols (kit) | Column chooser: `ChooseGridColumns`, `RestoreGridColumns`, `SetGridColumns`, and `BuildColumnsDialog` for `dlgColumns` |
| modGridFilter (kit) | Header filters: `GridFilter`, `SetGridFilter`, `SetColumnFilter`, `ClearGridFilters`, `GridFilterWhere`, and `BuildFilterDialog` for `dlgFilter` |
| modKitHelp (kit) | In-app help: `LoadHelpTopics` (help.txt into `zz_help`), `BuildHelpDialog`, `BuildHelpKey` (F1), `KitHelp` (the topic for what is on screen, search) |
| modTrust (kit) | `BuildTrustCheck` startup, `TrustNeeded`, `InTrustedLocation`, `ShowTrustHelp` (the `dlgTrust` page on trusting the folder) |
| modKitUtil (kit) | `FileExists`, `FolderExists`, `EnsureFolder`, `IsFormOpen`, `RunDialog`, `CenterOnApp`, `SetVisible`, `SetEnabled` |
| modBackdrop (kit) | `DimApp`, `UndimApp`: the dimmed backdrop behind modal dialogs |
| modDarkMode (kit), modPalette (gen), modTheme (gen) | Palette engine (section 7) |
| modPng (kit) | Draws anti-aliased PNGs for rounded pills, card corners, progress capsules and toggle switches |
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
Empty required fields are checked first, on the form's Save, with the kit's
`CheckRequired`: it outlines every empty one in `CLR_DESTRUCTIVE`, lists them all in one
message and puts the focus on the first, so the user fixes them in one pass.

**Session state** lives in `TempVars`. Nothing about the session is written to tables, and
saved queries can read `[TempVars]![X]`. Expose it through accessor functions such as
`CurUserId()` and `IsAdmin()`.

**Silent mode.** `g_Silent` is True during the build and tests. In that mode, `Notify`
prints to the debug window, raises on warnings and answers Yes to questions, and
`HandleError` raises the error again, so nothing can block the invisible Access.
Otherwise both show their message with the kit's `KitMsgBox`, never `MsgBox`: Windows
centres a `MsgBox` on the monitor, so on a large or second screen it opens outside
Access and Access waits behind it. For the same reason a question the user types an
answer to uses `KitInputBox`, never `InputBox`, with `masked:=True` for a PIN or password.

**Toasts for confirmations.** When the user only needs to know an action worked
("Saved", "Exported 24 rows"), show `KitToast msg, variant` instead of a message box:
a small toast in the bottom-right corner of `frmMain` that fades after a few seconds
and never takes the focus, so the user keeps typing. Variants are `success`, `info`,
`warning` and `destructive`. Anything the user must read, decide or fix stays a
`KitMsgBox` or an inline `lblMsg`; a toast disappears before some users read it.

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
   - The kit's modKitHelp loads it (`LoadHelpTopics`), builds the window
     (`BuildHelpDialog`) and the F1 key: `BuildShortcutKeys` (modKitKeys) writes the
     `AutoKeys` macro with F1 running `KitHelp()` beside Ctrl+S, Ctrl+N and Ctrl+F;
     `BuildHelpKey` writes F1 alone.
   - Topic keys match page keys, so Help (`=KitHelp()`), F1 and the sign-in Help link open
     the topic for whatever is on screen. A dialog, or a page with tabs, names its own
     topic with a `Public Property Get HelpTopic() As String` in its code-behind.
   - The search box over the topic list narrows it to the topics that mention the words
     and marks them in the topic shown, so users find an answer without knowing where
     it lives.
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
  hourglass, and shows a friendly message (`KitMsgBox`): "Something went wrong, so that action didn't
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

`LogError` and `RaiseAgain` add the failing line (`Erl`, from the build's line numbers)
to the procedure name: `pgAdmin.btnSave_Click line 42` is line 42 of
`src/forms/pgAdmin.vba`. `Erl` is per procedure, so for an error passed up by
`RaiseAgain` each trail entry and the entry point name the line of their own call.
Lines VBA won't take a number on (`Case`, `Else`, `ElseIf`, `Loop`, `Next`, a line
starting with a label) are logged at the numbered line before them, e.g. a failing
`Case 1: DoThing` at its `Select Case` line.

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

**Several users on one back end** (kit modKitLocks). A shared back end brings two
errors a single-user app never sees, and every app handles them the same way:
- **Locks** (3218, 3260, 3188 and the rest of `IsLockError`): someone is writing the same
  record this instant. Retry, then say "wait a moment, then save again". Run action
  queries through `KitExecute`, save bound forms with `KitSaveRecord`, and wrap your own
  DAO edits in a `LockWait` retry (the pattern is at the top of modKitLocks).
- **Write conflicts** (7787, 7878): someone saved the record after this user started
  editing it. Every bound form that edits shared data calls `KitNoteRecord Me` in its
  `Current` and `AfterUpdate` and `KitNoteEdits Me` in its `BeforeUpdate`, has a `Form_Error` that calls `KitFormError Me, DataErr, Response`, and
  saves with `KitSaveRecord`. On a conflict the user gets a `KitMsgBox` naming the fields
  they changed: Yes puts those back over the other user's record (they save again), No
  shows the other user's version. Saves Access starts itself (moving off the record,
  closing) keep Access's own Write Conflict window.

**Who did it** (kit modKitUser). Record people with `KitUser()`: the app's signed-in user,
else the Windows account from `WinUser()`. Never read `Environ("USERNAME")` for this:
anyone can change it before Access starts. Every table several users edit has
`created_by`, `created_at`, `updated_by` and `updated_at` (`AddStampFields`), filled by
`StampRecord Me` in the bound form's `BeforeUpdate` and by `StampSql()` or
`StampColumns()`/`StampValues()` in SQL writes.
On a form that also calls `KitNoteEdits` (modKitLocks), call `StampRecord Me` first, so a
write conflict puts the stamps back with the user's changes.
To recognise a person (sign-in, a person record), match `WinAccount()` (DOMAIN\user)
or `WinSid()`, never the bare `WinUser()`: two domains can both have a jsmith.

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
    background, and `OnClose = "=BackdropOnClose()"`, which removes a dim a VBA reset
    left behind (a dialog with its own `Form_Close` calls `BackdropOnClose` from it). A reset
    between `DimApp` and the dialog opening leaves the dim on until the next kit dialog opens or closes.
  - `NewReport recSrc, w`: Letter portrait, 0.5 in margins, `REPORT_W = 10800`, pop-up
    preview.
  - `SaveForm`, `Trace`
  - `BuildSplash [nextForm], [subtitle], [glyph]`: the startup splash screen (section 6.5).
- **Controls**
  - `Lbl`: transparent label.
  - `Txt`: input with `CLR_INPUT_BG`, a 1px `CLR_INPUT_BORDER` border and inner margins.
    A field box (not an `=` expression) and a `Cbo` show the focus ring: the border turns
    `CLR_PRIMARY` while the field has the focus (`FocusRing`, see §6.7).
  - `DateBox`: a date input, 120px wide, with its date left-aligned like text, and a
    30px outline calendar button 4px to its right that opens the kit's date picker
    (`PickDate`, `dlgDatePick`, built by `BuildDatePicker`): a month of day cells with
    previous and next arrows, the box's date highlighted, today in bold, and Today,
    Clear and Cancel. US federal holidays (`FederalHoliday`, on their observed day)
    are underlined in `CLR_DANGER_TEXT` and app dates (`MarkDate`, such as deadlines)
    in `CLR_WARN_TEXT`, which wins when a day is both; the tooltip names every mark.
    Days outside the month stay `CLR_MUTED_FG` with the underline and tooltip.
    Access's own picker is off: it shows only on focus, and
    `acCmdShowDatePicker` is refused once it is off. The pair is 154px
    (`DATEFIELD_W`), so it fits a half-width field column.
  - `SecretBox`: passwords and other sensitive values (PINs, SSNs, access codes). The
    text shows as asterisks (input mask `Password`, which also blocks copying while it is
    hidden), with a
    30px outline eye button 4px to its right: an eye (E890) to show the text, an eye
    struck through (ED1A) to hide it again (`ToggleSecret`). The form calls
    `HideSecrets Me` in `Form_Current`, so a value shown on one record is hidden on the
    next. Use it for every sensitive field. Masking only hides the text on screen:
    a password box is unbound, or hashed (`HashPassword`) in `BeforeUpdate`, before it
    is saved, and other sensitive values stay readable in tables, queries and exports.
    Single-record forms only: a continuous form shares one mask across its rows, so
    `SecretBox` raises error 5 when built on one.
  - `RO`: read-only text, borderless, transparent, no tab stop. Show read-only dates
    left-aligned (`TextAlign = 1`) so they line up with text values.
  - `Cbo`, `Lst`, `Chk`
  - `Toggle nm, src, x, y, [w], [h], [onChange]`: an on/off switch for a yes/no field
    (section 6.6). Use it for settings and flags that take effect at once; keep `Chk`
    for a choice saved with the rest of a form.
  - `Segmented nm, src, x, y, w, items, [h], [onChange]`: a segmented control for one of
    2 to 6 short options that changes what the page shows (section 6.6). Use a combo box
    for a longer or data-driven list.
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
  - `Banner nm, y, w, [vAnchor]`: the classification banner (section 6.5).
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
      the column or make it wrap. `s` marks the text matching the page's search box.
      A pill column draws the
      value as a rounded badge whose variant comes from the app's
      `StatusVariant(value)` (it gets a Variant, which may be Null, and returns a
      `BadgeColors` variant or "" for no pill). Map statuses there, never as colors.
    - The last column takes the remaining width and stretches with the window. A pill
      in the last column stops 8px further in, so it clears the table's scrollbar.
    - Rows are separated by a 1px line, and the selected row is highlighted in
      `CLR_PRIMARY_SOFT`.
    - The vertical scroll bar shows only when the rows overflow (`GridScroll`, the
      table's OnCurrent). The window's `Form_Resize` calls `GridScrollAll Me`, since
      a table anchored to the window changes height without a Current event.
    - Optional arguments, all off by default so existing tables don't change:
      - `zebra:=True` fills every other row with `CLR_ROW2`. The selected row still
        turns `CLR_PRIMARY_SOFT`.
      - `rowPx:=n` sets the row height in px. Use 38 for a roomy one-line table and
        56 for two lines. By default the height comes from `rowLines`.
      - `fit:=True` keeps the last column at its own width so `GridFit` can scale
        every column in proportion. Call `GridFit Me.sfList, "hdrList"` from the
        host's `Form_Resize`; it also moves the captions `GridHeaders` drew with that
        prefix. Columns never get narrower than designed.
    - **Search matches are marked.** On a page with a search box, give every column
      the search looks in the `s` option, and call `GridHighlight Me.sfList, term` with
      the box's text whenever it changes (with the filter). Each match shows in
      `CLR_WARNING_FG` on `CLR_WARNING_BG`, ignoring case; `""` clears the marks. An `s`
      column is a rich text box showing `GridMark([txtSrc<i>], [txtFind])`, where the
      hidden `txtSrc<i>` is bound to the field (a bare `[name]` in an expression would
      be the form's Name). `PrintGrid` and `ExportGrid` print and export that field.
      Use `s` on text columns only: not on pills, and not on dates or numbers, whose
      `Format` a rich text cell doesn't apply.
    - **Clicking a row opens it.** There are no `>` or Open buttons in rows: the row
      click selects the record and the host's `GridClick` shows it (in a detail panel,
      section 6.8, or a dialog).
    - All grids share one code-behind, `subGrid.vba`, which calls the parent's
      `GridClick(grid, dbl)`.
    - Each column's `Tag` holds its caption and design position, which `GridFit` and
      modGridOut read.
  - `GridHeaders prefix, cols, x, y, [parentCtl], [band], [w], [sortSf], [filterSf], [onFilter], [fit]`
    puts the captions above the grid in 8pt bold muted uppercase. `band:=True` draws them on a 22px `CLR_INSET`
    strip, `w` wide (default: the columns' widths), that stretches with the window.
    - Each caption is exactly as wide as its column's cells, so right-aligned and
      centred captions sit over their values. Pass the table's width as `w`: the last
      caption then reaches the table's right edge and stretches with it, like the last
      column. For a `fit:=True` table pass `fit:=True` here too, so the last caption
      keeps its column's width and `GridFit` scales it; without it the last caption is
      anchored to stretch and `GridFit` moves it past the column.
    - **Click to sort.** `sortSf:="sfList"` (the table's subform control) makes each
      caption sort the table by its column (`GridSort`): the first click sorts
      ascending, the next reverses. The sorted caption shows ▲ or ▼ after the text, or
      before it on a right-aligned column; the others show none. A column sorts by its
      field; give a fifth part, `"due|Due|80||due_date"`, when the cell shows formatted
      text (a date through `Format`) and the record source also returns the real
      field. A computed column (`"=..."`) sorts only with a fifth part.
    - The sort is the table's `OrderBy`, so `PrintGrid` and `ExportGrid` follow it. Rows
      with the same value keep id order. Sort a table in code through `GridSort` so its
      arrow stays right. A paginated table (`GridFooter`) goes back to page 1 in the new
      order (`RestartTablePages`).
    - **Click to filter.** `filterSf:="sfList"` (the same subform control) puts a funnel
      (Segoe MDL2 Assets Filter, 20px, muted) at the right end of each caption whose
      column shows a field; the caption's text stops 2px short of it, and the last
      caption takes its column's width. Clicking the funnel opens `dlgFilter`: the
      column's values (those the page's filter and the other columns' filters leave, plus
      those the column's own filter keeps even when no row shows them, in the column's sort order, `(Blanks)` last for Null and empty text), ten rows at a
      time, each with a check box glyph: ticked (`CLR_PRIMARY`) values show. The rows
      are kit-drawn labels, not a list box, because a list box shows its selection in
      Access's own colors whatever the palette. A click ticks or clears a value; the
      arrows and the mouse wheel scroll; Up, Down, Page Up, Page Down, Home and End move
      a `CLR_PRIMARY_SOFT` cursor row and Space ticks it. Select all ticks every value,
      or clears them, without scrolling; OK needs one ticked; Cancel changes nothing.
      Ticking every value removes the filter. While a column is filtered its funnel is
      white on a `CLR_PRIMARY` chip, like a selected filter chip, so it stands out by
      more than hue, and its tip says how many values show.
    - The table's `Filter` is the page's filter AND each column's (`[status] IN (...)`),
      so sorting, pagination, `PrintGrid` and `ExportGrid` follow it. A page that filters
      the table itself (search, chips) calls `SetGridFilter Me.sfList, where` instead of
      setting `Filter`; setting `Filter` directly replaces the header filters, and the
      funnels go back to muted at the next header call. `onFilter:="HeaderFiltered"`
      names a Public Sub on the form, called with the subform control's name after the
      user changes a filter, to update counts and the empty state; count with
      `GridFilterWhere`. Filters last while the page is open. Memo, OLE and attachment
      columns can't be filtered. The column chooser hides a hidden column's funnel too.
    - Build `dlgFilter` with `BuildFilterDialog` alongside the app's other dialogs.
  - **Column chooser** (modGridCols), for `fit:=True` tables only: in any other table
    the last column stretches with the window, so it can't move.
    - A Columns button (outline, top right of the page or card, next to Print and
      Export) calls `ChooseGridColumns Me.sfList, "hdrList"`. It opens `dlgColumns`,
      where the user hides, shows (button or double-click) and moves columns up and
      down, or resets to the design. At least one column stays shown.
    - Columns keep their design widths and `GridFit` scales them, so hiding one widens
      the rest. The captions move and hide with their columns. `PrintGrid` and
      `ExportGrid` print the columns in the order shown and leave out hidden ones.
    - The choice is saved per Windows user and table (registry, `<APP_NAME>\Columns`,
      named after the table's form). The host calls `RestoreGridColumns Me.sfList,
      "hdrList"` once the table's record source is set, and again after anything that
      shows every caption (such as a tab switch calling `SetVisible` on them). A saved
      layout that no longer matches the table's columns (a rebuild added or removed
      one) is ignored.
    - Build `dlgColumns` with `BuildColumnsDialog` alongside the app's other dialogs.
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
  still 9pt. Use one size per row. No other heights: no 24, 26 or 30px buttons. The
  exceptions are the `DateBox` calendar and `SecretBox` eye buttons, which take their
  input's 30px height.
- Vertical rhythm for stacked fields: 38px.
- Cards and KPI tiles are separated by 12px gaps.
- **Designer spacing.** The values in this section are the default. A designer who wants a
  tighter or roomier app calls `SetSpacing "compact"` (0.75) or `"comfortable"` (1.25), or a
  scale from 0.5 to 2, once before building. Write gaps and padding as
  `Sp(SP_CARD_GAP)`, `Sp(SP_PAGE)` and so on so they follow that choice; keep control sizes
  in `px()` (an input stays 30px). Without a `SetSpacing` call, `Sp(n)` equals `px(n)`.
- A KPI tile is 78px tall: a surface box with a 1px border, a 4px colored accent strip
  along the bottom, a 10pt caption and a 24pt number. A transparent `Link` over the whole
  tile makes it clickable.
- Dialog buttons sit bottom-right in the order **Cancel (outline)** then **primary or
  destructive** action. Cancel sets `.Cancel = True` and the primary action sets
  `.Default = True`. The inline error label sits bottom-left.
- The exception is the kit's stand-ins for Windows dialogs: the message box, the input
  box and the date picker (`dlgMessage`, `dlgInput`, `dlgDatePick`). They use 30px
  buttons, and the message and input boxes keep the Windows order (OK, then Cancel), so
  they read like the dialogs they replace.
- Header actions sit top-right, anchored right: the primary action rightmost, then
  outline, then destructive.

### 6.5 Application shell

- `frmMain` is a single form containing:
  - A classification **banner** at the top and bottom (`Banner`, 20px, full width,
    anchored). It shows the app's selected classification (`AppClassification`), which
    also limits the labels the app permits (`ClassificationPermitted`,
    `modKitClassification`). Build the banners last so they sit above the other controls,
    and shift the sidebar and page host down by one banner's height.
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
  - Before it swaps or closes the page (`QueueNav`, Appearance, `Form_Unload`), the shell
    asks `frmMain.CanLeavePage`. A page with unsaved work defines a public `CanLeave`
    (pgList: `CanLeave = SaveDetail(Me, "sfDetail")`), and the page stays when it returns
    False. Sidebar items are labels, so clicking one never takes the focus out of the
    page: without this, Access would drop the page's unsaved record.
  - `RefreshNav` restacks the visible items without gaps and highlights the current one.
- **Reopen**: `frmMain.Reopen pageKey` closes and reopens the shell on the same page,
  which is how a palette switch repaints everything. An `m_reopen` flag keeps `Form_Close`
  from signing the user out.
- **Dialogs**:
  - `RunDialog(name, openArgs)` opens the form with `acDialog` and returns True when the
    dialog set `TempVars("DialogResult") = "ok"`.
  - While a dialog is open the app behind it is dimmed, like a web modal: a see-through
    layer of `CLR_SIDEBAR` (45%) over the Access window. `RunDialog`, `KitMsgBox` and
    `KitInputBox` do this themselves and remove the dim when the dialog closes, or fails
    to open.
    `RunDialog name, openArgs, False` opens one undimmed. A modal form opened any other way
    calls `DimApp` before it and `UndimApp` after, on the error path too. The date picker
    is a small drop-down, so it isn't dimmed.
  - A dialog that wires its own close handler (`Ev F, "OnClose"`) replaces
    `=BackdropOnClose()` from `NewDialog`, so its `Form_Close` calls `BackdropOnClose` itself.
  - Other results come back in named TempVars.
  - Dialog input comes through `OpenArgs`.
  - Every dialog's `Form_Load` ends with `CenterOnApp Me`, after anything that changes
    its size. It centres the dialog over the Access window, on that window's monitor,
    and pulls it back inside the monitor's work area. `AutoCenter` alone can open a
    dialog high or on another monitor.
- **Splash screen** (optional): the app starts on `frmSplash` (`BuildSplash`), a borderless 480 x 280px
  pop-up in `CLR_SIDEBAR` with the app's MDL2 glyph, its name in 20pt white, a subtitle,
  and a footer under a `CLR_SIDEBAR_LINE` hairline: a status line on the left and the
  version on the right. `ThemeForm` gives it the user's palette. Once it has painted, its
  timer opens `nextForm` (the login form or `frmMain`); that form's Load does the start-up
  work (version checks, linking) and calls `SplashStatus "..."` to say what it is doing.
  The splash closes when the form is open, and stays at least 1.2 seconds so it doesn't
  flash. `KitMsgBox` closes it first, so a start-up error is never hidden behind it.
  A plain `MsgBox` or a file dialog doesn't close it, so call `CloseSplash` before one
  (for example a back-end file picker at start-up).
  `HideAccessSplash` writes the 1 x 1 pixel `<database name>.bmp` that replaces Access's
  own splash. It is a build output: ignore it in git and ship it beside the front end.
- **List subforms** (continuous forms of cards or rows) call back to their host with
  `Me.Parent.<PublicProc>`.
- **New version notice**: once the app has opened (after sign-in checks, before the
  first page), call `CheckForUpdate APP_VERSION, <release folder>`.
  - The release folder is wherever users copy the app from. The release script writes
    `version.txt` there, holding the release's version (`1.2.0`) on its first line.
  - A newer version gets one information `KitMsgBox` naming both versions and the folder,
    once per session. The app keeps running; nothing is copied for the user.
  - Pass `advice` (third argument) when "copy the new version from the folder" is the
    wrong next step, for example when the app's data lives inside its own file.
  - A blank folder or a missing file skips the check. A read failure is logged with
    `LogError` and never shown, so an offline share can't stop the app opening. An
    unreachable network folder can still delay opening by about 20 seconds.

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
  - Hide a pill with `SetVisible`, never `.Visible = False`: `SetVisible` hides its
    `_bg` image too, so closing or clearing a card leaves no empty pill behind.
    Showing it again restores a bound image; repaint a painted pill to show its image.
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
- **Toggle switch**: `Toggle` draws a 44 x 24px rounded track with a round knob
  (`TogglePng`): `CLR_PRIMARY` with the knob on the right when on, `CLR_MUTED_FG`
  with it on the left when off, the knob `CLR_SURFACE`, so the track and knob keep well
  over 3:1 against the surface and each other in Navy and Midnight. Center the track on its row: in a
  30px input row it sits 3px down. Access has no switch, so it is a hidden check box
  `<nm>_val` bound to the field (the value), the track image `<nm>_bg` bound to it, and a
  transparent button `<nm>` on top that takes clicks, the Tab stop and Space
  (`ToggleSwitch`). A transparent button draws no focus, so a `CLR_PRIMARY` ring
  (`<nm>_ring`, 2px wide, 2px outside the track) shows while it has the focus, and an input's
  `FocusRing` comes off as it does for a button; leave 4px clear around the track for it. A value set in code fires no `AfterUpdate`, so pass
  `onChange`, a Public Sub in the form's code-behind, to act on a change (for example,
  save the record at once). A misspelled `onChange` shows error 438 on the first click.
  Hide it with `SetVisible` on `<nm>`, and the track goes too; `SetEnabled` on `<nm>`
  stops clicks and the Tab stop but doesn't dim the track. Use it on single forms only:
  in a continuous form the ring shows on every row.
- **Segmented control**: `Segmented` lays 2 to 6 options side by side on a capsule track
  in `CLR_SECONDARY_BG`, 32px tall by default (28px beside a card title), sharing its width
  equally; give each option at least 40px, and its caption room at 9pt. The picked option sits on a
  `CLR_SURFACE` capsule thumb inset 3px, its caption bold `CLR_FG`; the others are
  `CLR_MUTED_FG`, like a web toggle group. It is a hidden text box `<nm>_val` holding the picked key
  (bound to a text or number field, or unbound and starting on the first item), the track
  image, one thumb image per option whose expression shows it only while that option is picked,
  so a bound control follows the record with no code, the captions with a conditional
  format, and a transparent button per option (`SegmentClick`) that takes clicks, a Tab stop
  and Space. The focus ring is the `Toggle`'s, around the whole track, so leave 4px clear.
  Act on a change in `onChange`; read the value with `SegmentValue` and set it with
  `SetSegment`, which calls no `onChange`. Single forms only.
- **Inline alert**: `Alert` is a message that stays on the page while it applies (an item
  past its due date, a read-only record, what a page's filter hides), where a toast would
  fade before it is read and a message box would interrupt. A panel tinted in the variant's
  soft color (`info`, `success`, `warning`, `destructive`) with a 4px accent bar and icon in
  `KitToast`'s colors, a 10pt bold title and the 9pt message in `CLR_FG`, 16px in from the
  left and 44px for the text. Give it the height of the lines it may need (64px fits a title
  and two lines); an untitled alert under 48px is compact, one line centered, for a
  page's header row. Set it at run time with `ShowAlert` (it repaints in the current palette)
  and `HideAlert`; `dismiss:=True` adds a close mark that hides it. Keep one alert per
  page area, and say what to do, not only what is wrong.
- **Breadcrumb**: `Breadcrumb` shows where a page sits (All items > In progress > Pump
  overhaul) on one 20px line, in the subtitle's place under the page title: 9pt
  `CLR_MUTED_FG` links, 7pt chevrons (MDL2 E76C) 16px wide between them, and the current
  page last in `CLR_FG`, not a link. Use it where a page is reached by going down levels;
  a page reached from the sidebar needs none. Build as many slots as the deepest path and set
  the path with `SetBreadcrumb` whenever the page's place changes. A path wider than the
  trail keeps its first and last crumbs and folds the middle into an ellipsis, as web
  breadcrumbs do.
- **Step indicator**: `Stepper` shows progress through a task done in steps (a new-record
  wizard, an approval route) above the step's own content, usually in a card with Back and
  Next at the bottom right. 28px circles centered in equal slots, 9pt captions 8px below,
  2px connectors 8px clear of the circles. Done steps are filled `CLR_PRIMARY` with a check
  mark (MDL2 E73E) and fill the connector after them; the current step is filled with its
  number and a bold `CLR_FG` caption; later steps are a 2px `CLR_INPUT_BORDER` ring with a
  `CLR_MUTED_FG` number. A click on a done step goes back to it (the `handler`); later steps
  can't be skipped to. Keep captions to one or two words.
- **Accordion**: `Accordion` splits a long page into sections the user opens as needed:
  40px headers with a 10pt semibold title and a chevron at the right (right when closed,
  down when open), a `CLR_BORDER` hairline under each, and the open section's body (a
  subform) 7px below its header and 8px above the next. Opening or closing moves the
  sections below, so keep the accordion last in its column. Use `oneOpen` when the sections
  are alternatives; leave it off when the user may compare them. Open the section most
  users need at first.
- **Avatar**: `Avatar` shows who owns or is assigned a record as their initials in a 32px
  circle (24px in a dense row), bold at 30% of the diameter in points, beside or instead of
  the name. The tone comes from the name, so a person keeps one color on every page; tones
  are the soft badge pairs, so they read in every palette. Initials are the first letters of
  the first and last words (ETN1 Rivera is ER). It is bound, so it follows the record on a
  continuous form with no code.
- **Dropdown menu**: `KitMenu` puts a record's or page's secondary actions (Duplicate,
  Archive, Delete) behind one ghost button, the MDL2 More glyph (E712, 32 x 28px) at the
  right of a card's title row, instead of a row of buttons. The menu is a `CLR_DIALOG`
  pop-up with a `CLR_INPUT_BORDER` hairline, 160 to 320px wide to fit its captions, 32px
  rows with a 10pt MDL2 icon and a 9pt caption, `CLR_OUT_HOVER` behind the row under the
  mouse or keys, and 9px separator rows with a `CLR_BORDER` line. It opens 4px under the
  button, left edges lined up, and flips right-aligned or above to stay on the monitor. Put
  destructive actions last, after a separator, as `!Caption` (`CLR_DESTRUCTIVE`), and
  confirm them with `KitMsgBox`; show an action that doesn't apply as disabled (`~key`)
  rather than leaving it out, so the menu keeps its shape. Keep it to about seven items.
- **Stat tiles**: a dashboard opens with a row of `StatTile`s, 112px tall, sharing the
  width with `Sp(SP_CARD_GAP)` between them: a 9pt muted label, the 22pt bold value, an
  MDL2 icon at the top right in `CLR_MUTED_FG`, and below the value a 64 x 20px change
  badge (success up, destructive down, secondary flat) with an 8pt muted caption saying
  what it compares with. Use `upIsGood:=False` where a rise is bad (overdue, cost). Four
  tiles fit a 1000px page; keep labels to two or three words.
- **Search box**: `SearchBox` is the page's search field: a 28 to 32px frame in the input
  colors with a magnifier (MDL2 E721) 8px in, the text, and a clear mark (E711) 8px from
  the right that shows once there is text. Filter as the user types (its handler), and keep
  the placeholder to what can be searched ("Search items or owners...").
- **Empty state component**: where a list may be empty for a reason the user can act on
  (a filter, nothing created yet), use `EmptyState` rather than a bare label: an icon that
  names the thing, a title saying what is missing, a line saying why or what to do, and a
  button that does it (Clear filters, New item).
- **Rating**: `Rating` shows and sets a 0 to 5 score as 24px MDL2 stars, filled
  `CLR_AMBER_500` and outlined `CLR_INPUT_BORDER`; a second click on the value clears it.
  Use `readOnly` to show a score on table rows.
- **Slider**: `Slider` picks a number from a short, even range (rows per page, a
  percentage in steps of 10) where the exact value matters less than its place in the range.
  The knob is 20px with a 2px `CLR_PRIMARY` ring on a 6px track; show the value at the
  right with `showValue`. Use a text box for exact or wide-ranging numbers.
- **Popover**: `KitPopover` explains a field or setting in a sentence or three, next to
  it, from a muted MDL2 help mark (E9CE) beside the caption. 280px wide, a 10pt bold title,
  9pt text, up to 8 lines. It closes on a click anywhere or Esc. Keep page-level help in the
  page title's "?" (`TitleTip`).
- **Keyboard and screen readers**: everything a mouse can click, a keyboard can reach.
  Access labels can't take the focus, so every clickable label (a link, a close mark, a
  crumb, a help mark) gets a `KeyButton` over it, and every transparent button (switches,
  segments, stars, slider steps, accordion headers) gets a `Caption`: it isn't drawn, but
  it is the name a screen reader reads. A focused control always shows it: a focus ring
  (inputs, `Toggle`, `Segmented`, `Rating`, `Slider`) or `FocusMark` (labels and accordion
  titles turn `CLR_PRIMARY` and, for text, underlined). Outlines that carry meaning (an
  empty star, a step not reached) use `CLR_MUTED_FG`, which keeps 3:1 against the surface;
  `CLR_INPUT_BORDER` is for decoration only.
- **Empty states**: a centered muted 10pt label shown in place of an empty list. Hide the
  list's subform while it shows (`SetVisible`): a subform always covers a label.
- **Toasts** (`KitToast`, `dlgToast`): a 360px borderless pop-up on `CLR_SURFACE` with
  a 1px `CLR_INPUT_BORDER` edge, a 4px accent bar on the left and a Segoe MDL2 icon in
  the variant's color (success `CLR_GREEN_600`, info `CLR_PRIMARY`, warning
  `CLR_WARN_TEXT`, destructive `CLR_ERROR`), then the message in 10pt `CLR_FG`, up to
  3 lines. It sits 16px in from `frmMain`'s bottom-right corner, holds for 3 seconds,
  fades over about a third of a second and closes; a click closes it at once.
- **Search box placeholder**: `Format = "@;""Search by ..."""`.
- **Keyboard shortcuts**: every app has the kit's Ctrl+S (save), Ctrl+N (new), Ctrl+F (find)
  and F1 (help) through `BuildShortcutKeys`. A page that has a New button answers
  Ctrl+N in its `Shortcut` function; name the search box `txtSearch` so Ctrl+F finds it;
  put the key in the button's tooltip ("Save (Ctrl+S)"). These keys replace Access's own
  Ctrl+S (save the object's design), Ctrl+N (new database) and Ctrl+F (Find dialog).
- **Files**: never ask users to type a path. Use `KitSaveAs`, `KitPickFile` or `KitPickFolder`
  (Windows' own windows, starting in the folder last used for the same purpose), and
  `ExportGridAs` for a table's Export button.
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
- **Focus ring**: `Txt` and `Cbo` set `OnGotFocus`/`OnLostFocus` to `=FocusRing(...)`,
  which paints the border `CLR_PRIMARY` and then gives back the color it found (so a
  required field's red outline survives tabbing through). Only `BorderColor` changes: a wider
  border would eat into the text area and shift the text, and an extra ring control
  would have to be shown and hidden on every focus change (flicker). Locked fields,
  read-only forms and continuous-form rows get no ring (one control's border is every
  row's). Only one field is ringed at a time, and `SubF` sets the subform control's
  `OnEnter`/`OnExit` to `FocusRingEnter`/`FocusRingClear`: Access fires no LostFocus for
  a field inside a subform when the focus moves out to the main form, and no GotFocus
  when it comes back to the same field. `Btn`, `Chk`, `Lst` and the focus sink clear
  the ring on GotFocus (`FocusRingClear`), because the subform's Enter re-rings its
  last field before a click on its Save button lands. Setting these events later
  (`Ev c, "OnGotFocus"`) replaces the ring, so such a handler calls
  `FocusRing Me, "<name>", True` (`False` on LostFocus), or `FocusRingEnter` /
  `FocusRingClear` for a subform control, itself.
- Read-only row text with no click handler is set to `Enabled = False` +
  `Locked = True`. That keeps its normal look while it can never take focus.
- **Disabled controls ignore theme colors.** Lock controls instead of disabling them when
  they must stay themed.
- **Anchored widths**: an anchored control's `Width` still returns the design value. Use
  `form.InsideWidth - form.Width` to get the extra width.
- **Form width only grows**: moving a control past a form's right edge at run time widens
  the form, and the anchored controls' design sizes grow with it. Access never narrows
  them again. To narrow, move that control back, take the same amount off each anchored
  control (`Width` if anchored both sides, `Left` if anchored right), then set the form's
  `Width` (GridFit does this for its captions' host).
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
- The detail form checks its required fields in `Form_BeforeUpdate` (`Cancel = Not
  CheckRequired(Me, ...)`) as well as on Save, so no save path skips them. A refused
  save keeps the record in the panel: `ShowDetail` and `ClearDetail` return False and
  change nothing, `GridClick` then puts the list's selection back on the panel's record,
  and a filter chip calls `SaveDetail Me, "sfDetail"` first and stops when it is False.
  The page's `CanLeave` does the same for the shell (section 6.5).
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
   - Pop-up forms (dialogs) also get a themed title bar and border in every palette,
     light included: `ThemeTitleBar` sets the caption to `CLR_DIALOG`, its text to
     `CLR_FG` and the border to `CLR_BORDER` through `DwmSetWindowAttribute`, and turns
     on dark caption buttons in dark palettes. `SetPalette` reapplies it to pop-ups that
     are open. The colors need Windows 11; Windows 10 (20H1 and later) honors only the
     dark-buttons setting, so dark palettes get its plain dark title bar there. Rejected
     attributes are ignored; nothing is shown or logged.
   - Code that sets a color at run time **must** wrap it: `lbl.ForeColor = Th(CLR_MUTED_FG)`.
   - The choice is saved per Windows user with `SaveSetting APP, "UI", "Theme",
     paletteKey`. With no saved choice, it follows the Windows app mode (light or dark).
   - Switching palettes: Appearance in the sidebar opens `dlgAppearance` (one row per
     palette with three swatches: sidebar, accent, page). Picking a row calls
     `SetPalette`, which saves it and calls `frmMain.Reopen`.
5. **Reports** use `PRINT_*` colors (black, grays, white paper) and are never repainted.

### 7.2 Rules

- Never write `RGB()` or a numeric color in VBA. Use a `CLR_` role, or a `FIXED` constant
  for semantic colors that must not change, such as status colors and classification
  banners.
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
- classification banners: green 0,122,51 (UNCLASSIFIED), purple 80,43,133 (CUI), blue 0,51,160
  (CONFIDENTIAL), red 200,16,46 (SECRET), orange 255,103,31 (TOP SECRET), yellow 247,234,72
  (TOP SECRET//SCI), with black text on the orange and yellow (`CLR_CLASS_TEXT_DARK`) and
  white on the rest
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

1. Copy `kit/`, `examples/build.ps1`, `tools/VbaLineNumbers.ps1` and `tools/make_themes.py`
   from this repository.
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
   - Set startup options, opening the first form through `BuildTrustCheck`.
   - Return `"OK"` or the step that failed.
5. Build the shell: collapsible sidebar with `NavItem`s, `sfPage` host, deferred
   `QueueNav`/`ShowPage`/`RefreshNav`, and `PageKey`/`Reopen`.
6. For each page, write a `Build_pg<Name>` using `PageHeader`, cards, grids and buttons,
   and a `src/forms/pg<Name>.vba` where every handler uses the standard error handler.
7. Copy the dialogs: `dlgAppearance`, `BuildMessageDialog` (for `Notify` and `HandleError`),
   `BuildToast` (for `KitToast`),
   `BuildInputDialog` (for `KitInputBox`),
   `BuildDatePicker` (for date fields), and the help window: `LoadHelpTopics` with
   `help.txt`, `BuildHelpDialog`, and `BuildShortcutKeys` in the startup step for Ctrl+S, Ctrl+N,
   Ctrl+F and F1 (`BuildHelpKey` instead for F1 alone). Optionally add the splash screen with `BuildSplash` and
   `HideAccessSplash`, and pass `frmSplash` to `BuildTrustCheck` as the startup form.
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
