# Kit demo

A small Access app built straight from this repository, to see the kit working and to
test changes to it. It has one local table of fictional items and six pages:

| Page | Shows |
|---|---|
| Cards and progress | `Card` (rounded and square), `Progress` at 0, 37 and 100 percent, a table inside a rounded card; the square card's buttons open `KitInputBox`, plain and masked; the rounded card's Primary button shows `KitMsgBox` and Outline shows `KitToast`; a `Segmented` control (All, Open, Done) beside the table's title filters it |
| Tables | `BuildGrid` with every option (header band, zebra rows, status pills, 38px rows, proportional fit through `GridFit`) next to a default table, both sorted by clicking a column caption; filter funnels on the first table's captions; the column chooser (Columns, top right) for the first table; `TabStrip` and chips with counts; search and a category filter to the right of the chips |
| Steps and sections | A `Stepper` in a card with Back and Next (click a check mark to go back) and a one-open `Accordion` of three sections, the first holding the items table and the last Preferences: a `Slider` (rows per page, with a `KitPopover` help mark) and a `Rating` |
| Dashboard | Four `StatTile`s (open, past due with down as good, complete, average done) with change badges, and the flagged items in a card with an `EmptyState` when none are flagged |
| List and detail | The list/detail pattern (`SelectFirstRow`, `ShowDetail`, `ClearDetail`), click-to-sort captions, `PrintGrid` and `ExportGrid` in the list's sort; the Due field is a `DateBox` whose date picker marks federal holidays and every item's due date (`MarkDate`); a compact `Toggle` switch (Flagged for review) on the detail panel's Save row; a `SecretBox` (the item's access code); an `Avatar` with the owner's initials beside the Owner field; Save confirms with a toast; a compact warning `Alert` in the header row counts open items past their due date (its close mark hides it); a `Breadcrumb` below the title (All items > status > item) whose crumbs pick the chip; the list's `SearchBox` (magnifier and clear mark, as on Tables) and its `EmptyState` when nothing matches; the detail card's More button (⋯) opens a `KitMenu` dropdown (Mark complete, Flag, Delete) |
| Form designer | The kit's form designer (`BuildDesigner`, modKitDesigner), drag any of 22 kit elements (card, button, label, progress, text, date, secret, toggle, segmented, slider, rating, stepper, search, stat tile, alert, avatar, empty state, table, chips, tab strip, breadcrumb, accordion) from Elements onto the canvas (or click it, then the canvas), drag to move, drag the corner handle to resize, all on an 8px grid; double-click an element, or Edit under Selected, for its settings in a dialog (`dlgDzProps`) that checks them against the kit builder's rules; Delete removes it; a layout saves under a name to `kit_layout` and `kit_layout_item` and opens again from the list; Form sets page or dialog, size and record source, elements bind to its fields and anchor to the window's edges, Tab order (`dlgDzTabs`) sets the Tab key's order; Generate builds it as a real form with the kit's builders, shows it (`dzPreview`), and writes the same calls as `designer\Build_pg<Name>.txt` with a `pg<Name>.vba` code-behind stub |

Help in the sidebar, or F1, opens the kit's help window (`KitHelp`) on the page's topic, with a search box over the topic list; the topics are in `src\help\help.txt`. Ctrl+F goes to a page's search box, Ctrl+S saves a changed record, and Ctrl+K opens the command palette over the pages and the items opened lately (`BuildShortcutKeys`, `KitPalette`). The sidebar collapses to icons, and a right-click on a list row opens its actions. Appearance in the sidebar switches palettes. The demo opens on the kit's splash screen
(`BuildSplash`), which opens the main window and closes; `HideAccessSplash` replaces
Access's own splash with `KitDemo.bmp`.
Folder trust in the sidebar opens the kit's trust page (`dlgTrust`), which also opens by
itself at startup when the demo's folder isn't a Trusted Location.

## Build and open

```powershell
cd examples\demo
.\build.ps1            # new KitDemo.accdb with seed data
.\build.ps1 -UiOnly    # rebuild code and forms, keep the data
```

The code loaded into KitDemo.accdb has line numbers so the error log names the failing
line (`-NoLineNumbers` leaves them out). Edit the files here, not the built copy.

Then open `examples\demo\KitDemo.accdb`. The build stops and prints any Visual Basic or
Access dialog it meets (a compile error, an error message) instead of hanging.

## Release

`release.ps1` builds a locked showcase copy from a tag `kitdemo-v<version>`, in a clean
export in TEMP, so this folder and its KitDemo.accdb are never touched:

```powershell
git tag kitdemo-v1.0.0 <merged commit>
.\release.ps1 -Version 1.0.0     # writes <repo>\Releases\KitDemo\v1.0.0
```

It runs the kit checks (`tests\native-smoke.ps1`), builds the demo, stops if the build
logged an error, and writes `KitDemo.accde` (the locked app), `KitDemo.accdb` (the full,
editable program), `KitDemo.bmp`, `design\` (every form as JSON and a SQL Server script
for the tables; only release builds make these), `RELEASE-NOTES.md` and `SHA256SUMS.txt`.
Released copies go in the private navy-access-programs repository under `KitDemo/v<version>/`.

## What to check after a kit change

Click through every page in one light palette (Navy) and one dark palette (Midnight),
watching for Visual Basic error dialogs as well as the error log:

- Splash: opening `KitDemo.accdb` shows no Access logo, then the splash centred in the
  palette's sidebar color with the icon, "Kit Demo", the subtitle, "Build <stamp>" on the
  right and a status line; it closes once the main window is up. The splash takes the palette
  saved last, so to see it in another palette pick one in Appearance, then close and reopen
  the file.
- Help: on each page, Help in the sidebar and F1 open the help window on that page's
  topic (Tables opens Tables; Choosing columns is indented under it). Typing "funnel" in the
  search box leaves only Tables and marks the word; "zzz" shows "No topics match your
  search."; Enter opens the first match; Esc closes it. With help open, F1 keeps its topic.
  Check it in Navy and Midnight: the topic text, its marks and the list stay readable.
- Cards: corners are round in every palette; hovering the buttons and cards doesn't flicker.
- Progress: 0 shows only the track, 37 is a third full, 100 is full.
- Toasts: each click on Outline shows the next variant (success, info, warning,
  destructive) 16px in from the window's bottom-right corner; it fades after 3 seconds;
  a click on any part of it closes it; a toast while one shows replaces it. On List and
  detail, click into Notes, then Save, then press Shift+Tab while "Item saved." shows:
  the focus moves to Notes, so the keyboard never went to the toast.
- Input box (square card): Ask opens with "Sailor" selected, OK reports the text (an
  empty box reports ""), Cancel and Esc report Cancelled; Ask masked shows * for each
  character and reports only the length.
- Tables: the band, zebra rows, pills and 38px rows show; the selected row is highlighted;
  widening the window widens every column of the first table and moves its captions; the
  chip and tab counts follow the search and category. Columns opens the column chooser:
  hiding a column widens the rest and hides its caption, Move up/down reorders columns and
  captions together, Reset puts the design back, Cancel changes nothing, and the last shown
  column can't be hidden. The choice is still there after closing and reopening the demo,
  and clicking a caption still sorts after columns are moved.
- Header filters (first table): each caption shows a muted funnel at its
  right end. Clicking it lists the column's values (Due in date order, `(Blanks)` last);
  clearing values' ticks hides their rows, the funnel turns white on a primary chip, and the chip
  and tab counts and the empty state follow. Select all reads Clear all when every
  value is ticked and doesn't scroll the list; the arrows, mouse wheel and keyboard
  (arrows, Page Up/Down, Home, End, Space) work; OK is disabled with none; Cancel changes nothing. The filters
  stay when a chip, the search or the category changes, and when a caption sorts.
  Hiding a filtered column in the column chooser hides its funnel.
- Focus ring: clicking or tabbing into the search box, the category filter or a detail
  field turns its border the palette's primary color, and leaving it restores the border,
  with no flicker and the text not moving.
- Dialogs: Appearance and the card buttons' messages open centred over Access, on its
  monitor, with Access restored, maximized, or moved partly off the edge of a screen.
- List and detail: the first row opens on load and after each filter; a click shows that
  row; editing and saving updates the list; a search with no match shows both empty
  states; Print previews and Export writes exactly the rows shown, with the captions.
- Folder trust: the page shows this folder's path and the steps, nothing is cut off, and
  Close (or Esc) closes it. Opened from a folder that isn't a Trusted Location, with the
  code turned off, Access first shows its own "active content is blocked" notice, then
  the demo shows the page on its own: its folder box says to look under File, then Info,
  Enable Content works while the page is open, and Close quits Access. After Enable Content, Access remembers the file as
  trusted, so the page doesn't open over the main window.
- Date picker (Due's calendar button): October 2026 shows Columbus Day (12th) underlined in
  the danger color and item due dates (1st, 4th, 7th...) in the warning color, with
  tooltips naming them; July 2026 marks Friday the 3rd as Independence Day (observed).
  Changing an item's due date and saving moves its mark.
- Access code: opens as asterisks; the eye shows the code and turns into a struck-through
  eye, a second click hides it; the next item opens hidden again.
- Toggle (Flagged for review, left of Save in the detail panel): clicking the track or its
  label, or Tab to it and Space, slides the knob and saves the record at once; the track is
  the palette's primary color when on and its muted text color when off; a ring shows while
  it has the focus, inside the panel. The table gained a column, so build without
  `-UiOnly` once.
- Form designer: dragging Card from Elements shows an outline that follows the pointer over
  the canvas and snaps to the grid; releasing places the card there, selected; releasing off
  the canvas places nothing. Clicking Button and then the canvas places a button there. A
  selected element can be dragged anywhere on the canvas but not off it, and resized from its
  corner handle (a button stays 24 to 48px tall, a progress bar 4 to 16px). Moving and
  resizing don't flicker the page, and labels and buttons stay on top of cards. Delete
  (the button, or the key with the canvas focused) removes it; the arrow keys move it. Save
  with a name, New, then pick the name under Open a layout: the same elements come back.
  Leaving the page with unsaved changes asks to save them. Generate opens a window titled
  "Preview: <name>" where every element sits exactly where it was on the canvas, now as the
  real kit control (a rounded card, a themed button, a progress capsule), and writes
  Build_pg<Name>.txt and pg<Name>.vba to a designer folder beside KitDemo.accdb.
  Check it in Navy and Midnight.
- Form designer elements: every one of the 22 tiles places its element (the ones past
  Progress as an outlined box with a muted caption: the options of a segmented control,
  stars for a rating, initials for an avatar). Double-click each (or Edit) for an "Element
  settings" dialog with only the rows that kind uses; OK applies them and the box redraws;
  bad settings (one segmented option, a slider step that doesn't divide its range, 2 stars,
  a stepper's current step past the end) are refused with a message and change nothing.
  More segmented options or more stars widen the element. Generate with one of each shows
  them all as real kit controls in the preview (they work there but call no code), and
  pg<Name>.vba has a handler for each.
- Form designer data and behaviour: Form opens "Form settings" with title, width, height,
  page or dialog and record source (demo_item is in the list). A size smaller than the
  elements reach, or past the canvas, is refused; a smaller size shrinks the outlined area
  and elements can't be dragged past it; Dialog paints it in the dialog colour. With
  demo_item as the record source, a text box's settings list its fields under Field (bind it
  to name), and a toggle's (flagged); Generate shows the first item's name and flag, and the
  preview can't change them. On a page, anchor a button Top right and a card Stretches across
  and down: resizing the preview window keeps the button at the right edge and grows the
  card. Tab order lists the focusable elements in reading order; Move up and OK, then
  Generate: Tab in the preview follows the new order. Saved, New and reopened, all of it
  comes back. A dialog layout generates dlg<Name>, titled with its title. Check it in Navy
  and Midnight.
- Form designer tables, chips, tabs, breadcrumb and accordion: a Table's settings list the
  tables and queries under Lists; demo_item with the columns left empty fills them with its
  first fields, and a table or query without an id field is refused. A second Tab strip or
  Breadcrumb is refused with a toast. Chips widen with their captions, an accordion grows
  with its sections and body height, and settings that don't fit the form are refused.
  Generate: the table lists every demo_item row with its column captions, Stretches across
  widens it with the preview window; the first chip and tab are on, and clicking another
  chip or tab moves it; the breadcrumb's crumbs do nothing, and the current one reads in
  CLR_FG in Midnight too; an accordion header opens and closes its section, whose body is
  the page's color in Midnight, not white. Tab goes
  through the breadcrumb and the accordion headers in reading order and skips the table,
  whose rows take clicks. Build_pg<Name>.txt starts with the table's BuildGrid and the accordion's BuildBlank, and pg<Name>.vba
  has chips1Chip, tabs1Tab, crumb1Crumb, GridClick and a Form_Load. Check it in Navy and
  Midnight.
- Dialogs: Appearance in the sidebar and the buttons on Cards and progress dim the whole
  Access window behind their dialog; the dim goes when the dialog closes, and the dialog
  itself is not dimmed.

## Files

```
build.ps1                 builds KitDemo.accdb (stages kit\forms and src\forms together)
release.ps1               builds a locked release from a kitdemo-v<version> tag
src/modules/modDemo.bas   the names the kit calls, and the pages' shared helpers
src/modules/modDemoBuild.bas  BuildAll, the table and seed data, Build_<form>
src/modules/modPalette.bas    generated: python ..\..\tools\make_themes.py src
src/themes/light/modTheme.bas generated, as above
src/forms/*.vba           code-behind for frmMain and the pages
```
