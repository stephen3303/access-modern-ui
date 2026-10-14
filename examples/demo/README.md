# Kit demo

A small Access app built straight from this repository, to see the kit working and to
test changes to it. It has one local table of fictional items and three pages:

| Page | Shows |
|---|---|
| Cards and progress | `Card` (rounded and square), `Progress` at 0, 37 and 100 percent, a table inside a rounded card; the square card's buttons open `KitInputBox`, plain and masked; the rounded card's Primary button shows `KitMsgBox` and Outline shows `KitToast` |
| Tables | `BuildGrid` with every option (header band, zebra rows, status pills, 38px rows, proportional fit through `GridFit`) next to a default table, both sorted by clicking a column caption; filter funnels on the first table's captions; the column chooser (Columns, top right) for the first table; `TabStrip` and chips with counts; search and a category filter to the right of the chips |
| List and detail | The list/detail pattern (`SelectFirstRow`, `ShowDetail`, `ClearDetail`), click-to-sort captions, `PrintGrid` and `ExportGrid` in the list's sort; the Due field is a `DateBox` whose date picker marks federal holidays and every item's due date (`MarkDate`); a compact `Toggle` switch (Flagged for review) on the detail panel's Save row; a `SecretBox` (the item's access code); Save confirms with a toast |

Appearance in the sidebar switches palettes. The demo opens on the kit's splash screen
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

## What to check after a kit change

Click through every page in one light palette (Navy) and one dark palette (Midnight),
watching for Visual Basic error dialogs as well as the error log:

- Splash: opening `KitDemo.accdb` shows no Access logo, then the splash centred in the
  palette's sidebar color with the icon, "Kit Demo", the subtitle, "Build <stamp>" on the
  right and a status line; it closes once the main window is up. The splash takes the palette
  saved last, so to see it in another palette pick one in Appearance, then close and reopen
  the file.
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
  clearing values hides their rows, the funnel turns the primary color, and the chip
  and tab counts and the empty state follow. Select all reads Clear all when every
  value is selected; OK is disabled with none; Cancel changes nothing. The filters
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
  code turned off, the demo shows the page on its own. After Enable Content, Access
  remembers the file as trusted, so the page doesn't open over the main window.
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
- Dialogs: Appearance in the sidebar and the buttons on Cards and progress dim the whole
  Access window behind their dialog; the dim goes when the dialog closes, and the dialog
  itself is not dimmed.

## Files

```
build.ps1                 builds KitDemo.accdb (stages kit\forms and src\forms together)
src/modules/modDemo.bas   the names the kit calls, and the pages' shared helpers
src/modules/modDemoBuild.bas  BuildAll, the table and seed data, Build_<form>
src/modules/modPalette.bas    generated: python ..\..\tools\make_themes.py src
src/themes/light/modTheme.bas generated, as above
src/forms/*.vba           code-behind for frmMain and the pages
```
