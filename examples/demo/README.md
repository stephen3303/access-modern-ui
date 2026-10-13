# Kit demo

A small Access app built straight from this repository, to see the kit working and to
test changes to it. It has one local table of fictional items and three pages:

| Page | Shows |
|---|---|
| Cards and progress | `Card` (rounded and square), `Progress` at 0, 37 and 100 percent, a table inside a rounded card; the square card's buttons open `KitInputBox`, plain and masked |
| Tables | `BuildGrid` with every option (header band, zebra rows, status pills, 38px rows, proportional fit through `GridFit`) next to a default table, both sorted by clicking a column caption; `TabStrip` and chips with counts; search and a category filter to the right of the chips |
| List and detail | The list/detail pattern (`SelectFirstRow`, `ShowDetail`, `ClearDetail`), click-to-sort captions, `PrintGrid` and `ExportGrid` in the list's sort; the Due field is a `DateBox` whose date picker marks federal holidays and every item's due date (`MarkDate`); a compact `Toggle` switch (Flagged for review) on the detail panel's Save row |

Appearance in the sidebar switches palettes.

## Build and open

```powershell
cd examples\demo
.\build.ps1            # new KitDemo.accdb with seed data
.\build.ps1 -UiOnly    # rebuild code and forms, keep the data
```

Then open `examples\demo\KitDemo.accdb`. The build stops and prints any Visual Basic or
Access dialog it meets (a compile error, an error message) instead of hanging.

## What to check after a kit change

Click through every page in one light palette (Navy) and one dark palette (Midnight),
watching for Visual Basic error dialogs as well as the error log:

- Cards: corners are round in every palette; hovering the buttons and cards doesn't flicker.
- Progress: 0 shows only the track, 37 is a third full, 100 is full.
- Input box (square card): Ask opens with "Sailor" selected, OK reports the text (an
  empty box reports ""), Cancel and Esc report Cancelled; Ask masked shows * for each
  character and reports only the length.
- Tables: the band, zebra rows, pills and 38px rows show; the selected row is highlighted;
  widening the window widens every column of the first table and moves its captions; the
  chip and tab counts follow the search and category.
- Focus ring: clicking or tabbing into the search box, the category filter or a detail
  field turns its border the palette's primary color, and leaving it restores the border,
  with no flicker and the text not moving.
- Dialogs: Appearance and the card buttons' messages open centred over Access, on its
  monitor, with Access restored, maximized, or moved partly off the edge of a screen.
- List and detail: the first row opens on load and after each filter; a click shows that
  row; editing and saving updates the list; a search with no match shows both empty
  states; Print previews and Export writes exactly the rows shown, with the captions.
- Date picker (Due's calendar button): October 2026 shows Columbus Day (12th) underlined in
  the danger color and item due dates (1st, 4th, 7th...) in the warning color, with
  tooltips naming them; July 2026 marks Friday the 3rd as Independence Day (observed).
  Changing an item's due date and saving moves its mark.
- Toggle (Flagged for review, left of Save in the detail panel): clicking the track or its
  label, or Tab to it and Space, slides the knob and saves the record at once; the track is
  the palette's primary color when on and its muted text color when off; a ring shows while
  it has the focus, inside the panel. The table gained a column, so build without
  `-UiOnly` once.

## Files

```
build.ps1                 builds KitDemo.accdb (stages kit\forms and src\forms together)
src/modules/modDemo.bas   the names the kit calls, and the pages' shared helpers
src/modules/modDemoBuild.bas  BuildAll, the table and seed data, Build_<form>
src/modules/modPalette.bas    generated: python ..\..\tools\make_themes.py src
src/themes/light/modTheme.bas generated, as above
src/forms/*.vba           code-behind for frmMain and the pages
```
