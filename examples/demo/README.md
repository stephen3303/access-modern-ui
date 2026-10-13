# Kit demo

A small Access app built straight from this repository, to see the kit working and to
test changes to it. It has one local table of fictional items and three pages:

| Page | Shows |
|---|---|
| Cards and progress | `Card` (rounded and square), `Progress` at 0, 37 and 100 percent, a table inside a rounded card |
| Tables | `BuildGrid` with every option (header band, zebra rows, status pills, 38px rows, proportional fit through `GridFit`) next to a default table; `TabStrip` and chips with counts; search and a category filter to the right of the chips |
| List and detail | The list/detail pattern (`SelectFirstRow`, `ShowDetail`, `ClearDetail`), `PrintGrid` and `ExportGrid`; the Due field is a `DateBox` whose date picker marks federal holidays and every item's due date (`MarkDate`) |

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
- Tables: the band, zebra rows, pills and 38px rows show; the selected row is highlighted;
  widening the window widens every column of the first table and moves its captions; the
  chip and tab counts follow the search and category.
- List and detail: the first row opens on load and after each filter; a click shows that
  row; editing and saving updates the list; a search with no match shows both empty
  states; Print previews and Export writes exactly the rows shown, with the captions.
- Date picker (Due's calendar button): October 2026 shows Columbus Day (12th) underlined in
  the danger color and item due dates (1st, 4th, 7th...) in the warning color, with
  tooltips naming them; July 2026 marks Friday the 3rd as Independence Day (observed).
  Changing an item's due date and saving moves its mark.

## Files

```
build.ps1                 builds KitDemo.accdb (stages kit\forms and src\forms together)
src/modules/modDemo.bas   the names the kit calls, and the pages' shared helpers
src/modules/modDemoBuild.bas  BuildAll, the table and seed data, Build_<form>
src/modules/modPalette.bas    generated: python ..\..\tools\make_themes.py src
src/themes/light/modTheme.bas generated, as above
src/forms/*.vba           code-behind for frmMain and the pages
```
