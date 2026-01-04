# ChromaForge Release Notes

## Beta 2.1 - 2026-01-04

### Added
- Versioned build: the app now advertises Beta 2.1 in the window title and a footer label at the bottom of the UI.
- Splash screen version line: the splash now shows the Beta 2.1 label under the "Created by Koma" credit.
- Layout Editor JSON quick-load tools:
  - Recent JSON list (scans the input root for `sprite_sheets` or `tilemaps` folders depending on the selected type).
  - One-click Load button for the selected recent JSON entry.
  - Drag-and-drop loading of `.json` layout files directly onto the Layout Editor canvas.
- Layout Editor copy workflow: "Copy Selected" duplicates the selected items and drops them into the next empty grid cell.
- Layout Editor naming clarity: the name field now reads "Spritesheet name" or "Tilemap name" based on the selected type.
- Spritesheet auto-export names now use a capitalized suffix (`_Spritesheet`) for consistency.

### Changed
- Layout Editor output naming behavior:
  - Spritesheet exports use the exact name entered in the name field (e.g., `DEFBOD.png` and `DEFBOD.json`).
  - Tilemap exports use the exact name entered (e.g., `ROOMA.png`, `ROOMA.json`, and `ROOMA.csv` when metadata is enabled).
- Layout Editor load/save flow now goes through a shared loader so recent lists stay in sync after loads and saves.
- Input folder Browse now refreshes the Layout Editor folder list immediately.
- Spritesheet grouping prefix logic now handles numeric-first filenames by using the next meaningful token when appropriate (keeps numeric prefixes like `1012-F4-E`, but prefers tokens like `DEFBOD` when the first token is numeric).

### Layout Editor Highlights (Beta 2.1)
- Manual placement canvas with:
  - Grid or free-form layout mode.
  - Snap to grid toggle.
  - Full grid line rendering across every row/column.
  - Optional alignment guides overlay.
  - Mouse-wheel zoom for fine placement.
- Positioning tools:
  - X/Y fields with Apply.
  - Align Horizontal / Align Vertical.
  - Center in Cell (centers within the current grid cell).
  - Remove Selected (deletes items from canvas only).
  - Copy Selected (duplicates items into next available empty grid cell).
- Export outputs:
  - Spritesheets saved under `sprite_sheets` with matching JSON.
  - Tilemaps saved under `tilemaps` with matching JSON and optional CSV metadata.

### Build Notes
- Packaged with PyInstaller (one-file, no console) and includes `tkinterdnd2` resources so drag-and-drop works in the EXE build.

