# ChromaForge
ChromaForge is a desktop app that batch-edits PNGs, groups outputs by filename prefixes, and builds sprite sheets and tilemaps using a simple GUI.

Version: Beta 3.1.1

## Launch
- Run `ChromaForge.exe` from the main folder.

## What It Does
- Walks folders under the input (Old) directory and writes outputs under the output (New) directory.
- Converts exact hex colors to transparency, or replaces colors based on the mode you pick.
- Routes files into prefix-named subfolders (lowercase) based on the filename prefix, not the original subfolder. Non-matching files go into `needs_sorting`.
- Renames files by removing known prefixes like `Characters_<number>_`, `Inventory_<number>_`, `FX_<number>_`, `Chars_<number>_`, and `MapGFX_<number>_`.
- Logs only files where conversions happen.
- Preview routing shows counts per prefix before running.
- Scan for prefixes lists detected prefixes so you can choose which ones to use.
- Keep prefixes preserves original filenames when checked.
- Pre-check scan warns about naming conflicts and offers overwrite or add-copy options.

## Tabs
- Color Mode: color removal/fill/replace with prefix cleanup and output grouping.
- Spritesheet Mode: groups frames by prefix within each folder and writes sheets to a `sprite_sheets` subfolder.
- Tilemap Mode: builds tilemaps per folder in a `tilemaps` subfolder and can export CSV metadata.
- Layout Editor (Assemble): manual drag-and-drop layout with grid/free-form, optional guides, JSON layouts saved alongside output, plus recent JSON and drag-and-drop JSON loading.
- Layout Editor (Split): load a spritesheet, configure a grid, select cells, and export frames back out as PNGs.

## Release Notes
See `CHANGELOG.md` in the app root for detailed release notes.

## Modes
- Make transparent: turns exact color matches into full transparency.
- Fill with color: fills transparent pixels (and optional shadows) with a solid hex color.
- Replace color: swaps one exact hex color for another.

## Folder Rules
- Process all folders (default), or pick specific folders.
- Per-folder mode: Process, Skip, or Rename only.
- Skip prefix folders/files that already exist in the output (default).

## Presets
- Save and load presets for different workflows.
- Recent presets list for quick reuse.

## Themes and Settings
- Light and Dark themes.
- Theme and window size auto-save on change.
- Settings are stored in the `settings` folder next to the EXE.

## Tips
- Use exact 6-digit hex values (like `#00FF00`).
- Folder list updates with the Refresh button if new folders are added to the input directory.
