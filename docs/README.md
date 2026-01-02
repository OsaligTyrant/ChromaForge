# ChromaForge
ChromaForge is a desktop app that batch-edits PNGs, mirrors folder structure, and cleans filename prefixes using a simple GUI.

## Launch
- Run `ChromaForge.exe` from the main folder.

## What It Does
- Walks folders under the input (Old) directory and writes outputs under the output (New) directory.
- Converts exact hex colors to transparency, or replaces colors based on the mode you pick.
- Renames files by removing known prefixes like `Characters_<number>_`, `Inventory_<number>_`, `FX_<number>_`, `Chars_<number>_`, and `MapGFX_<number>_`.
- Logs only files where conversions happen.

## Modes
- Make transparent: turns exact color matches into full transparency.
- Fill with color: fills transparent pixels (and optional shadows) with a solid hex color.
- Replace color: swaps one exact hex color for another.

## Folder Rules
- Process all folders (default), or pick specific folders.
- Per-folder mode: Process, Skip, or Rename only.
- Skip folders/files that already exist in the output (default).

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
