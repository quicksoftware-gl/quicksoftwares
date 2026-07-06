# Windows_Pic

Product images shown in the Software Catalog for the **Windows** sheet
of `Softwares.xlsx`.

## Naming convention

The filename must match the **Excel row number** of the software.

- Row 1 is the header, so no image is needed for it.
- Row 2 (the first software) → `2.jpg`
- Row 3 → `3.jpg`
- …and so on.

Supported extensions: `.jpg`, `.png`, `.webp` (tried in that order).
Missing images fall back to a "No image" placeholder in the UI.

## After adding/removing images

Run `./update-pics.py` from the repo root so the frontend knows the file is available. Otherwise the tile falls back to the "No image" placeholder.
