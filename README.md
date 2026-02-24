# FastGenerator

FastGenerator is a SolidWorks VBA macro for batch-exporting drawing files (`.SLDDRW`) into multiple output formats.

## What it does today

- Lets you choose a **source folder** containing `.SLDDRW` files
- Lets you choose a **destination folder**
- Displays drawings in a list and supports **select all / deselect all**
- Exports selected drawings to one or more of these formats:
  - `.dxf`
  - `.dwg`
  - `.PDF`
  - `.JPG`
  - `.TIF`
  - `.edrw`

## Tech stack

- Language: **VBA (Visual Basic for Applications)**
- Platform: **SolidWorks macro environment**
- Core API: **SolidWorks VB API** (`OpenDoc6`, `SaveAs`, `CloseDoc`)

## Project structure

- `Script.vb` — macro logic + form event handlers
- `README.md` — project documentation

## Quick usage

1. Open SolidWorks.
2. Load and run the macro.
3. Click the source-folder button and pick a folder with `.SLDDRW` files.
4. Select the drawings you want to process.
5. Select one or more output formats.
6. Pick a destination folder.
7. Run export.

## Current limitations

- No per-file error report or summary log
- Extension filter is uppercase-only (`*.SLDDRW`) when listing files
- Path separators are mixed (`\` and `/`) in export paths
- No automatic subfolder organization per format
- No overwrite policy prompt (depends on current SaveAs behavior)

## Practical improvements (recommended)

### 1) Reliability first

- Add explicit checks for:
  - source folder exists
  - destination folder exists
  - at least one file selected
  - at least one format selected
- Log success/failure per drawing (CSV or TXT report)
- Catch and display detailed `longstatus/longwarnings` information

### 2) Better file handling

- Support both `.slddrw` and `.SLDDRW`
- Use consistent path joining
- Optional output folders by format (`PDF/`, `DXF/`, etc.)
- Optional naming templates (e.g., drawing number + revision)

### 3) UX improvements

- Progress indicator during batch export
- “Export complete” summary: processed, succeeded, failed
- Persist last used folders and selected formats

### 4) Engineering quality

- Split export logic into reusable functions
- Add comments for API call intent and expected error codes
- Include a small sample workflow in docs/screenshots

## Compatibility notes

This project depends on the SolidWorks macro runtime and API behavior. Keep SolidWorks version notes in this README when changes are made.

## Next suggested milestone

**Milestone 1: Stable Batch Export**

- Add validation + robust error handling
- Add export report file
- Normalize extension/path behavior

That would make this immediately useful for production drafting workflows.
