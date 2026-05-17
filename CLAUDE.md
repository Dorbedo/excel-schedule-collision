# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel VBA tool for detecting layout and schedule collisions on construction/production sites. Written entirely in VBA for Excel compatibility. The main workbook is `excel-schedule-collision.xlsm`.

## Architecture

All VBA source files live in `src/` as `.bas` modules. They must be manually imported into the Excel workbook via the VBA editor (`Alt+F11` → File → Import File). After editing a module in Excel, export it back to `src/` to keep the repo in sync.

### Module responsibilities

| Module | Role |
|---|---|
| `Main.bas` | Single entry point for all user-facing macros (`Run*` subs). All buttons/shortcuts should point here. |
| `TableSetup.bas` | Creates the `Vorgangsliste` sheet with the task table (MS Project–compatible columns + Status dropdown + Farbe column). |
| `CalibrationSetup.bas` | Creates the `Kalibrierung` sheet with `Kalibrierung_X` and `Kalibrierung_Y` tables. |
| `LayoutImage.bas` | Import workflow: file dialog → layout name → embed image on a dedicated sheet → trigger calibration if new. Also exposes `GetLayoutSheet()` and `GetLayoutPicture()` for other modules. |
| `Calibration.bas` | Rectangle-based calibration (`PlaceCalibrationRect` / `ApplyCalibrationRect`) and calibration marker display (`ShowCalibrationMarkers` / `HideCalibrationMarkers`). Reads/writes the `Kalibrierung_X/Y` tables. |
| `ColorHelper.bas` | Parses HEX (`#RRGGBB`, `#AARRGGBB`) and RGB (`R,G,B`, `R,G,B,A`) color strings, applies fill to the `Farbe` column, and reads existing cell fills back as `#AARRGGBB`. Default alpha = `0x80`. |
| `CollisionChecker.bas` | Core collision logic stub — `HasCollision()` and `CheckAll()` (not yet implemented). |
| `FormBuilder.bas` | Programmatically creates `frmLayoutSelector` UserForm via the VBE API. Requires *Trust access to VBA project object model* in Trust Center. Reserved for future use — currently not called. |

### Key data model

- **Vorgangsliste** — main task table (sheet + ListObject). Columns: ID, Layoutname, Vorgangsname, Dauer, Anfang, Ende, Vorgänger, Bereich/Fläche, Verantwortlicher, Status, Farbe.
- **Kalibrierung_X / Kalibrierung_Y** — calibration tables (sheet `Kalibrierung`). Columns: Layout, Bezeichnung, Pixelposition, Abstand (m). Pixelposition is stored in Excel sheet points (not image pixels). Multiple layouts share the same tables, distinguished by the `Layout` column.
- **Layout sheets** — one sheet per imported layout, named after the layout (sanitized). Contains one embedded picture shape named `Bild_<sheetname>`.

### Naming conventions for shapes

| Prefix | Purpose |
|---|---|
| `Bild_<layout>` | Embedded layout image |
| `Kalibrierung_<layout>` | Calibration rectangle (temporary, removed after `ApplyCalibrationRect`) |
| `KalMkr_X_<layout>_<label>` | X-axis calibration marker |
| `KalMkr_Y_<layout>_<label>` | Y-axis calibration marker |

## VBA-specific conventions

- **Umlauts** are never written literally in `.bas` source files — use `Chr()` codes: ä=228, ö=246, ü=252, Ä=196, Ö=214, Ü=220, ß=223. This avoids encoding corruption on import.
- **Long string literals** that span many lines must be built with separate `c = c & "..."` assignments — VBA enforces a maximum of 24 line continuations (`_`) per statement.
- **Late binding** (`As Object`) on VBE Designer controls only exposes the base `MSForms.Control` interface. Type-specific properties (`Style`, `Default`, `Cancel`) must be set at runtime (e.g., in `UserForm_Initialize`), not during form construction.
- `Chr()` only accepts values 0–255. Unicode code points (e.g., `→` = 8594) will throw *Invalid procedure call*.

## Workflow

1. `RunSetup` — creates Vorgangsliste + Kalibrierung sheets (run once on a fresh workbook).
2. `RunImportLayout` — import a layout image; prompts for calibration points if the layout is new.
3. `RunPlaceCalibrationRect` / `RunApplyCalibrationRect` — refine calibration via a draggable rectangle.
4. `RunShowCalibrationMarkers` / `RunHideCalibrationMarkers` — visual verification of calibration.
5. `RunColorProcessing` — applies colors from the `Farbe` column to cell fills.
6. `RunCollisionCheck` — collision detection (not yet implemented).

## TODO

- Interpolation logic in `Calibration.bas` to convert pixel/point positions to real-world metres.
- `CollisionChecker.CheckAll` implementation.
- Layout dropdown selector (`FormBuilder.bas` / `frmLayoutSelector`) — deferred due to VBE late-binding constraints.
