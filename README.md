# WinSIF MON PySide6 Migration

This repository contains a standalone Python/PySide6 migration of the visible `MONOFUNI.xls` workbook UI and the first Python service boundaries for the VBA calculation code.

## What Exists

- `winsif-mon`: launches a PySide6 desktop shell with friendly workflow navigation.
- Workbook-backed defaults for the main input, result, report, and plotting sheets.
- Modern light UI theme with named fields and unit-aware controls instead of raw Excel cell references.
- Embedded terrain profile plotting for the `F02` terrain data.
- Structured Support Geometry and Span Geometry pages for the workbook `F03`/`F04` workflow.
- Structured Line Verification and Custom Load Cases pages for the workbook `F05`/`F06` workflow.
- Solver preparation service for F05: line input assembly, run parameters, friction defaults, vehicle positions, and load preparation.
- `winsif-inventory`: scans the exported VBA project and, when `xlrd` is installed, the legacy `.xls` workbook metadata.
- A hydraulic-tension dialog scaffold based on `Forms/idraulic_input.frm`.
- Python models/services that separate future calculation logic from worksheet storage.

## Run

```bash
python3 -m pip install -r requirements.txt
python3 -m winsif_mon.main
```

If you are using the local virtual environment from this project:

```bash
.venv/bin/python -m winsif_mon.main
```

Inventory:

```bash
python3 -m winsif_mon.inventory --json
```

## Migration Status

This is not yet a full calculation replacement for the workbook. The visible sheets, defaults, data tables, first plot, and F05 solver preparation path are migrated, but the iterative numerical tension loop in `Modulo1.esegui_verifica` still needs to be ported before calculated output parity can be claimed.

See `docs/verification_workflow.md` for the Excel-vs-PySide comparison process.
