# MONOFUNI VBA Flow Notes

## What This Workbook Is

`MONOFUNI.xls` is an Excel/VBA engineering calculator for a monocable ropeway/chairlift line. Excel is not just a spreadsheet here: it is the UI framework, data store, intermediate calculation memory, chart/report renderer, and print/export surface.

The PySide6 migration should therefore not translate `Cells(row, col)` literally. Each important cell range needs a named model first, then the UI binds to those names.

## Startup

Execution starts in `ThisWorkbook.Workbook_Open`.

1. Creates/uses `C:\TMP_SIF`.
2. Reads an activation key file and computes a machine code from the Windows volume serial.
3. Writes `MONOFUNI_TEST_PSWKEY = activate` when activation succeeds.
4. Activates `HOME`.
5. Reads `c:\tmp_sif\prg_path` into global `mia_path`.
6. Sets `test_ricalcolo = True`.

For the PySide6 app, the activation and `C:\TMP_SIF` path behavior should be compatibility-only, not part of the calculation core.

## Main Pages

- `HOME`: startup/language entry page. It stores Italian/English flags.
- `MENU`: navigation hub.
- `F01`: general plant data, rope data, vehicle data, friction/wind/roller data. `Modulo1.assegna_generali` copies this sheet into global calculation variables.
- `F02`: terrain profile for ascent and descent branches. `Modulo1.leggi_terreno` reads the profile from row 15 onward.
- `F03` / `Linea`: absolute support/tower data. It can derive relative span data in `F04`.
- `F04` / `Linea1`: relative span data. It can derive absolute support/tower data in `F03`.
- `F05`: main calculation control page. Buttons call `prepara_calcolo` and `esegui_verifica`.
- `F06`/`F07`: load cases and friction/attribution data.
- `F08`/`F09`/`F10`/`F11`/`F12`/`F20`: outputs, charts, max/min tables, power tables, general report, and tension-laying report.
- `F13`: labels/translations used to update captions and reports.
- `ROPE_1`, `ROPE_2`, `ROPE_3`: rope catalog tables used by `F01` setup buttons.
- `STORE05`, `STORE06`, `STORE13`: intermediate calculation/result stores.

## Calculation Entry Points

The main buttons are in the `F05` sheet module, exported as `Foglio11.cls`:

- `calcola_normale_Click`: normal calculation, then variable tension, then max/min, powers, and report.
- `calcola_ancorata_Click`: anchored calculation.
- `calcola_idraul_Click`: hydraulic pre-calculation, then opens `idraulic_input`.
- `calcola_varten_Click`: variable-tension calculation.

All of these eventually call `Modulo1.esegui_verifica(XIndex)`.

`XIndex` changes the calculation mode:

- `0`: normal verification
- `1`: hydraulic pre-check and dialog preparation
- `2`: final hydraulic verification after dialog input
- `4`: anchored/fixed reference calculation
- `5`: variable tension calculation

## What `esegui_verifica` Does

At a high level:

1. Allocates large global arrays for two line branches, supports, spans, vehicles, forces, tensions, deflections, and max/min values.
2. Reads the selected hypothesis matrix from `F05` cells `D10:I24`; `OOO` means one enabled state, `XXX` means another enabled state.
3. Computes vehicle spacing/step size from plant type, rope length, speed, and capacity.
4. For each selected hypothesis and each running mode, runs the line calculation twice: forward gear and reverse gear.
5. For each vehicle displacement step, calls `posizione_veicoli`, calculates span/support effects, iterates rope tensions until the station/tensioner condition is satisfied, then stores max/min values.
6. Writes intermediate and final results to `STORE06` and `STORE13`.
7. Calls `analisi` to summarize raw stored results.
8. Copies the hypothesis matrix into output/report sheets.

The calculation is self-contained in the workbook/VBA code. The hard part is not missing formulas; it is untangling workbook-global mutable state into named Python structures.

## Store Sheets

- `STORE06`: detailed line tables used by `F08`, `F10`, `F12`, and report rendering.
- `STORE13`: per-hypothesis tension/power summary data used by `F11` and later report sections.
- `STORE05`: additional max/min support data used by output table rendering.

In Python these should become result dataclasses/lists, not hidden worksheet tables.

## Current PySide6 Migration Status

Implemented:

- `F01` is now a workbook-backed page with named fields and Excel defaults.
- `F01` uses unit-aware PySide controls and stores values by names such as `working_speed`, not by UI cell addresses.
- `F01` rope setup reads `ROPE_1`, `ROPE_2`, and `ROPE_3` catalogs.
- `F02` is now a workbook-backed terrain profile page using named `TerrainPoint` records and an embedded PySide chart.
- `F03` / `Linea` is now a structured Support Geometry page with ascent supports, descent supports, and support/foundation drawing dimensions as separate named tables.
- `F04` / `Linea1` is now a structured Span Geometry page with first-distance/first-height fields and ascent/descent span tables.
- `F05` is now a structured Line Verification page with plant run parameters and the 15 x 6 hypothesis matrix.
- `F06` is now a structured Custom Load Cases page with span rows and 10 ascent/descent imposed-load hypothesis pairs.
- F05 buttons now call the Python solver preparation service, which ports input assembly, line-length/run-parameter derivation, F07 friction loading, F06 load preparation, and `Modulo6.posizione_veicoli`.
- Visible long sheets such as `F07`, `F08`, `F09`, `F10`, `F11`, `F12`, `F13`, and `F20` still load their workbook defaults into friendly PySide pages while their typed page models are pending.
- `idraulic_input` has an initial dialog scaffold and arithmetic for the visible hydraulic totals.
- Inventory tooling confirms `34` VBA components, `182` procedures, and `22` workbook sheets.

Not implemented yet:

- Typed page for `F07` beyond the current workbook-backed table view.
- Full Python port of the iterative span/tension loop inside `esegui_verifica`, including `campata`, `memorizza`, `analisi`, `stampa_maxmin`, `stampa_potenze`, and `rel_gen`.
- Report rendering from `STORE05/STORE06/STORE13`.

## Important Missing Or External Pieces

The workbook itself is readable and contains the main sheets, defaults, rope catalogs, and store sheets. No critical calculation sheet appears missing.

External/non-portable behavior found in the VBA:

- Windows volume serial activation through `DriveInfo.bas`.
- `C:\TMP_SIF` temporary files.
- `MF_CREA_WINDXF.EXE` for DXF generation, not present in this workspace.
- `view_line.exe`, not present in this workspace.
- `MF_AUTOCAD.dwg`, present.
- HTML help folders, partly present under `SOSTEGNI_MONOFUNE`.

The calculation can be ported without those `.exe` files, but DXF/viewer functionality cannot be reproduced exactly unless those executables or their expected output format are available.
