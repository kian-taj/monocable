# Agent Prompt: Port Faithful STORE05 `analisi` Aggregation

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

Native max/min packaging currently exists in `src/winsif_mon/services/iterative_solver.py`, but it is simplified. It now combines active cases, but it is not a faithful port of VBA `analisi`.

Workbook-backed max/min parsing is in `src/winsif_mon/domain/max_min_results.py`.

Authoritative VBA source:

- `VBA_Export_20260720_120646/Modules/Modulo1.bas`
- `Public Sub analisi()`
- Labels/procedures inside or near `analisi`:
  - `tabella_maxmin`
  - `spoolPMMFMM`
  - `leggiPMMFMM`
  - `assegna_rulli`
  - `stampa_maxmin`

Existing useful golden defaults:

- First `STORE05` normal max span:
  - `Valle` to `AV`
  - tension `9837.15372805009`
  - sag `0.025423470370434634`
- First `STORE05` support row:
  - support `AV`
  - support tension `9838.035785608645`
  - roller count `2.0`

## Task

Replace approximate native STORE05/max-min aggregation with a faithful Python implementation of the VBA `analisi` result aggregation.

## Implementation Requirements

- Preserve the existing domain dataclasses in `src/winsif_mon/domain/max_min_results.py`.
- Implement aggregation in native solver/result layer without writing workbook cells.
- Match VBA row semantics for:
  - span max/min tension
  - sag max/min
  - valley/mountain angle max/min
  - support tension max/min
  - deviation max/min
  - pressure max/min
  - friction max/min
  - roller count
  - unit deviation / unit pressure
  - validation/test labels such as `-NV-`
- Port the logic that compares ascent/descent when `ns == nd`.
- Port support validation thresholds using already-loaded general inputs:
  - max support roller load
  - max compression roller load
  - support/compression roller admissible deviation if available
  - double-acting roller thresholds if available
- If a field is not yet loaded in `GeneralInput`, add it explicitly with a named field instead of using workbook cell addresses in solver code.
- Keep workbook reader as golden fixture and fallback.

## Acceptance Criteria

- Native `MaxMinCase` matches workbook `STORE05` for representative rows:
  - first, middle, last span rows
  - first, middle, last support rows
  - ascent and descent branches
  - normal, variable-tension, anchored, and hydraulic result families where available
- Existing tests continue to pass.
- New tests compare native values against `load_max_min_case()`.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

