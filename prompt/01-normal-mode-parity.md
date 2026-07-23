# Agent Prompt: Strengthen Normal-Mode Native Solver Parity

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

The native iterative solver is in `src/winsif_mon/services/iterative_solver.py`. The solver boundary is `src/winsif_mon/services/solver.py`. Workbook-backed golden readers live in:

- `src/winsif_mon/domain/line_results.py` for `STORE06`
- `src/winsif_mon/domain/max_min_results.py` for `STORE05`
- `src/winsif_mon/domain/verification_plots.py` for `STORE13` traces
- `src/winsif_mon/domain/power_summary.py` for `STORE13` summaries

The authoritative VBA source is `VBA_Export_20260720_120646/Modules/Modulo1.bas`, especially `esegui_verifica`, `calcolo`, `campata`, `memorizza`, `risultati`, `potenza`, and `analisi`.

Current known tests pass:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

Last known result was `37 passed`.

The current native normal solver has tolerance-based tests for the first default `STORE06` and `STORE13` values. It is not fully parity-proven across all active cases, branches, spans, supports, directions, or summaries.

Default active cases:

- Hypothesis 1 / condition 2, store index `2`
- Hypothesis 2 / condition 2, store index `8`

Useful golden values already used:

- First `STORE13` normal trace, case 2:
  - offset `0.0`
  - ascent tension `7471.770942704483`
  - descent tension `11369.716244164345`
  - motive force `3897.945301459862`
- First `STORE06` normal span, case 2:
  - span `Valle` to `AV`
  - valley tension `7530.783599668728`
  - mountain tension `7530.783599668727`
  - sag `0.025423470370434634`
- First `STORE05` normal max span:
  - span `Valle` to `AV`
  - tension `9837.15372805009`
  - sag `0.025423470370434634`
- First `STORE05` support row:
  - support `AV`
  - support tension `9838.035785608645`
  - roller count `2.0`

## Task

Expand normal-mode parity tests and fix native solver discrepancies found by those tests.

## Implementation Requirements

- Add tests for both default active normal cases, store indexes `2` and `8`.
- Compare native output against workbook golden readers for:
  - `STORE13`: first, second, middle valid, and last valid trace point for forward and reverse.
  - `STORE06`: first, middle, and last span rows for ascent and descent.
  - `STORE06`: first, middle, and last support rows for ascent and descent.
  - `STORE05`: first, middle, and last max/min rows.
  - `STORE13` power summaries for both cases and both directions.
- Use workbook readers as golden fixtures; do not hardcode large tables.
- If a mismatch is found, inspect the relevant VBA formula and fix the Python solver rather than widening tolerance by default.
- Keep workbook-store fallback behavior unchanged.
- Do not mark `parity_complete=True` unless broad parity is actually proven.

## Acceptance Criteria

- Normal-mode native tests cover both default active cases and both directions.
- Existing tests continue to pass.
- New tests document any tolerated delta explicitly with `pytest.approx(..., abs=...)`.
- Tolerances should be tightened as much as current numeric behavior allows.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

