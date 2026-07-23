# Agent Prompt: Fix Workbook STORE13 Trace Over-Read

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

Workbook-backed power trace parsing is in `src/winsif_mon/domain/verification_plots.py`.

Native solver output is in `src/winsif_mon/services/iterative_solver.py`. The native solver currently sweeps fixed-grip default vehicle offsets from `0.0` to `15.0` in `1.5 m` steps, giving 11 points. That matches the VBA displacement loop:

```vb
For spz = spziniz To spzfin Step passo
```

For the default fixed-grip setup:

- `spziniz = 0`
- `spzfin = eg = 15.0`
- `passo = eg / avanzamento = 1.5`

The workbook reader currently appears to over-read `STORE13` trace blocks into summary or adjacent data. Existing tests may assert inflated trace counts such as 47 forward rows and 28 reverse rows; these are likely reader artifacts, not actual trace-loop counts.

Relevant VBA source:

- `VBA_Export_20260720_120646/Modules/Modulo1.bas`
- `calcolo`: writes trace rows in 5-value groups: ascent tension, descent tension, stroke, offset, motive force.
- `potenza`: reads trace rows until the trace block ends and then writes summaries separately.
- `creacalcolo`: computes `nn13 = Int(spzfin - spziniz) / passo + 1`.

## Task

Fix the workbook `STORE13` trace reader so it stops at the actual trace endpoint and does not read summary or unrelated rows as trace points.

## Implementation Requirements

- Inspect `load_power_trace_case()` and `_read_trace()` in `src/winsif_mon/domain/verification_plots.py`.
- Determine trace count from workbook/default run parameters when possible instead of scanning up to a fixed 450-row window.
- For the default workbook, valid trace count should align with the VBA displacement sweep, expected 11 points for normal fixed-grip cases.
- Keep the parser general enough for non-default spacing by deriving count from F05 parameters where available:
  - spacing / `eg`
  - step size / `passo`
  - advancement steps / `avanzamento`
- Preserve result dataclasses and UI API.
- Do not change the native solver to match an over-read workbook parser.

## Acceptance Criteria

- `load_power_trace_case(ResultFamily.NORMAL, first_active_ref).forward` returns only real trace points.
- Workbook reader and native solver agree on default trace count.
- Existing UI pages still render traces.
- Tests are updated to assert the corrected count and representative trace values.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

