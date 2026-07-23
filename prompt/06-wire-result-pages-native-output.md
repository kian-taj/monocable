# Agent Prompt: Wire Result Pages to Native Solver Output

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

Native solver results are produced through `run_solver_setup()` in `src/winsif_mon/services/solver.py`.

Native iterative results live in `SolverRunResult.native_iterative_result`.

Result pages currently use workbook-backed readers as authoritative:

- `src/winsif_mon/ui/line_tools_pages.py`
- `src/winsif_mon/ui/result_pages.py`
- `src/winsif_mon/domain/report.py`

Workbook readers:

- `load_line_result_case()`
- `load_power_trace_case()`
- `load_max_min_case()`
- `load_power_summary_case()`

Important constraint:

Do not switch UI pages blindly to native output. Native output should be preferred only when the relevant native result family is marked parity-complete. Workbook fallback must remain available.

## Task

Add a result-provider path so UI pages can consume native solver output when safe, otherwise fall back to workbook stores.

## Implementation Requirements

- Add a small result-provider abstraction or explicit optional parameter flow.
- UI pages should prefer native output only when:
  - `last_solver_result.native_iterative_result` exists
  - the requested mode/family/case exists in native output
  - parity is marked complete for that requested family/mode
- Otherwise, keep current workbook readers.
- Source labels must remain visible:
  - `"Native iterative solver"`
  - `"Workbook STORE06 defaults"`
  - `"Workbook STORE05 defaults"`
  - `"Workbook STORE13 defaults"`
- Avoid changing domain dataclass shapes unless necessary.
- Preserve existing workbook fallback tests.
- Add a UI smoke test that creates `MainWindow` offscreen, runs normal calculation preparation, and verifies pages can refresh without crashing.

## Acceptance Criteria

- Result pages still work with workbook fallback.
- Native result path is used only when explicitly parity-complete.
- No result page crashes when `last_solver_result` is `None`.
- Existing tests continue to pass.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

