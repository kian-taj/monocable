# Agent Prompt: Track Parity Completion by Mode and Result Family

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

`SolverRunResult` in `src/winsif_mon/services/solver.py` has a boolean `parity_complete`, currently false.

`IterativeSolverResult` in `src/winsif_mon/services/iterative_solver.py` also has `parity_complete`, currently false.

The native solver supports result families:

- `ResultFamily.NORMAL`
- `ResultFamily.PLUS_TEN`
- `ResultFamily.MINUS_TEN`
- `ResultFamily.HYDRAULIC`

Modes:

- `SolverMode.NORMAL`
- `SolverMode.VARIABLE_TENSION`
- `SolverMode.ANCHORED`
- `SolverMode.HYDRAULIC_PRECHECK`
- `SolverMode.HYDRAULIC_FINAL`

Current status:

- Some normal-mode tests pass with tolerance.
- Variable tension, anchored, and hydraulic have native paths, but full parity is not proven.
- Result pages should not use native output as authoritative until parity is proven for the requested mode/family.

## Task

Replace the single boolean parity status with family/mode-level parity tracking.

## Implementation Requirements

- Add explicit parity metadata to native results, for example:
  - `parity_complete_families: set[ResultFamily]`
  - or `parity_status: dict[ResultFamily, bool]`
- Keep backward-compatible `parity_complete` property if existing code/tests use it.
- `parity_complete` should return true only if every relevant family in the result is complete.
- Initially mark no family complete unless current tests truly prove it.
- Add helper methods such as:
  - `is_family_parity_complete(family: ResultFamily) -> bool`
  - optional mode-level helper if useful for UI.
- Update UI/native result-provider logic to check family-level parity.
- Add tests verifying:
  - default native result has no authoritative families unless explicitly marked.
  - manually constructed parity-complete results are recognized.
  - UI fallback remains workbook-backed when family parity is false.

## Acceptance Criteria

- No mode claims global parity prematurely.
- Result pages can make family-specific native/fallback decisions.
- Existing tests continue to pass.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

