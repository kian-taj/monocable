# Agent Prompt: Complete Hydraulic Final Iteration

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

Hydraulic arithmetic lives in `src/winsif_mon/services/hydraulic.py`.

Hydraulic UI dialog lives in `src/winsif_mon/ui/hydraulic_dialog.py`.

Native solver mode handling is in `src/winsif_mon/services/iterative_solver.py`.

Current state:

- `SolverMode.HYDRAULIC_PRECHECK` creates a native 7x3 `corsaz`-style reference matrix.
- `SolverMode.HYDRAULIC_FINAL` currently produces hydraulic-family result objects using nominal tension, but it does not fully implement the VBA final hydraulic tension iteration.
- The UI precheck flow loads native reference values into `HydraulicState`, then accepting the dialog triggers final hydraulic preparation.

Authoritative VBA:

- `VBA_Export_20260720_120646/Modules/Modulo1.bas`
- `XIndex = 1`: hydraulic precheck
- `XIndex = 2`: hydraulic final
- `precalcolo`
- `precalcsub`
- `creacalcolo`

Critical VBA final loop:

```vb
corsatentativo = (sviluppomax + sviluppomin) / 4
spostamento = corsazero - corsatentativo
DELTA = (corsacil - spostamento)
If Abs(DELTA) <= 0.01 Then Exit Do
If tentativo% = 0 Then
    incr = Tinizio * 0.4
Else
    incr = incr / 2
End If
Tinizio = Tinizio - Sgn(DELTA) * incr
tentativo% = tentativo% + 1
```

## Task

Complete native `HYDRAULIC_FINAL` so it uses dialog-derived hydraulic values and performs the VBA tension iteration.

## Implementation Requirements

- Thread hydraulic final inputs from UI state into solver input:
  - `corsacil`: cylinder stroke
  - selected reference option:
    - bare rope reference (`Opt1`)
    - unloaded rope reference
  - `corsazero`: selected reference stroke value
- Add typed fields; do not pass workbook cell addresses into solver code.
- Modify `HydraulicDialog` so the final selected option and cylinder stroke are retained in `HydraulicState`.
- Modify `run_solver_setup()` / `LineInput` or a nearby typed object so hydraulic final mode receives those values.
- Implement the final tension iteration in `iterative_solver.py`.
- Return hydraulic-family native result objects for line results, traces, max/min, and power summaries.
- Keep current workbook fallback.

## Acceptance Criteria

- Hydraulic final output changes when dialog cylinder/reference inputs change.
- Native hydraulic final trace values are compared against workbook `STORE13` hydraulic blocks:
  - forward block around row `600`
  - reverse block around row `700`
- Tests cover:
  - precheck matrix shape and representative values
  - final iteration convergence
  - final hydraulic family result objects
- Existing tests continue to pass.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

