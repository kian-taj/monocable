# Agent Prompt: Improve Anchored Mode Parity

Continue the WinSIF MON PySide6 migration in `/home/kian/Desktop/WinSif_MON`.

## Current Context

Native anchored mode now runs in `src/winsif_mon/services/iterative_solver.py`.

Current behavior:

- `SolverMode.ANCHORED` no longer raises `NotImplementedError`.
- It computes a bare stopped reference length and performs a rope-at-ground length iteration.
- It produces distinct native tensions from normal mode.
- Existing tests assert a representative native anchored output, but they do not prove workbook parity.

Authoritative VBA:

- `VBA_Export_20260720_120646/Modules/Modulo1.bas`
- `XIndex = 4`
- Reference bare-rope calculation:
  - `kk% = 5`
  - `xx% = 1`
  - `precalcsub`
- Reference rope-at-ground length:
  - `LRF_FUNE_A_TERRA`
- Anchored iteration:
  - compute `LGH_FUNE_A_TERRA`
  - `scostamento = LRF_FUNE_A_TERRA - LGH_FUNE_A_TERRA`
  - stop at `Abs(scostamento) <= 0.001`
  - adjust `Tinizio`

Important VBA details:

```vb
DeltaT = Sheets("F05").Cells(22, 16)
Sviluppo = sum(Ci + s(k,w))
Alleltot = sum(svcamp * 10 * Tmean / modulo / area)
Alltetot = sum(svcamp * KT * DeltaT)
LGH_FUNE_A_TERRA = Sviluppo - Alleltot - Alltetot
```

## Task

Refine anchored mode to match workbook anchored behavior and add parity tests.

## Implementation Requirements

- Verify the Python source of `DeltaT` matches `Sheets("F05").Cells(22,16)`, not a nearby display/parameter cell.
- Confirm the reference run uses the same tension multiplier sequence as VBA `precalcsub`.
- Compare native anchored result objects against workbook store blocks for anchored/normal family rows.
- Tighten the secant-style tension update to match VBA:
  - first update uses `Tinizio / 10`
  - later updates use `DELTA = DELTA * Z / (zz - Z)`
- Ensure trial iterations do not pollute final extrema before convergence.
- Keep normal and variable-tension behavior unchanged.

## Acceptance Criteria

- Anchored native output matches workbook anchored values for:
  - first trace point
  - first and last span result
  - first support result
  - max/min first rows
- Tests prove anchored is distinct from normal and parity-aligned with workbook golden fixtures.
- Existing tests continue to pass.

Run:

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q
```

