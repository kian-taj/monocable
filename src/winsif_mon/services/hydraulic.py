from __future__ import annotations

from dataclasses import dataclass

from winsif_mon.models import HydraulicState


@dataclass(frozen=True, slots=True)
class HydraulicResult:
    thermal_stroke: float
    total_nominal: float
    total_plus_10: float
    total_minus_10: float
    max_stroke: float
    cylinder_stroke: float


def calculate_hydraulic(state: HydraulicState) -> HydraulicResult:
    """Replicate the arithmetic performed by `idraulic_input.modifica_parametri`.

    The source VBA stores values in a 1-based `corsaz(14, 3)` array. This Python
    representation keeps only the dialog-visible 7x3 reference matrix plus the
    three editable hydraulic additions.
    """
    refs = state.reference_values
    thermal = state.thermal_stroke
    e1 = state.extra_stroke_1
    e2 = state.extra_stroke_2

    if state.zeta != "F":
        nominal = refs[4][0] + refs[5][0] + refs[6][0] + thermal + e1 + e2
        plus = refs[4][1] + refs[5][1] + refs[6][1] + thermal + e1 + e2
        minus = refs[4][2] + refs[5][2] + refs[6][2] + thermal + e1 + e2
    else:
        nominal = refs[5][0] + refs[6][0] + thermal + e1 + e2
        plus = refs[5][1] + refs[6][1] + thermal + e1 + e2
        minus = refs[5][2] + refs[6][2] + thermal + e1 + e2

    max_stroke = max(nominal, plus, minus)
    cylinder = int(nominal * 10.0) / 10.0 + 0.1
    return HydraulicResult(
        thermal_stroke=thermal,
        total_nominal=nominal,
        total_plus_10=plus,
        total_minus_10=minus,
        max_stroke=max_stroke,
        cylinder_stroke=cylinder,
    )


def format_vba_number(value: float) -> str:
    """Approximate the `Format(..., "###0.00")` display used by the VBA UI."""
    return f"{value:.2f}"
