from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from winsif_mon.domain.f01_general import load_f01_defaults
from winsif_mon.domain.geometry import GeometryState, SpanRow, load_geometry_state
from winsif_mon.domain.verification import load_verification_state


@dataclass(slots=True)
class LayingSpanRow:
    span_start: str
    span_end: str
    horizontal_distance_m: float
    height_difference_m: float
    laying_tension_da_n: float


@dataclass(slots=True)
class LayingTensionState:
    ambient_temperature_c: float
    reference_tension_da_n: float
    ascent_rows: list[LayingSpanRow]
    descent_rows: list[LayingSpanRow]
    temperature_curve: list[tuple[float, float, float]]


def load_laying_tension_state(
    geometry: GeometryState | None = None,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> LayingTensionState:
    geometry = geometry or load_geometry_state(workbook_path)
    general = load_f01_defaults(workbook_path)
    verification = load_verification_state(workbook_path)
    reference = _number(general.value("rated_tightener_tension"))
    if reference <= 0:
        reference = 25000.0
    ambient = verification.parameters.local_temperature_c
    ascent = _rows(geometry.ascent_spans, ambient, reference)
    descent_source = geometry.descent_spans if any(span.code for span in geometry.descent_spans) else geometry.ascent_spans
    descent = _rows(descent_source, ambient, reference)
    curve = [
        (temperature, _temperature_tension(reference, temperature, 0), _temperature_tension(reference, temperature, 1))
        for temperature in range(-20, 31)
    ]
    return LayingTensionState(ambient, reference, ascent, descent, curve)


def _rows(spans: list[SpanRow], temperature: float, reference: float) -> list[LayingSpanRow]:
    rows: list[LayingSpanRow] = []
    for span in spans:
        if not span.code:
            continue
        start, _sep, end = span.code.partition("-")
        rows.append(
            LayingSpanRow(
                span_start=start,
                span_end=end,
                horizontal_distance_m=_number(span.horizontal_distance_m),
                height_difference_m=_number(span.height_difference_m),
                laying_tension_da_n=_temperature_tension(reference, temperature, len(rows)),
            )
        )
    return rows


def _temperature_tension(reference: float, temperature: float, index: int) -> float:
    return max(reference * (1.0 - 0.00012 * temperature) - index * 2.5, 0.0)


def _number(value) -> float:
    if value == "":
        return 0.0
    return float(value)
