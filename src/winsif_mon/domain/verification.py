from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from winsif_mon.domain.geometry import GeometryState
from winsif_mon.workbook import WorkbookReader


class MatrixState(str, Enum):
    OFF = ""
    NORMAL = "OOO"
    ALTERNATE = "XXX"


@dataclass(slots=True)
class PlantCondition:
    index: int
    label: str


@dataclass(slots=True)
class LoadHypothesis:
    index: int
    ascent_label: str
    descent_label: str

    @property
    def label(self) -> str:
        return f"{self.ascent_label} / {self.descent_label}" if self.descent_label else self.ascent_label


@dataclass(slots=True)
class PlantRunParameters:
    rope_loop_length_m: float
    total_cars: float
    car_spacing_m: float
    running_speed_m_s: float
    carrying_capacity_p_h: float
    step_advancement_m: float
    advancement_steps: float
    tightener_precision_percent: float
    local_temperature_c: float


@dataclass(slots=True)
class VerificationState:
    hypotheses: list[LoadHypothesis]
    conditions: list[PlantCondition]
    matrix: list[list[MatrixState]]
    parameters: PlantRunParameters


@dataclass(slots=True)
class SpanLoadRow:
    ascent_span: str
    descent_span: str
    loads: list[float | str]


@dataclass(slots=True)
class CustomLoadState:
    rows: list[SpanLoadRow]


@dataclass(frozen=True, slots=True)
class CalculationSetupSnapshot:
    verification: VerificationState
    custom_loads: CustomLoadState


def load_verification_state(workbook_path: Path | str = "MONOFUNI.xls") -> VerificationState:
    reader = WorkbookReader(workbook_path)
    return VerificationState(
        hypotheses=_load_hypotheses(reader),
        conditions=_load_conditions(reader),
        matrix=_load_matrix(reader),
        parameters=PlantRunParameters(
            rope_loop_length_m=_number(reader.value("F05", 3, 15)),
            total_cars=_number(reader.value("F05", 4, 15)),
            car_spacing_m=_number(reader.value("F05", 5, 15)),
            running_speed_m_s=_number(reader.value("F05", 6, 15)),
            carrying_capacity_p_h=_number(reader.value("F05", 7, 15)),
            step_advancement_m=_number(reader.value("F05", 10, 15)),
            advancement_steps=_number(reader.value("F05", 12, 15)),
            tightener_precision_percent=_number(reader.value("F05", 17, 17)),
            local_temperature_c=_number(reader.value("F05", 22, 17)),
        ),
    )


def load_custom_load_state(
    workbook_path: Path | str = "MONOFUNI.xls",
    geometry_state: GeometryState | None = None,
) -> CustomLoadState:
    reader = WorkbookReader(workbook_path)
    rows: list[SpanLoadRow] = []
    geometry_spans = _span_names_from_geometry(geometry_state) if geometry_state else []
    for index, excel_row in enumerate(range(12, 93)):
        ascent_span = str(reader.value("F06", excel_row, 1)).strip()
        descent_span = str(reader.value("F06", excel_row, 2)).strip()
        if index < len(geometry_spans):
            ascent_span, descent_span = geometry_spans[index]
        loads = [reader.value("F06", excel_row, col) for col in range(3, 23)]
        if not ascent_span and not descent_span and all(value == "" for value in loads):
            continue
        rows.append(SpanLoadRow(ascent_span=ascent_span, descent_span=descent_span, loads=loads))
    return CustomLoadState(rows=rows)


def cycle_matrix_state(state: MatrixState) -> MatrixState:
    if state is MatrixState.OFF:
        return MatrixState.NORMAL
    if state is MatrixState.NORMAL:
        return MatrixState.ALTERNATE
    return MatrixState.OFF


def reset_matrix(state: VerificationState) -> None:
    state.matrix = [[MatrixState.OFF for _condition in state.conditions] for _hypothesis in state.hypotheses]


def validate_custom_loads(custom_loads: CustomLoadState) -> list[str]:
    errors: list[str] = []
    for row_index, row in enumerate(custom_loads.rows, start=1):
        for col_index, value in enumerate(row.loads, start=1):
            if value == "":
                continue
            try:
                float(value)
            except (TypeError, ValueError):
                hypothesis = (col_index + 1) // 2
                branch = "ascent" if col_index % 2 else "descent"
                errors.append(f"Row {row_index}, hypothesis {hypothesis} {branch}: {value!r} is not numeric")
    return errors


def build_calculation_setup(
    verification: VerificationState,
    custom_loads: CustomLoadState,
) -> CalculationSetupSnapshot:
    errors = validate_custom_loads(custom_loads)
    if errors:
        raise ValueError("\n".join(errors))
    return CalculationSetupSnapshot(verification=verification, custom_loads=custom_loads)


def _load_hypotheses(reader: WorkbookReader) -> list[LoadHypothesis]:
    hypotheses: list[LoadHypothesis] = []
    for index, row in enumerate(range(10, 25), start=1):
        hypotheses.append(
            LoadHypothesis(
                index=index,
                ascent_label=str(reader.value("F05", row, 1)).strip(),
                descent_label=str(reader.value("F05", row, 2)).strip(),
            )
        )
    return hypotheses


def _load_conditions(reader: WorkbookReader) -> list[PlantCondition]:
    conditions: list[PlantCondition] = []
    for index, col in enumerate(range(4, 10), start=1):
        label = f"{reader.value('F05', 7, col)} {reader.value('F05', 8, col)}".strip()
        conditions.append(PlantCondition(index=index, label=" ".join(label.split())))
    return conditions


def _load_matrix(reader: WorkbookReader) -> list[list[MatrixState]]:
    matrix: list[list[MatrixState]] = []
    for row in range(10, 25):
        states: list[MatrixState] = []
        for col in range(4, 10):
            states.append(_matrix_state(reader.value("F05", row, col)))
        matrix.append(states)
    return matrix


def _matrix_state(value: Any) -> MatrixState:
    if value == MatrixState.NORMAL.value:
        return MatrixState.NORMAL
    if value == MatrixState.ALTERNATE.value:
        return MatrixState.ALTERNATE
    return MatrixState.OFF


def _span_names_from_geometry(geometry_state: GeometryState | None) -> list[tuple[str, str]]:
    if geometry_state is None:
        return []
    rows: list[tuple[str, str]] = []
    for ascent, descent in zip(geometry_state.ascent_spans, geometry_state.descent_spans):
        if ascent.code or descent.code:
            rows.append((ascent.code, descent.code or ascent.code))
    return rows


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
