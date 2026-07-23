from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from winsif_mon.domain.geometry import GeometryState, load_geometry_state
from winsif_mon.domain.verification import MatrixState, VerificationState, load_verification_state
from winsif_mon.workbook import WorkbookReader

VBA_PI = 3.141


class ResultFamily(str, Enum):
    NORMAL = "normal"
    PLUS_TEN = "plus_ten"
    MINUS_TEN = "minus_ten"
    HYDRAULIC = "hydraulic"

    @property
    def label(self) -> str:
        return {
            ResultFamily.NORMAL: "Normal / anchored",
            ResultFamily.PLUS_TEN: "+10% tension",
            ResultFamily.MINUS_TEN: "-10% tension",
            ResultFamily.HYDRAULIC: "Hydraulic",
        }[self]

    @property
    def store_row(self) -> int:
        return {
            ResultFamily.NORMAL: 5,
            ResultFamily.HYDRAULIC: 2005,
            ResultFamily.PLUS_TEN: 4005,
            ResultFamily.MINUS_TEN: 6005,
        }[self]


@dataclass(frozen=True, slots=True)
class ResultCaseRef:
    hypothesis_index: int
    condition_index: int
    state: MatrixState
    label: str

    @property
    def store_index(self) -> int:
        return (self.hypothesis_index - 1) * 6 + self.condition_index


@dataclass(slots=True)
class LineResultRow:
    span_start: str = ""
    span_end: str = ""
    valley_tension_da_n: float | str = ""
    mountain_tension_da_n: float | str = ""
    sag_m: float | str = ""
    valley_angle_deg: float | str = ""
    mountain_angle_deg: float | str = ""
    support_code: str = ""
    support_tension_da_n: float | str = ""
    total_deviation_deg: float | str = ""
    pressure_da_n: float | str = ""
    friction_da_n: float | str = ""
    roller_count: float | str = ""
    unit_deviation_deg: float | str = ""
    unit_pressure_da_n: float | str = ""


@dataclass(slots=True)
class LineResultBranch:
    name: str
    rows: list[LineResultRow]


@dataclass(slots=True)
class LineResultCase:
    family: ResultFamily
    reference: ResultCaseRef
    source: str
    flag_average: int
    ascent: LineResultBranch
    descent: LineResultBranch


def active_result_cases(verification: VerificationState) -> list[ResultCaseRef]:
    refs: list[ResultCaseRef] = []
    for hypothesis_index, row in enumerate(verification.matrix, start=1):
        for condition_index, state in enumerate(row, start=1):
            if state is MatrixState.OFF:
                continue
            hypothesis = verification.hypotheses[hypothesis_index - 1]
            condition = verification.conditions[condition_index - 1]
            refs.append(
                ResultCaseRef(
                    hypothesis_index=hypothesis_index,
                    condition_index=condition_index,
                    state=state,
                    label=f"{hypothesis.index}. {hypothesis.label} - {condition.label}",
                )
            )
    return refs


def load_line_result_case(
    family: ResultFamily = ResultFamily.NORMAL,
    reference: ResultCaseRef | None = None,
    verification: VerificationState | None = None,
    geometry: GeometryState | None = None,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> LineResultCase:
    verification = verification or load_verification_state(workbook_path)
    geometry = geometry or load_geometry_state(workbook_path)
    reference = reference or _first_reference(verification)
    reader = WorkbookReader(workbook_path)
    start_row = family.store_row
    col_ascent = 1 + reference.store_index
    col_descent = 92 + reference.store_index
    flag_average = int(_number(reader.value("STORE06", start_row, col_ascent)))
    ascent_count = int(_number(reader.value("STORE06", start_row + 1, col_ascent)))
    descent_count = int(_number(reader.value("STORE06", start_row + 2, col_ascent)))
    return LineResultCase(
        family=family,
        reference=reference,
        source="Workbook STORE06 defaults",
        flag_average=flag_average,
        ascent=LineResultBranch(
            name="Ascent Branch Results",
            rows=_read_branch(
                reader,
                start_row + 3,
                col_ascent,
                ascent_count,
                flag_average,
                _roller_counts(geometry.ascent_supports),
            ),
        ),
        descent=LineResultBranch(
            name="Descent Branch Results",
            rows=_read_branch(
                reader,
                start_row + 3,
                col_descent,
                descent_count,
                flag_average,
                _roller_counts(geometry.descent_supports) or _roller_counts(geometry.ascent_supports),
            ),
        ),
    )


def _first_reference(verification: VerificationState) -> ResultCaseRef:
    refs = active_result_cases(verification)
    if refs:
        return refs[0]
    hypothesis = verification.hypotheses[0]
    condition = verification.conditions[0]
    return ResultCaseRef(
        hypothesis_index=1,
        condition_index=1,
        state=MatrixState.OFF,
        label=f"{hypothesis.index}. {hypothesis.label} - {condition.label}",
    )


def _read_branch(
    reader: WorkbookReader,
    start_row: int,
    col: int,
    support_count: int,
    flag_average: int,
    roller_counts: dict[str, float],
) -> list[LineResultRow]:
    rows: list[LineResultRow] = []
    rr = start_row
    for index in range(1, max(support_count, 0)):
        span_values = [reader.value("STORE06", rr + offset, col) for offset in range(7)]
        rr += 7
        rows.append(_span_row(span_values))
        if flag_average == 0:
            span_values = [reader.value("STORE06", rr + offset, col) for offset in range(5)]
            rr += 5
            rows.append(_span_row(["", "", *span_values]))
        if index != support_count - 1:
            support_values = [reader.value("STORE06", rr + offset, col) for offset in range(5)]
            rr += 5
            rows.append(_support_row(support_values, roller_counts))
            if flag_average == 0:
                support_values = [reader.value("STORE06", rr + offset, col) for offset in range(4)]
                rr += 4
                rows.append(_support_row(["", *support_values], roller_counts))
    return rows


def _span_row(values: list[Any]) -> LineResultRow:
    return LineResultRow(
        span_start=str(values[0]).strip(),
        span_end=str(values[1]).strip(),
        valley_tension_da_n=_value(values[2]),
        mountain_tension_da_n=_value(values[3]),
        sag_m=_value(values[4]),
        valley_angle_deg=_degrees(values[5]),
        mountain_angle_deg=_degrees(values[6]),
    )


def _support_row(values: list[Any], roller_counts: dict[str, float]) -> LineResultRow:
    support_code = str(values[0]).strip()
    roller_count = roller_counts.get(support_code, "") if support_code else ""
    deviation = _degrees(values[2])
    pressure = _value(values[3])
    unit_deviation = deviation / roller_count if isinstance(deviation, float) and roller_count else ""
    unit_pressure = pressure / roller_count if isinstance(pressure, float) and roller_count else ""
    return LineResultRow(
        support_code=support_code,
        support_tension_da_n=_value(values[1]),
        total_deviation_deg=deviation,
        pressure_da_n=pressure,
        friction_da_n=_value(values[4]),
        roller_count=roller_count,
        unit_deviation_deg=unit_deviation,
        unit_pressure_da_n=unit_pressure,
    )


def _roller_counts(supports: list[Any]) -> dict[str, float]:
    counts: dict[str, float] = {}
    for support in supports:
        code = getattr(support, "code", "")
        rollers = getattr(support, "roller_quantity", "")
        if code and rollers != "":
            counts[str(code)] = float(rollers)
    return counts


def _value(value: Any) -> float | str:
    if value == "":
        return ""
    return float(value) if isinstance(value, (int, float)) else str(value)


def _degrees(value: Any) -> float | str:
    if value == "":
        return ""
    if isinstance(value, (int, float)):
        return float(value) * 180.0 / VBA_PI
    return str(value)


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
