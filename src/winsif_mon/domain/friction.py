from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

from winsif_mon.domain.geometry import GeometryState
from winsif_mon.workbook import WorkbookReader


class FrictionMode(str, Enum):
    PERCENT = "%"
    ABSOLUTE = "A"

    @property
    def label(self) -> str:
        if self is FrictionMode.PERCENT:
            return "Percentage (%)"
        return "Absolute (N/roller)"

    @property
    def unit(self) -> str:
        if self is FrictionMode.PERCENT:
            return "%"
        return "N/roller"


@dataclass(slots=True)
class FrictionSettings:
    default_mode: FrictionMode
    steady_value: float
    braking_value: float


@dataclass(slots=True)
class FrictionAssignmentRow:
    support_code: str
    mode: FrictionMode
    steady_value: float
    braking_value: float


@dataclass(slots=True)
class FrictionAssignmentState:
    settings: FrictionSettings
    ascent_rows: list[FrictionAssignmentRow]
    descent_rows: list[FrictionAssignmentRow]


def load_friction_assignment_state(workbook_path: Path | str = "MONOFUNI.xls") -> FrictionAssignmentState:
    reader = WorkbookReader(workbook_path)
    settings = FrictionSettings(
        default_mode=FrictionMode.PERCENT,
        steady_value=_number(reader.value("F07", 4, 9)),
        braking_value=_number(reader.value("F07", 5, 9)),
    )
    return FrictionAssignmentState(
        settings=settings,
        ascent_rows=_load_rows(reader, code_col=2, mode_col=3, steady_col=4, braking_col=5),
        descent_rows=_load_rows(reader, code_col=7, mode_col=8, steady_col=9, braking_col=10),
    )


def assign_default_friction(
    supports: list[Any],
    settings: FrictionSettings,
) -> list[FrictionAssignmentRow]:
    rows: list[FrictionAssignmentRow] = []
    usable = [support for support in supports if getattr(support, "code", "")]
    for support in usable[1:-1]:
        rows.append(
            FrictionAssignmentRow(
                support_code=str(support.code),
                mode=settings.default_mode,
                steady_value=settings.steady_value,
                braking_value=settings.braking_value,
            )
        )
    return rows


def reset_friction_from_geometry(
    state: FrictionAssignmentState,
    geometry: GeometryState,
) -> None:
    state.ascent_rows = assign_default_friction(geometry.ascent_supports, state.settings)
    descent_supports = [support for support in geometry.descent_supports if support.code]
    if len(descent_supports) < 2:
        descent_supports = [support for support in geometry.ascent_supports if support.code]
    state.descent_rows = assign_default_friction(descent_supports, state.settings)


def validate_friction_state(state: FrictionAssignmentState) -> list[str]:
    errors: list[str] = []
    for branch_name, rows in (("ascent", state.ascent_rows), ("descent", state.descent_rows)):
        for index, row in enumerate(rows, start=1):
            if not row.support_code.strip():
                errors.append(f"{branch_name} row {index}: support code is required")
            if row.mode not in (FrictionMode.PERCENT, FrictionMode.ABSOLUTE):
                errors.append(f"{branch_name} row {index}: invalid friction mode")
            if row.steady_value < 0:
                errors.append(f"{branch_name} row {index}: steady value must be non-negative")
            if row.braking_value < 0:
                errors.append(f"{branch_name} row {index}: braking value must be non-negative")
    return errors


def _load_rows(
    reader: WorkbookReader,
    code_col: int,
    mode_col: int,
    steady_col: int,
    braking_col: int,
) -> list[FrictionAssignmentRow]:
    rows: list[FrictionAssignmentRow] = []
    for excel_row in range(12, 93):
        code = str(reader.value("F07", excel_row, code_col)).strip()
        if not code:
            continue
        rows.append(
            FrictionAssignmentRow(
                support_code=code,
                mode=_mode(reader.value("F07", excel_row, mode_col)),
                steady_value=_number(reader.value("F07", excel_row, steady_col)),
                braking_value=_number(reader.value("F07", excel_row, braking_col)),
            )
        )
    return rows


def _mode(value: Any) -> FrictionMode:
    if str(value).strip().upper() == FrictionMode.ABSOLUTE.value:
        return FrictionMode.ABSOLUTE
    return FrictionMode.PERCENT


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
