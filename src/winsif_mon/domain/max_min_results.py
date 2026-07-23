from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from winsif_mon.domain.line_results import ResultFamily
from winsif_mon.workbook import WorkbookReader

VBA_PI = 3.141


@dataclass(slots=True)
class MaxMinRow:
    span_start: str = ""
    span_end: str = ""
    span_state: str = ""
    tension_da_n: float | str = ""
    sag_m: float | str = ""
    valley_angle_deg: float | str = ""
    mountain_angle_deg: float | str = ""
    support_code: str = ""
    support_state: str = ""
    support_tension_da_n: float | str = ""
    deviation_deg: float | str = ""
    pressure_da_n: float | str = ""
    friction_da_n: float | str = ""
    roller_count: float | str = ""
    unit_deviation_deg: float | str = ""
    unit_pressure_da_n: float | str = ""
    test: str = ""


@dataclass(slots=True)
class MaxMinCase:
    family: ResultFamily
    ascent_rows: list[MaxMinRow]
    descent_rows: list[MaxMinRow]
    source: str = "Workbook STORE05 defaults"


START_ROWS = {
    ResultFamily.NORMAL: 3,
    ResultFamily.HYDRAULIC: 42,
    ResultFamily.PLUS_TEN: 82,
    ResultFamily.MINUS_TEN: 122,
}


def load_max_min_case(
    family: ResultFamily = ResultFamily.NORMAL,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> MaxMinCase:
    reader = WorkbookReader(workbook_path)
    start = START_ROWS[family]
    ascent_count = int(_number(reader.value("STORE05", start, 2)))
    descent_count = int(_number(reader.value("STORE05", start + 1, 2)))
    return MaxMinCase(
        family=family,
        ascent_rows=_read_branch(reader, start + 2, 3, ascent_count),
        descent_rows=_read_branch(reader, start + 2, 84, descent_count),
    )


def _read_branch(reader: WorkbookReader, start_row: int, start_col: int, support_count: int) -> list[MaxMinRow]:
    rows: list[MaxMinRow] = []
    for support_index in range(2, max(support_count, 0) + 1):
        row = start_row
        col = start_col + support_index - 2
        span_start = str(reader.value("STORE05", row, col)).strip()
        span_end = str(reader.value("STORE05", row + 1, col)).strip()
        rows.append(
            MaxMinRow(
                span_start=span_start,
                span_end=span_end,
                span_state="max",
                tension_da_n=_value(reader.value("STORE05", row + 2, col)),
                sag_m=_value(reader.value("STORE05", row + 3, col)),
                valley_angle_deg=_degrees(reader.value("STORE05", row + 4, col)),
                mountain_angle_deg=_degrees(reader.value("STORE05", row + 5, col)),
            )
        )
        rows.append(
            MaxMinRow(
                span_state="min",
                tension_da_n=_value(reader.value("STORE05", row + 6, col)),
                sag_m=_value(reader.value("STORE05", row + 7, col)),
                valley_angle_deg=_degrees(reader.value("STORE05", row + 8, col)),
                mountain_angle_deg=_degrees(reader.value("STORE05", row + 9, col)),
            )
        )
        if support_index != support_count:
            support_row = row + 10
            rows.append(_support_row(reader, support_row, col, "max"))
            rows.append(_support_row(reader, support_row + 9, col, "min"))
    return rows


def _support_row(reader: WorkbookReader, row: int, col: int, state: str) -> MaxMinRow:
    return MaxMinRow(
        support_code=str(reader.value("STORE05", row, col)).strip(),
        support_state=state,
        support_tension_da_n=_value(reader.value("STORE05", row + 1, col)),
        deviation_deg=_degrees(reader.value("STORE05", row + 2, col)),
        pressure_da_n=_value(reader.value("STORE05", row + 3, col)),
        friction_da_n=_value(reader.value("STORE05", row + 4, col)),
        roller_count=_value(reader.value("STORE05", row + 5, col)),
        unit_deviation_deg=_degrees(reader.value("STORE05", row + 6, col)),
        unit_pressure_da_n=_value(reader.value("STORE05", row + 7, col)),
        test=str(reader.value("STORE05", row + 8, col)).strip(),
    )


def _value(value: Any) -> float | str:
    if value == "":
        return ""
    return float(value) if isinstance(value, (int, float)) else str(value)


def _degrees(value: Any) -> float | str:
    if value == "":
        return ""
    if isinstance(value, (int, float)):
        return float(value) * 180 / VBA_PI
    return str(value)


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
