from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from winsif_mon.domain.labels import load_labels
from winsif_mon.domain.line_results import ResultFamily
from winsif_mon.workbook import WorkbookReader


@dataclass(slots=True)
class PowerSummaryRow:
    direction: str
    case_index: int
    load_hypothesis: str
    plant_condition: str
    mean_force_da_n: float | str
    max_force_da_n: float | str
    motor_inertia_da_n: float | str
    motor_stress_da_n: float | str
    efficiency: float | str
    power_kw: float | str
    slip: float | str
    length_m: float | str
    total_tension_da_n: float | str


@dataclass(slots=True)
class PowerSummaryCase:
    family: ResultFamily
    rows: list[PowerSummaryRow]
    source: str = "Workbook STORE13 defaults"


START_ROWS = {
    ResultFamily.NORMAL: (4, 100),
    ResultFamily.PLUS_TEN: (200, 300),
    ResultFamily.MINUS_TEN: (400, 500),
    ResultFamily.HYDRAULIC: (600, 700),
}


def load_power_summary_case(
    family: ResultFamily = ResultFamily.NORMAL,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> PowerSummaryCase:
    reader = WorkbookReader(workbook_path)
    labels = load_labels(workbook_path)
    forward, reverse = START_ROWS[family]
    rows = _read_direction(reader, forward, "Forward", labels.load_hypotheses, labels.plant_conditions)
    rows.extend(_read_direction(reader, reverse, "Reverse", labels.load_hypotheses, labels.plant_conditions))
    return PowerSummaryCase(family=family, rows=rows)


def _read_direction(
    reader: WorkbookReader,
    start_row: int,
    direction: str,
    hypothesis_labels: list[str],
    condition_labels: list[str],
) -> list[PowerSummaryRow]:
    rows: list[PowerSummaryRow] = []
    for offset in range(90):
        row = start_row + offset
        col = 187
        values = [reader.value("STORE13", row, col + index) for index in range(15)]
        case_index = int(_number(values[0]))
        if case_index <= 0:
            continue
        hypothesis_index = (case_index - 1) // 6
        condition_index = (case_index - 1) % 6
        rows.append(
            PowerSummaryRow(
                direction=direction,
                case_index=case_index,
                load_hypothesis=hypothesis_labels[hypothesis_index],
                plant_condition=condition_labels[condition_index],
                mean_force_da_n=_value(values[1]),
                motor_inertia_da_n=_value(values[2]),
                motor_stress_da_n=_value(values[3]),
                efficiency=_value(values[4]),
                power_kw=_value(values[5]),
                slip=_value(values[6]),
                length_m=_value(values[7]),
                total_tension_da_n=_value(values[9]),
                max_force_da_n=_value(values[12]),
            )
        )
    return rows


def _value(value: Any) -> float | str:
    if value == "":
        return ""
    return float(value) if isinstance(value, (int, float)) else str(value)


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
