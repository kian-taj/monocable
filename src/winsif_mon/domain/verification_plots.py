from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from winsif_mon.domain.line_results import ResultCaseRef, ResultFamily, active_result_cases
from winsif_mon.domain.verification import VerificationState, load_verification_state
from winsif_mon.workbook import WorkbookReader


@dataclass(frozen=True, slots=True)
class PowerTracePoint:
    offset_m: float
    ascent_tension_da_n: float
    descent_tension_da_n: float
    motive_force_da_n: float


@dataclass(frozen=True, slots=True)
class PowerTraceCase:
    family: ResultFamily
    reference: ResultCaseRef
    forward: list[PowerTracePoint]
    reverse: list[PowerTracePoint]
    source: str = "Workbook STORE13 defaults"


TRACE_START_ROWS = {
    ResultFamily.NORMAL: (4, 100),
    ResultFamily.PLUS_TEN: (200, 300),
    ResultFamily.MINUS_TEN: (400, 500),
    ResultFamily.HYDRAULIC: (600, 700),
}


def load_power_trace_case(
    family: ResultFamily = ResultFamily.NORMAL,
    reference: ResultCaseRef | None = None,
    verification: VerificationState | None = None,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> PowerTraceCase:
    verification = verification or load_verification_state(workbook_path)
    reference = reference or active_result_cases(verification)[0]
    reader = WorkbookReader(workbook_path)
    forward_row, reverse_row = TRACE_START_ROWS[family]
    col = 1 + reference.store_index
    return PowerTraceCase(
        family=family,
        reference=reference,
        forward=_read_trace(reader, forward_row + 1, col),
        reverse=_read_trace(reader, reverse_row + 1, col),
    )


def _read_trace(reader: WorkbookReader, start_row: int, col: int) -> list[PowerTracePoint]:
    points: list[PowerTracePoint] = []
    row = start_row
    while row < start_row + 450:
        ascent = _number(reader.value("STORE13", row, col))
        descent = _number(reader.value("STORE13", row + 1, col))
        _stroke = _number(reader.value("STORE13", row + 2, col))
        offset = _number(reader.value("STORE13", row + 3, col))
        motive = _number(reader.value("STORE13", row + 4, col))
        if ascent <= 0 and descent == 0 and motive == 0:
            break
        points.append(PowerTracePoint(offset, ascent, descent, motive))
        row += 5
    return points


def _number(value) -> float:
    if value == "":
        return 0.0
    return float(value)
