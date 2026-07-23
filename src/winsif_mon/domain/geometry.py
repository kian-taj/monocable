from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from winsif_mon.domain.terrain import TerrainProfile, load_terrain_profile
from winsif_mon.workbook import WorkbookReader


@dataclass(slots=True)
class SupportRow:
    line_number: int
    code: str
    rope_distance_m: float | str
    ground_elevation_m: float | str
    tower_height_m: float | str
    rope_elevation_m: float | str
    roller_quantity: float | str


@dataclass(slots=True)
class FoundationDrawingRow:
    height_from_rope_to_base_m: float | str
    tower_inclination_deg: float | str
    foundation_protrusion_m: float | str
    plinth_height_m: float | str


@dataclass(slots=True)
class SpanRow:
    span_number: int
    code: str
    horizontal_distance_m: float | str
    height_difference_m: float | str


@dataclass(slots=True)
class GeometryState:
    ascent_supports: list[SupportRow]
    descent_supports: list[SupportRow]
    foundation_rows: list[FoundationDrawingRow]
    first_progressive_distance_m: float
    first_rope_height_m: float
    ascent_spans: list[SpanRow]
    descent_spans: list[SpanRow]


def load_geometry_state(workbook_path: Path | str = "MONOFUNI.xls") -> GeometryState:
    reader = WorkbookReader(workbook_path)
    return GeometryState(
        ascent_supports=_load_support_rows(reader, start_col=2),
        descent_supports=_load_support_rows(reader, start_col=9),
        foundation_rows=_load_foundation_rows(reader),
        first_progressive_distance_m=_number(reader.value("F04", 6, 5)),
        first_rope_height_m=_number(reader.value("F04", 7, 5)),
        ascent_spans=_load_span_rows(reader, start_col=2),
        descent_spans=_load_span_rows(reader, start_col=6),
    )


def convert_supports_to_spans(supports: list[SupportRow]) -> list[SpanRow]:
    rows = _non_empty_supports(supports)
    spans: list[SpanRow] = []
    for index, (previous, current) in enumerate(zip(rows, rows[1:]), start=1):
        previous_distance = _number(previous.rope_distance_m)
        current_distance = _number(current.rope_distance_m)
        previous_height = _number(previous.rope_elevation_m)
        current_height = _number(current.rope_elevation_m)
        spans.append(
            SpanRow(
                span_number=index,
                code=f"{previous.code}-{current.code}",
                horizontal_distance_m=current_distance - previous_distance,
                height_difference_m=current_height - previous_height,
            )
        )
    return spans


def convert_spans_to_supports(
    spans: list[SpanRow],
    first_progressive_distance_m: float,
    first_rope_height_m: float,
    terrain_profile: TerrainProfile,
    branch: str,
) -> list[SupportRow]:
    supports: list[SupportRow] = []
    distance = first_progressive_distance_m
    rope_height = first_rope_height_m
    first_code = _first_support_code(spans)
    supports.append(
        _support_from_distance(
            line_number=1,
            code=first_code,
            distance=distance,
            rope_height=rope_height,
            terrain_profile=terrain_profile,
            branch=branch,
        )
    )
    for index, span in enumerate(_non_empty_spans(spans), start=2):
        distance += _number(span.horizontal_distance_m)
        rope_height += _number(span.height_difference_m)
        code = _last_support_code(span.code) or f"S{index}"
        supports.append(
            _support_from_distance(
                line_number=index,
                code=code,
                distance=distance,
                rope_height=rope_height,
                terrain_profile=terrain_profile,
                branch=branch,
            )
        )
    return supports


def update_support_ground_from_terrain(
    row: SupportRow,
    terrain_profile: TerrainProfile,
    branch: str,
) -> SupportRow:
    distance = _number(row.rope_distance_m)
    ground = interpolate_ground_elevation(terrain_profile, branch, distance)
    tower_height = ""
    if row.rope_elevation_m != "":
        tower_height = _number(row.rope_elevation_m) - ground
    return replace(row, ground_elevation_m=ground, tower_height_m=tower_height)


def update_support_rope_elevation(row: SupportRow) -> SupportRow:
    if row.ground_elevation_m == "" or row.tower_height_m == "":
        return row
    return replace(row, rope_elevation_m=_number(row.ground_elevation_m) + _number(row.tower_height_m))


def update_support_tower_height(row: SupportRow) -> SupportRow:
    if row.rope_elevation_m == "" or row.ground_elevation_m == "":
        return row
    return replace(row, tower_height_m=_number(row.rope_elevation_m) - _number(row.ground_elevation_m))


def interpolate_ground_elevation(profile: TerrainProfile, branch: str, distance_m: float) -> float:
    points = profile.ascent if branch == "ascent" else profile.descent
    numeric = [
        (_number(point.progressive_distance_m), _number(point.ground_elevation_m))
        for point in points
        if point.progressive_distance_m != "" and point.ground_elevation_m != ""
    ]
    if not numeric:
        return 0.0
    if distance_m <= numeric[0][0]:
        return numeric[0][1]
    for (previous_distance, previous_elevation), (next_distance, next_elevation) in zip(numeric, numeric[1:]):
        if distance_m <= next_distance:
            if next_distance == previous_distance:
                return previous_elevation
            ratio = (distance_m - previous_distance) / (next_distance - previous_distance)
            return previous_elevation + ratio * (next_elevation - previous_elevation)
    return numeric[-1][1]


def _load_support_rows(reader: WorkbookReader, start_col: int) -> list[SupportRow]:
    rows: list[SupportRow] = []
    for excel_row in range(15, 95):
        line_number = int(_number(reader.value("F03", excel_row, start_col)) or len(rows) + 1)
        rows.append(
            SupportRow(
                line_number=line_number,
                code=str(reader.value("F03", excel_row, start_col + 1)).strip(),
                rope_distance_m=reader.value("F03", excel_row, start_col + 2),
                ground_elevation_m=reader.value("F03", excel_row, start_col + 3),
                tower_height_m=reader.value("F03", excel_row, start_col + 4),
                rope_elevation_m=reader.value("F03", excel_row, start_col + 5),
                roller_quantity=reader.value("F03", excel_row, start_col + 6),
            )
        )
    return rows


def _load_foundation_rows(reader: WorkbookReader) -> list[FoundationDrawingRow]:
    rows: list[FoundationDrawingRow] = []
    for excel_row in range(15, 95):
        values = [reader.value("F03", excel_row, col) for col in range(16, 20)]
        if all(value == "" for value in values):
            rows.append(FoundationDrawingRow("", "", "", ""))
            continue
        rows.append(FoundationDrawingRow(*values))
    return rows


def _load_span_rows(reader: WorkbookReader, start_col: int) -> list[SpanRow]:
    rows: list[SpanRow] = []
    for excel_row in range(15, 95):
        span_number = int(_number(reader.value("F04", excel_row, start_col)) or len(rows) + 1)
        rows.append(
            SpanRow(
                span_number=span_number,
                code=str(reader.value("F04", excel_row, start_col + 1)).strip(),
                horizontal_distance_m=reader.value("F04", excel_row, start_col + 2),
                height_difference_m=reader.value("F04", excel_row, start_col + 3),
            )
        )
    return rows


def _support_from_distance(
    line_number: int,
    code: str,
    distance: float,
    rope_height: float,
    terrain_profile: TerrainProfile,
    branch: str,
) -> SupportRow:
    ground = interpolate_ground_elevation(terrain_profile, branch, distance)
    return SupportRow(
        line_number=line_number,
        code=code,
        rope_distance_m=distance,
        ground_elevation_m=ground,
        tower_height_m=rope_height - ground,
        rope_elevation_m=rope_height,
        roller_quantity="",
    )


def _non_empty_supports(rows: list[SupportRow]) -> list[SupportRow]:
    return [row for row in rows if row.code or row.rope_distance_m != "" or row.rope_elevation_m != ""]


def _non_empty_spans(rows: list[SpanRow]) -> list[SpanRow]:
    return [row for row in rows if row.code or row.horizontal_distance_m != "" or row.height_difference_m != ""]


def _first_support_code(spans: list[SpanRow]) -> str:
    for span in spans:
        first, _separator, _last = span.code.partition("-")
        if first:
            return first
    return "S1"


def _last_support_code(code: str) -> str:
    _first, separator, last = code.partition("-")
    return last if separator else ""


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
