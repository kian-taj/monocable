from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from winsif_mon.workbook import WorkbookReader


@dataclass(slots=True)
class TerrainPoint:
    stake_number: int
    stake_code: str
    progressive_distance_m: float | str
    ground_elevation_m: float | str
    left_slope_percent: float | str = ""
    right_slope_percent: float | str = ""


@dataclass(slots=True)
class TerrainProfile:
    ascent: list[TerrainPoint]
    descent: list[TerrainPoint]


def load_terrain_profile(workbook_path: Path | str = "MONOFUNI.xls") -> TerrainProfile:
    reader = WorkbookReader(workbook_path)
    return TerrainProfile(
        ascent=_load_branch(reader, start_col=2, slope_start_col=12),
        descent=_load_branch(reader, start_col=7, slope_start_col=14),
    )


def _load_branch(reader: WorkbookReader, start_col: int, slope_start_col: int) -> list[TerrainPoint]:
    points: list[TerrainPoint] = []
    for row in range(15, 1115):
        distance = reader.value("F02", row, start_col + 2)
        elevation = reader.value("F02", row, start_col + 3)
        if distance == "" and elevation == "":
            continue
        if _as_float(distance) == 0.0 and _as_float(elevation) == 0.0:
            continue
        stake_number = reader.value("F02", row, start_col)
        points.append(
            TerrainPoint(
                stake_number=int(_as_float(stake_number) or len(points) + 1),
                stake_code=str(reader.value("F02", row, start_col + 1)).strip(),
                progressive_distance_m=distance,
                ground_elevation_m=elevation,
                left_slope_percent=reader.value("F02", row, slope_start_col),
                right_slope_percent=reader.value("F02", row, slope_start_col + 1),
            )
        )
    return points


def _as_float(value: Any) -> float:
    if value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0
