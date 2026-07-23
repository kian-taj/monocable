from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from winsif_mon.domain.f01_general import load_f01_defaults
from winsif_mon.domain.geometry import GeometryState, load_geometry_state
from winsif_mon.domain.line_results import ResultFamily
from winsif_mon.domain.max_min_results import load_max_min_case
from winsif_mon.domain.power_summary import load_power_summary_case
from winsif_mon.domain.verification import load_verification_state


@dataclass(frozen=True, slots=True)
class ReportSection:
    title: str
    rows: list[tuple[str, str]]


@dataclass(frozen=True, slots=True)
class GeneralReport:
    title: str
    sections: list[ReportSection]
    source: str = "Workbook defaults and STORE result tables"


def build_general_report(
    family: ResultFamily = ResultFamily.NORMAL,
    geometry: GeometryState | None = None,
    workbook_path: Path | str = "MONOFUNI.xls",
) -> GeneralReport:
    general = load_f01_defaults(workbook_path)
    verification = load_verification_state(workbook_path)
    geometry = geometry or load_geometry_state(workbook_path)
    max_min = load_max_min_case(family, workbook_path)
    power = load_power_summary_case(family, workbook_path)
    return GeneralReport(
        title=str(general.value("plant_description")),
        sections=[
            ReportSection(
                "Plant",
                [
                    ("Description", str(general.value("plant_description"))),
                    ("Location", str(general.value("plant_location"))),
                    ("Carrying capacity", f"{general.value('carrying_capacity_per_hour')} p/h"),
                    ("Running speed", f"{general.value('working_speed')} m/s"),
                ],
            ),
            ReportSection(
                "Line",
                [
                    ("Ascent supports", str(len([row for row in geometry.ascent_supports if row.code]))),
                    ("Descent supports", str(len([row for row in geometry.descent_supports if row.code]))),
                    ("Rope loop length", f"{verification.parameters.rope_loop_length_m:g} m"),
                    ("Car spacing", f"{verification.parameters.car_spacing_m:g} m"),
                ],
            ),
            ReportSection(
                "Results",
                [
                    ("Result family", family.label),
                    ("Max/min ascent rows", str(len(max_min.ascent_rows))),
                    ("Power summary rows", str(len(power.rows))),
                    ("Source", "Workbook STORE05/STORE13 defaults until native solver parity exists"),
                ],
            ),
        ],
    )
