from __future__ import annotations

from dataclasses import dataclass, fields
from pathlib import Path
from typing import Any

from winsif_mon.workbook import WorkbookReader


@dataclass(frozen=True, slots=True)
class FieldSpec:
    name: str
    label: str
    row: int
    value_col: int
    unit_col: int | None = None


GENERAL_FIELD_SPECS: tuple[FieldSpec, ...] = (
    FieldSpec("plant_description", "Description of the plant", 3, 4),
    FieldSpec("plant_location", "Plant installation location", 4, 4),
    FieldSpec("committee_details", "Committee/ further details", 5, 4),
    FieldSpec("carrying_rope_description", "Rope type description", 7, 4),
    FieldSpec("carrying_rope_outer_diameter", "Outer diameter", 8, 5, 4),
    FieldSpec("carrying_rope_metallic_section", "Metallic section", 9, 5, 4),
    FieldSpec("carrying_rope_unit_weight", "Unit weight", 10, 5, 4),
    FieldSpec("carrying_rope_outer_wire_diameter", "Outer wire diameter", 11, 5, 4),
    FieldSpec("carrying_rope_min_break_load", "Min.break load", 12, 5, 4),
    FieldSpec("carrying_rope_modulus_elasticity", "Modulus of elasticity of the rope", 13, 5, 4),
    FieldSpec("tension_rope_description", "Tension rope type description", 15, 4),
    FieldSpec("tension_rope_outer_diameter", "Outer diameter", 16, 5, 4),
    FieldSpec("tension_rope_metallic_section", "Metallic section", 17, 5, 4),
    FieldSpec("tension_rope_unit_weight", "Unit weight", 18, 5, 4),
    FieldSpec("tension_rope_max_wire_diameter", "Max.wire diameters", 19, 5, 4),
    FieldSpec("tension_rope_min_break_load", "Min.break load", 20, 5, 4),
    FieldSpec("tension_rope_branches", "Tension rope branches", 21, 5, 4),
    FieldSpec("signal_cable_description", "Description of the type of cable", 23, 4),
    FieldSpec("signal_cable_outer_diameter", "Outer diameter", 24, 5, 4),
    FieldSpec("signal_cable_metallic_section", "Metallic section of the cable holder rope", 25, 5, 4),
    FieldSpec("signal_cable_unit_weight", "Unit weight of the cable", 26, 5, 4),
    FieldSpec("signal_cable_ice_sleeve_thickness", "Ice sleeve thickness", 27, 5, 4),
    FieldSpec("signal_cable_min_break_load", "Min.break load", 28, 5, 4),
    FieldSpec("car_gap_ascent", "Gap (ascent branch) of NR. cars", 39, 5, 4),
    FieldSpec("rated_tightener_tension", "Rated value of the tightener tension", 40, 5, 4),
    FieldSpec("carrying_capacity_per_hour", "Carrying capacity/ hour", 41, 5, 4),
    FieldSpec("working_speed", "Working speed", 42, 5, 4),
    FieldSpec("persons_per_car", "Number of persons per car", 43, 5, 4),
    FieldSpec("person_weight", "Weight of a person", 44, 5, 4),
    FieldSpec("empty_car_weight", "Total weight of the empty car", 45, 5, 4),
    FieldSpec("laden_car_weight", "Total weight of the laden car", 46, 5, 4),
    FieldSpec("startup_acceleration", "Start up acceleration", 47, 5, 4),
    FieldSpec("stopping_deceleration_1", "Stopping deceleration (type 1)", 48, 5, 4),
    FieldSpec("stopping_deceleration_2", "Stopping deceleration (type 2)", 49, 5, 4),
    FieldSpec("stopping_deceleration_3", "Stopping deceleration (type 3)", 50, 5, 4),
    FieldSpec("power_unit_equivalent_weight", "Equivalent weight of the power unit and end wheels", 51, 5, 4),
    FieldSpec("power_unit_efficiency", "Power unit efficiency", 52, 5, 4),
    FieldSpec("friction_running", "Rope-rollers % friction (during running)", 53, 5, 4),
    FieldSpec("friction_braking", "Rope-rollers % friction (during braking)", 54, 5, 4),
    FieldSpec("station_rope_deviation_angle", "Angle of deviation of the rope at the station", 55, 5, 4),
    FieldSpec("acceleration_beam_force", "Driving force on acceleration beams", 56, 5, 4),
    FieldSpec("driving_pulley_diameter", "Driving pulley diameter", 57, 5, 4),
    FieldSpec("snub_pulley_diameter", "Snub pulley diameter", 58, 5, 4),
    FieldSpec("rope_distance", "Distance between ropes", 59, 5, 4),
    FieldSpec("car_type", "Type of car", 60, 5),
    FieldSpec("clamps_per_car", "Number of clamps per car", 61, 5, 4),
    FieldSpec("empty_car_crosswind_surface", "Empty car surface exposed to cross wind", 64, 5, 4),
    FieldSpec("laden_car_crosswind_surface", "Laden car surface exposed to cross wind", 65, 5, 4),
    FieldSpec("running_crosswind_thrust", "Cross wind thrust with the plant running", 66, 5, 4),
    FieldSpec("out_of_service_crosswind_thrust", "Cross wind thrust with the plant out of commission", 67, 5, 4),
    FieldSpec("roller_type", "Type of rollers", 68, 5),
    FieldSpec("support_roller_weight", "Weight of the support roller", 69, 5, 4),
    FieldSpec("compression_roller_weight", "Weight of the compression roller", 70, 5, 4),
    FieldSpec("support_roller_diameter", "Diameter of the support roller", 71, 5, 4),
    FieldSpec("compression_roller_diameter", "Diameter of the compression roller", 72, 5, 4),
    FieldSpec("max_support_roller_deviation", "Max.admitted deviation on the support roller", 73, 5, 4),
    FieldSpec("max_compression_roller_deviation", "Max.admitted deviation on the compression roller", 74, 5, 4),
    FieldSpec("max_support_roller_load", "Max.admitted load on the support roller", 75, 5, 4),
    FieldSpec("max_compression_roller_load", "Max.admitted load on the compression roller", 76, 5, 4),
    FieldSpec("double_acting_roller_type", "Type of roller for double acting roller assembly", 77, 5),
    FieldSpec("double_acting_roller_diameter", "Double acting roller assembly roller diameter", 78, 5, 4),
    FieldSpec("double_acting_roller_weight", "Weight of the roller for double acting roller assembly", 79, 5, 4),
    FieldSpec("double_acting_roller_max_load", "Max.admitted load for double acting roller", 80, 5, 4),
    FieldSpec("vehicle_vertical_height", "Vertical vehicle height", 81, 5, 4),
    FieldSpec("vehicle_width", "Width of the vehicle", 82, 5, 4),
    FieldSpec("vehicle_admissible_inclination", "Admissible vehicle inclination", 83, 5, 4),
    FieldSpec("pulley_envelopment_corner", "Envelopment corner of the rope on the pulley", 84, 5, 4),
    FieldSpec("pulley_rope_friction_coefficient", "Friction coefficient of the rope on the pulley", 85, 5, 4),
    FieldSpec("valley_station_time", "Time of way in the valley station", 86, 5, 4),
    FieldSpec("mountain_station_time", "Time of way in the mountain station", 87, 5, 4),
)


@dataclass(slots=True)
class F01GeneralData:
    values: dict[str, Any]
    units: dict[str, str]

    def value(self, name: str) -> Any:
        return self.values.get(name, "")

    def set_value(self, name: str, value: Any) -> None:
        self.values[name] = value


@dataclass(frozen=True, slots=True)
class RopeCatalogRow:
    description: str
    diameter: float
    metallic_section: float
    unit_weight: float
    wire_diameter: float
    breaking_load_1770: float
    breaking_load_1960: float
    breaking_load_2160: float
    minimum_failure_1770: float
    minimum_failure_1960: float
    minimum_failure_2160: float


def load_f01_defaults(workbook_path: Path | str = "MONOFUNI.xls") -> F01GeneralData:
    reader = WorkbookReader(workbook_path)
    values: dict[str, Any] = {}
    units: dict[str, str] = {}
    for spec in GENERAL_FIELD_SPECS:
        values[spec.name] = reader.value("F01", spec.row, spec.value_col)
        if spec.unit_col:
            units[spec.name] = str(reader.value("F01", spec.row, spec.unit_col)).strip()
    return F01GeneralData(values=values, units=units)


def load_rope_catalog(sheet_name: str, workbook_path: Path | str = "MONOFUNI.xls") -> list[RopeCatalogRow]:
    reader = WorkbookReader(workbook_path)
    rows: list[RopeCatalogRow] = []
    current_description = ""
    for row in range(10, 403):
        first = reader.value(sheet_name, row, 2)
        if isinstance(first, str) and first.strip():
            current_description = first.strip()
            continue
        if not isinstance(first, (int, float)) or first == 0:
            continue
        rows.append(
            RopeCatalogRow(
                description=current_description,
                diameter=float(first),
                metallic_section=_number(reader.value(sheet_name, row, 3)),
                unit_weight=_number(reader.value(sheet_name, row, 4)),
                wire_diameter=_number(reader.value(sheet_name, row, 5)),
                breaking_load_1770=_number(reader.value(sheet_name, row, 6)),
                breaking_load_1960=_number(reader.value(sheet_name, row, 7)),
                breaking_load_2160=_number(reader.value(sheet_name, row, 8)),
                minimum_failure_1770=_number(reader.value(sheet_name, row, 9)),
                minimum_failure_1960=_number(reader.value(sheet_name, row, 10)),
                minimum_failure_2160=_number(reader.value(sheet_name, row, 11)),
            )
        )
    return rows


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)


def dataclass_field_names(cls: type) -> list[str]:
    return [field.name for field in fields(cls)]
