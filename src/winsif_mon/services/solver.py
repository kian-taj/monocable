from __future__ import annotations

from dataclasses import dataclass, replace
from enum import Enum
from math import cos, pi, sqrt
from typing import Any

from winsif_mon.domain.f01_general import F01GeneralData, load_f01_defaults
from winsif_mon.domain.friction import FrictionAssignmentState
from winsif_mon.domain.geometry import GeometryState, SpanRow, SupportRow, load_geometry_state
from winsif_mon.domain.terrain import TerrainProfile, load_terrain_profile
from winsif_mon.domain.verification import (
    CustomLoadState,
    MatrixState,
    VerificationState,
    build_calculation_setup,
    load_custom_load_state,
    load_verification_state,
)
from winsif_mon.models import CalculationMode
from winsif_mon.workbook import WorkbookReader

VBA_PI = 3.141
VBA_G = 0.981


class SolverMode(str, Enum):
    NORMAL = "normal"
    HYDRAULIC_PRECHECK = "hydraulic_precheck"
    HYDRAULIC_FINAL = "hydraulic_final"
    ANCHORED = "anchored"
    VARIABLE_TENSION = "variable_tension"


@dataclass(frozen=True, slots=True)
class GeneralInput:
    plant_name: str
    location: str
    rope_unit_weight_da_n_m: float
    rope_modulus_n_mm2: float
    rope_metallic_area_mm2: float
    rated_tension_da_n: float
    capacity_p_h: float
    speed_m_s: float
    persons_per_car: float
    empty_car_weight: float
    loaded_car_weight: float
    acceleration_m_s2: float
    deceleration_1_m_s2: float
    deceleration_2_m_s2: float
    deceleration_3_m_s2: float
    drive_equivalent_weight: float
    drive_efficiency: float
    friction_running_percent: float
    friction_braking_percent: float
    station_deviation_angle: float
    acceleration_beam_force: float
    drive_pulley_diameter_mm: float
    return_pulley_diameter_mm: float
    support_roller_weight: float
    compression_roller_weight: float
    max_support_roller_load: float
    max_compression_roller_load: float
    zeta: str = "F"
    drive_station: str = "VALLE"
    tension_station: str = "MONTE"
    rope_loop_type: str = "T"
    vehicles_per_group: float = 1.0
    groups_per_branch: float = 0.0
    vehicle_spacing_m: float = 0.0


@dataclass(frozen=True, slots=True)
class LineBranch:
    name: str
    supports: list[SupportRow]
    spans: list[SpanRow]
    support_progressive_m: list[float]
    developed_length_m: float


@dataclass(frozen=True, slots=True)
class FrictionRow:
    tower_code: str
    friction_type: str
    steady_value: float
    braking_value: float


@dataclass(frozen=True, slots=True)
class LineInput:
    general: GeneralInput
    terrain: TerrainProfile
    geometry: GeometryState
    verification: VerificationState
    custom_loads: CustomLoadState
    ascent: LineBranch
    descent: LineBranch
    ascent_friction: list[FrictionRow]
    descent_friction: list[FrictionRow]


@dataclass(frozen=True, slots=True)
class RunParameters:
    rope_loop_length_m: float
    total_cars: float
    car_spacing_m: float
    running_speed_m_s: float
    carrying_capacity_p_h: float
    step_m: float
    step_count: int
    vehicle_spacing_m: float
    vehicles_per_group: int


@dataclass(frozen=True, slots=True)
class VehiclePositions:
    ascent: list[float]
    descent: list[float]


@dataclass(frozen=True, slots=True)
class LoadPreparation:
    hypothesis_index: int
    condition_index: int
    ascent_loads_da_n: list[float]
    descent_loads_da_n: list[float]


@dataclass(frozen=True, slots=True)
class SolverRunResult:
    mode: SolverMode
    run_parameters: RunParameters
    active_cases: list[tuple[int, int, MatrixState]]
    vehicle_positions: VehiclePositions
    prepared_loads: list[LoadPreparation]
    status: str
    native_iterative_result: Any | None = None
    parity_complete: bool = False


def mode_from_calculation_mode(mode: CalculationMode) -> SolverMode:
    return SolverMode(mode.value)


def build_line_input(
    general_data: F01GeneralData | None = None,
    terrain: TerrainProfile | None = None,
    geometry: GeometryState | None = None,
    verification: VerificationState | None = None,
    custom_loads: CustomLoadState | None = None,
    friction_state: FrictionAssignmentState | None = None,
    workbook_path: str = "MONOFUNI.xls",
) -> LineInput:
    general_data = general_data or load_f01_defaults(workbook_path)
    terrain = terrain or load_terrain_profile(workbook_path)
    geometry = geometry or load_geometry_state(workbook_path)
    verification = verification or load_verification_state(workbook_path)
    custom_loads = custom_loads or load_custom_load_state(workbook_path, geometry)
    general = _general_input(general_data)
    ascent = _branch("ascent", geometry.ascent_supports)
    descent_supports = _non_empty_supports(geometry.descent_supports)
    if len(descent_supports) < 2:
        descent_supports = _descent_fallback_supports(geometry)
    descent = _branch("descent", descent_supports)
    if friction_state is None:
        ascent_friction, descent_friction = load_friction_defaults(workbook_path)
    else:
        ascent_friction = _friction_rows(friction_state.ascent_rows)
        descent_friction = _friction_rows(friction_state.descent_rows)
    return LineInput(
        general=general,
        terrain=terrain,
        geometry=geometry,
        verification=verification,
        custom_loads=custom_loads,
        ascent=ascent,
        descent=descent,
        ascent_friction=ascent_friction,
        descent_friction=descent_friction,
    )


def run_solver_setup(
    mode: SolverMode,
    line_input: LineInput,
) -> SolverRunResult:
    build_calculation_setup(line_input.verification, line_input.custom_loads)
    run_parameters = derive_run_parameters(line_input)
    active_cases = _active_cases(line_input.verification)
    positions = position_vehicles(
        offset_m=0.0,
        ascent_length_m=line_input.ascent.developed_length_m,
        descent_length_m=line_input.descent.developed_length_m,
        drive_pulley_diameter_mm=line_input.general.drive_pulley_diameter_mm,
        car_spacing_m=run_parameters.vehicle_spacing_m,
        vehicles_per_group=run_parameters.vehicles_per_group,
        intra_group_spacing_m=run_parameters.car_spacing_m,
    )
    prepared_loads = [
        prepare_loads(line_input, hypothesis_index=hypothesis, condition_index=condition)
        for hypothesis, condition, state in active_cases
        if state is MatrixState.NORMAL
    ]
    native_iterative_result = None
    if mode in {
        SolverMode.NORMAL,
        SolverMode.VARIABLE_TENSION,
        SolverMode.ANCHORED,
        SolverMode.HYDRAULIC_PRECHECK,
        SolverMode.HYDRAULIC_FINAL,
    }:
        from winsif_mon.services.iterative_solver import run_iterative_solver

        native_iterative_result = run_iterative_solver(mode, line_input)
    return SolverRunResult(
        mode=mode,
        run_parameters=run_parameters,
        active_cases=active_cases,
        vehicle_positions=positions,
        prepared_loads=prepared_loads,
        status=(
            "Input assembly, line geometry, vehicle positioning, friction defaults, load preparation, and the "
            "native iterative solver scaffold completed. Workbook stores remain the golden results until parity "
            "tests pass."
        ),
        native_iterative_result=native_iterative_result,
    )


def derive_run_parameters(line_input: LineInput) -> RunParameters:
    general = line_input.general
    rope_loop_length = (
        line_input.ascent.developed_length_m
        + line_input.descent.developed_length_m
        + VBA_PI * (general.drive_pulley_diameter_mm / 1000.0 + general.return_pulley_diameter_mm / 1000.0) / 2.0
    )
    if general.zeta == "I":
        spacing = rope_loop_length / (2.0 * general.groups_per_branch) if general.groups_per_branch else 0.0
        total_cars = 2.0 * general.groups_per_branch * general.vehicles_per_group
        effective_capacity = general.capacity_p_h
        vehicle_spacing = general.vehicle_spacing_m
        vehicles_per_group = max(int(general.vehicles_per_group), 1)
    else:
        theoretical_spacing = 3600.0 / general.capacity_p_h * general.persons_per_car * general.speed_m_s
        if general.zeta == "F":
            total_cars = _vba_vehicle_round(rope_loop_length / theoretical_spacing)
            spacing = theoretical_spacing
            effective_capacity = general.capacity_p_h
        else:
            total_cars = _vba_vehicle_round(
                (line_input.ascent.developed_length_m + line_input.descent.developed_length_m) / theoretical_spacing
            )
            spacing = theoretical_spacing
            effective_capacity = general.capacity_p_h
        vehicle_spacing = spacing
        vehicles_per_group = 1
    advancement_steps = int(line_input.verification.parameters.advancement_steps or 10)
    step = spacing / advancement_steps if advancement_steps else 0.0
    return RunParameters(
        rope_loop_length_m=rope_loop_length,
        total_cars=total_cars,
        car_spacing_m=spacing,
        running_speed_m_s=general.speed_m_s,
        carrying_capacity_p_h=effective_capacity,
        step_m=step,
        step_count=advancement_steps,
        vehicle_spacing_m=vehicle_spacing,
        vehicles_per_group=vehicles_per_group,
    )


def position_vehicles(
    offset_m: float,
    ascent_length_m: float,
    descent_length_m: float,
    drive_pulley_diameter_mm: float,
    car_spacing_m: float,
    vehicles_per_group: int,
    intra_group_spacing_m: float,
) -> VehiclePositions:
    vehicles_per_group = max(vehicles_per_group, 1)
    free_length = car_spacing_m - (vehicles_per_group - 1) * intra_group_spacing_m
    ascent = _vehicle_branch_positions(offset_m, ascent_length_m, free_length, vehicles_per_group, intra_group_spacing_m)

    if offset_m <= free_length:
        descent_start = car_spacing_m - offset_m - VBA_PI * drive_pulley_diameter_mm / 1000.0 / 2.0 - (
            vehicles_per_group - 1
        ) * intra_group_spacing_m
        if descent_start >= 0:
            descent = _vehicle_branch_positions(
                descent_start,
                descent_length_m,
                free_length,
                vehicles_per_group,
                intra_group_spacing_m,
            )
            return VehiclePositions(ascent=ascent, descent=descent)
    descent_start = car_spacing_m - offset_m - VBA_PI * drive_pulley_diameter_mm / 1000.0 / 2.0
    if descent_start < 0:
        descent_start = free_length + descent_start
    else:
        group_count = int(descent_start / intra_group_spacing_m)
        if group_count > 0:
            group_count -= 1
        descent_start = descent_start - group_count * intra_group_spacing_m
    descent = _vehicle_branch_positions(
        descent_start,
        descent_length_m,
        free_length,
        vehicles_per_group,
        intra_group_spacing_m,
    )
    return VehiclePositions(ascent=ascent, descent=descent)


def prepare_loads(line_input: LineInput, hypothesis_index: int, condition_index: int) -> LoadPreparation:
    ascent_span_count = max(len(line_input.ascent.spans), 0)
    descent_span_count = max(len(line_input.descent.spans), 0)
    if hypothesis_index == 5:
        return LoadPreparation(hypothesis_index, condition_index, [0.0] * ascent_span_count, [0.0] * descent_span_count)
    if hypothesis_index > 5:
        load_row_index = hypothesis_index - 6
        ascent: list[float] = []
        descent: list[float] = []
        for row in line_input.custom_loads.rows[: max(ascent_span_count, descent_span_count)]:
            ascent.append(_number(row.loads[load_row_index * 2]) * VBA_G)
            descent.append(_number(row.loads[load_row_index * 2 + 1]) * VBA_G)
        return LoadPreparation(hypothesis_index, condition_index, ascent[:ascent_span_count], descent[:descent_span_count])

    qs = {1: 1, 2: 0, 3: 0, 4: 1}
    qd = {1: 0, 2: 0, 3: 1, 4: 1}
    ascent_loaded = bool(qs[hypothesis_index])
    descent_loaded = bool(qd[hypothesis_index])
    ascent_weight = line_input.general.loaded_car_weight if ascent_loaded else line_input.general.empty_car_weight
    descent_weight = line_input.general.loaded_car_weight if descent_loaded else line_input.general.empty_car_weight
    return LoadPreparation(
        hypothesis_index,
        condition_index,
        [ascent_weight * VBA_G] * ascent_span_count,
        [descent_weight * VBA_G] * descent_span_count,
    )


def load_friction_defaults(workbook_path: str = "MONOFUNI.xls") -> tuple[list[FrictionRow], list[FrictionRow]]:
    reader = WorkbookReader(workbook_path)
    ascent: list[FrictionRow] = []
    descent: list[FrictionRow] = []
    for row in range(12, 93):
        ascent_code = str(reader.value("F07", row, 2)).strip()
        descent_code = str(reader.value("F07", row, 7)).strip()
        if ascent_code:
            ascent.append(
                FrictionRow(
                    tower_code=ascent_code,
                    friction_type=str(reader.value("F07", row, 3)).strip() or "%",
                    steady_value=_number(reader.value("F07", row, 4)),
                    braking_value=_number(reader.value("F07", row, 5)),
                )
            )
        if descent_code:
            descent.append(
                FrictionRow(
                    tower_code=descent_code,
                    friction_type=str(reader.value("F07", row, 8)).strip() or "%",
                    steady_value=_number(reader.value("F07", row, 9)),
                    braking_value=_number(reader.value("F07", row, 10)),
                )
            )
    return ascent, descent


def _general_input(data: F01GeneralData) -> GeneralInput:
    return GeneralInput(
        plant_name=str(data.value("plant_description")),
        location=str(data.value("plant_location")),
        rope_unit_weight_da_n_m=_number(data.value("carrying_rope_unit_weight")) * VBA_G,
        rope_modulus_n_mm2=_number(data.value("carrying_rope_modulus_elasticity")),
        rope_metallic_area_mm2=_number(data.value("carrying_rope_metallic_section")),
        rated_tension_da_n=_number(data.value("rated_tightener_tension")),
        capacity_p_h=_number(data.value("carrying_capacity_per_hour")),
        speed_m_s=_number(data.value("working_speed")),
        persons_per_car=_number(data.value("persons_per_car")),
        empty_car_weight=_number(data.value("empty_car_weight")),
        loaded_car_weight=_number(data.value("laden_car_weight")),
        acceleration_m_s2=_number(data.value("startup_acceleration")),
        deceleration_1_m_s2=_number(data.value("stopping_deceleration_1")),
        deceleration_2_m_s2=_number(data.value("stopping_deceleration_2")),
        deceleration_3_m_s2=_number(data.value("stopping_deceleration_3")),
        drive_equivalent_weight=_number(data.value("power_unit_equivalent_weight")),
        drive_efficiency=_number(data.value("power_unit_efficiency")),
        friction_running_percent=_number(data.value("friction_running")),
        friction_braking_percent=_number(data.value("friction_braking")),
        station_deviation_angle=_number(data.value("station_rope_deviation_angle")),
        acceleration_beam_force=_number(data.value("acceleration_beam_force")),
        drive_pulley_diameter_mm=_number(data.value("driving_pulley_diameter")),
        return_pulley_diameter_mm=_number(data.value("snub_pulley_diameter")),
        support_roller_weight=_number(data.value("support_roller_weight")),
        compression_roller_weight=_number(data.value("compression_roller_weight")),
        max_support_roller_load=_number(data.value("max_support_roller_load")),
        max_compression_roller_load=_number(data.value("max_compression_roller_load")),
        vehicle_spacing_m=_number(data.value("car_gap_ascent")),
    )


def _branch(name: str, supports: list[SupportRow]) -> LineBranch:
    non_empty = _non_empty_supports(supports)
    spans: list[SpanRow] = []
    progressive = [0.0]
    developed = 0.0
    for index, (previous, current) in enumerate(zip(non_empty, non_empty[1:]), start=1):
        horizontal = _number(current.rope_distance_m) - _number(previous.rope_distance_m)
        vertical = _number(current.rope_elevation_m) - _number(previous.rope_elevation_m)
        chord = sqrt(horizontal**2 + vertical**2)
        developed += chord
        progressive.append(developed)
        spans.append(
            SpanRow(
                span_number=index,
                code=f"{previous.code}-{current.code}",
                horizontal_distance_m=horizontal,
                height_difference_m=vertical,
            )
        )
    return LineBranch(name=name, supports=non_empty, spans=spans, support_progressive_m=progressive, developed_length_m=developed)


def _vehicle_branch_positions(
    start: float,
    branch_length: float,
    free_length: float,
    vehicles_per_group: int,
    intra_group_spacing: float,
) -> list[float]:
    positions: list[float] = [start]
    while True:
        for _index in range(vehicles_per_group - 1):
            next_position = positions[-1] + intra_group_spacing
            if next_position > branch_length:
                return positions
            positions.append(next_position)
        next_position = positions[-1] + free_length
        if next_position > branch_length:
            return positions
        positions.append(next_position)


def _active_cases(verification: VerificationState) -> list[tuple[int, int, MatrixState]]:
    active: list[tuple[int, int, MatrixState]] = []
    for hypothesis_index, row in enumerate(verification.matrix, start=1):
        for condition_index, state in enumerate(row, start=1):
            if state is not MatrixState.OFF:
                active.append((hypothesis_index, condition_index, state))
    return active


def _friction_rows(rows) -> list[FrictionRow]:
    return [
        FrictionRow(
            tower_code=row.support_code,
            friction_type=row.mode.value,
            steady_value=float(row.steady_value),
            braking_value=float(row.braking_value),
        )
        for row in rows
        if row.support_code
    ]


def _non_empty_supports(rows: list[SupportRow]) -> list[SupportRow]:
    return [row for row in rows if row.code or row.rope_distance_m != "" or row.rope_elevation_m != ""]


def _descent_fallback_supports(geometry: GeometryState) -> list[SupportRow]:
    ascent = _non_empty_supports(geometry.ascent_supports)
    rows: list[SupportRow] = []
    for index, support in enumerate(ascent):
        descent_rollers = ""
        if index < len(geometry.descent_supports):
            descent_rollers = geometry.descent_supports[index].roller_quantity
        rows.append(
            replace(
                support,
                roller_quantity=descent_rollers if descent_rollers != "" else support.roller_quantity,
            )
        )
    return rows


def _vba_vehicle_round(value: float) -> float:
    remainder = abs(int(value) - value)
    if remainder > 0:
        if remainder > 0.5:
            return float(int(value) + 1)
        return float(int(value) - 1)
    return float(value)


def _number(value: Any) -> float:
    if value == "":
        return 0.0
    return float(value)
