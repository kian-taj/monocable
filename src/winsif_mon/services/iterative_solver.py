from __future__ import annotations

from dataclasses import dataclass
from math import atan, cos, sin, sqrt, tan

from winsif_mon.domain.line_results import (
    LineResultBranch,
    LineResultCase,
    LineResultRow,
    ResultCaseRef,
    ResultFamily,
)
from winsif_mon.domain.max_min_results import MaxMinCase, MaxMinRow
from winsif_mon.domain.power_summary import PowerSummaryCase, PowerSummaryRow
from winsif_mon.domain.verification import MatrixState
from winsif_mon.domain.verification_plots import PowerTraceCase, PowerTracePoint
from winsif_mon.services.solver import (
    LineBranch,
    LineInput,
    SolverMode,
    VBA_G,
    VBA_PI,
    derive_run_parameters,
    position_vehicles,
    prepare_loads,
)


@dataclass(frozen=True, slots=True)
class IterativeCaseResult:
    family: ResultFamily
    reference: ResultCaseRef
    power_trace: PowerTraceCase
    forward_line_results: LineResultCase
    power_summary: PowerSummaryCase

    @property
    def forward_trace(self) -> list[PowerTracePoint]:
        return self.power_trace.forward

    @property
    def reverse_trace(self) -> list[PowerTracePoint]:
        return self.power_trace.reverse


@dataclass(frozen=True, slots=True)
class IterativeSolverResult:
    mode: SolverMode
    cases: list[IterativeCaseResult]
    max_min_results: dict[ResultFamily, MaxMinCase]
    hydraulic_reference_values: list[list[float]] | None = None
    parity_complete: bool = False


@dataclass(slots=True)
class _MotionParameters:
    acceleration: float
    friction_percent: float


@dataclass(slots=True)
class _SpanCalculation:
    valley_tension: float
    mountain_tension: float
    support_tension: float
    deviation_rad: float
    pressure_da_n: float
    friction_da_n: float
    valley_angle_rad: float
    mountain_angle_rad: float
    horizontal_component: float


@dataclass(slots=True)
class _Extrema:
    maxps: list[list[float]]
    minps: list[list[float]]
    fmmax: list[list[float]]
    fmmin: list[list[float]]
    maxav: list[list[float]]
    minav: list[list[float]]
    maxam: list[list[float]]
    minam: list[list[float]]
    maxtv: list[list[float]]
    mintv: list[list[float]]
    maxtm: list[list[float]]
    mintm: list[list[float]]
    maxdev: list[list[float]]
    mindev: list[list[float]]
    maxattr: list[list[float]]
    minattr: list[list[float]]
    maxtp: list[list[float]]
    mintp: list[list[float]]
    maxsq: list[float]
    max_slip: float = 0.0
    ascent_tension_max: float = 0.0
    ascent_tension_min: float = 99999.0
    descent_tension_max: float = 0.0
    descent_tension_min: float = 99999.0
    stroke_max: float = -999999.0
    stroke_min: float = 999999.0
    motive_max: float = -999999.0
    motive_min: float = 999999.0
    ground_length_last: float = 0.0


@dataclass(slots=True)
class _StepState:
    tcv: list[list[float]]
    tcm: list[list[float]]
    tp: list[list[float]]
    dp: list[list[float]]
    pp: list[list[float]]
    friction: list[list[float]]
    av: list[list[float]]
    am: list[list[float]]
    hc: list[list[float]]
    fm: list[list[float]]
    sag_development: list[list[float]]
    elastic_extension: list[list[float]]


@dataclass(slots=True)
class _VehiclesInSpans:
    counts: list[list[int]]
    bo: list[list[list[float]]]


def run_iterative_solver(mode: SolverMode, line_input: LineInput) -> IterativeSolverResult:
    if mode is SolverMode.HYDRAULIC_PRECHECK:
        return _run_hydraulic_precheck(line_input)
    if mode is SolverMode.HYDRAULIC_FINAL:
        return _run_hydraulic_final(line_input)
    if mode not in {SolverMode.NORMAL, SolverMode.VARIABLE_TENSION, SolverMode.ANCHORED}:
        raise NotImplementedError(f"{mode.value} iterative solver parity is not implemented yet.")
    active_cases = []
    for hypothesis_index, row in enumerate(line_input.verification.matrix, start=1):
        for condition_index, state in enumerate(row, start=1):
            if state is MatrixState.NORMAL:
                active_cases.append((hypothesis_index, condition_index, state))
    anchored_reference_ground_length = None
    if mode is SolverMode.ANCHORED:
        reference = _calculate_direction(line_input, 5, 1, "forward", 0.9)
        anchored_reference_ground_length = reference.extrema.ground_length_last
    if mode is SolverMode.VARIABLE_TENSION:
        families = (
            (ResultFamily.PLUS_TEN, 1.0 + line_input.verification.parameters.tightener_precision_percent / 100.0),
            (ResultFamily.MINUS_TEN, 1.0 - line_input.verification.parameters.tightener_precision_percent / 100.0),
        )
    elif mode is SolverMode.ANCHORED:
        families = ((ResultFamily.NORMAL, 1.0),)
    else:
        families = ((ResultFamily.NORMAL, 1.0),)
    case_results: list[IterativeCaseResult] = []
    max_min_results: dict[ResultFamily, MaxMinCase] = {}
    for family, tension_multiplier in families:
        family_cases = [
            _run_case(
                line_input,
                hypothesis,
                condition,
                state,
                family,
                tension_multiplier,
                anchored_reference_ground_length,
            )
            for hypothesis, condition, state in active_cases
        ]
        case_results.extend(family_cases)
        if family_cases:
            max_min_results[family] = _combined_max_min_case(family_cases)
    return IterativeSolverResult(mode=mode, cases=case_results, max_min_results=max_min_results, parity_complete=False)


def _run_hydraulic_precheck(line_input: LineInput) -> IterativeSolverResult:
    reference_values = [[0.0 for _column in range(3)] for _row in range(7)]
    for row_index, (hypothesis_index, condition_index) in enumerate(((5, 1), (2, 1), (1, 2), (4, 2))):
        for column_index, tension_multiplier in enumerate((1.0, 1.1, 0.9)):
            result = _calculate_direction(
                line_input,
                hypothesis_index,
                condition_index,
                "forward",
                tension_multiplier,
            )
            reference_values[row_index][column_index] = _hydraulic_stroke_reference(result)
    for column_index in range(3):
        reference_values[4][column_index] = reference_values[0][column_index] - reference_values[1][column_index]
        reference_values[5][column_index] = reference_values[1][column_index] - reference_values[2][column_index]
        reference_values[6][column_index] = reference_values[2][column_index] - reference_values[3][column_index]
    return IterativeSolverResult(
        mode=SolverMode.HYDRAULIC_PRECHECK,
        cases=[],
        max_min_results={},
        hydraulic_reference_values=reference_values,
        parity_complete=False,
    )


def _run_hydraulic_final(line_input: LineInput) -> IterativeSolverResult:
    # The final VBA path iterates on dialog-provided cylinder stroke values. Until those values are carried in
    # LineInput, produce the hydraulic STORE family with the nominal tension path so result objects are available.
    active_cases = []
    for hypothesis_index, row in enumerate(line_input.verification.matrix, start=1):
        for condition_index, state in enumerate(row, start=1):
            if state is MatrixState.NORMAL:
                active_cases.append((hypothesis_index, condition_index, state))
    cases = [
        _run_case(line_input, hypothesis, condition, state, ResultFamily.HYDRAULIC, 1.0)
        for hypothesis, condition, state in active_cases
    ]
    max_min_results = {ResultFamily.HYDRAULIC: _combined_max_min_case(cases)} if cases else {}
    return IterativeSolverResult(
        mode=SolverMode.HYDRAULIC_FINAL,
        cases=cases,
        max_min_results=max_min_results,
        hydraulic_reference_values=_run_hydraulic_precheck(line_input).hydraulic_reference_values,
        parity_complete=False,
    )


def _run_case(
    line_input: LineInput,
    hypothesis_index: int,
    condition_index: int,
    state: MatrixState,
    family: ResultFamily,
    tension_multiplier: float,
    anchored_reference_ground_length: float | None = None,
) -> IterativeCaseResult:
    reference = ResultCaseRef(
        hypothesis_index=hypothesis_index,
        condition_index=condition_index,
        state=state,
        label=(
            f"{line_input.verification.hypotheses[hypothesis_index - 1].index}. "
            f"{line_input.verification.hypotheses[hypothesis_index - 1].label} - "
            f"{line_input.verification.conditions[condition_index - 1].label}"
        ),
    )
    forward = _calculate_direction(
        line_input,
        hypothesis_index,
        condition_index,
        "forward",
        tension_multiplier,
        anchored_reference_ground_length,
    )
    reverse = _calculate_direction(
        line_input,
        hypothesis_index,
        condition_index,
        "reverse",
        tension_multiplier,
        anchored_reference_ground_length,
    )
    power_trace = PowerTraceCase(
        family=family,
        reference=reference,
        forward=forward.trace,
        reverse=reverse.trace,
        source="Native iterative solver",
    )
    power_summary = _power_summary(line_input, family, reference, forward, reverse, condition_index)
    return IterativeCaseResult(
        family=family,
        reference=reference,
        power_trace=power_trace,
        forward_line_results=_line_results(line_input, family, reference, forward.extrema, flag_average=0),
        power_summary=power_summary,
    )


@dataclass(slots=True)
class _DirectionResult:
    trace: list[PowerTracePoint]
    extrema: _Extrema


def _calculate_direction(
    line_input: LineInput,
    hypothesis_index: int,
    condition_index: int,
    direction: str,
    tension_multiplier: float,
    anchored_reference_ground_length: float | None = None,
) -> _DirectionResult:
    run_parameters = derive_run_parameters(line_input)
    load_preparation = prepare_loads(line_input, hypothesis_index, condition_index)
    motion = _motion_parameters(line_input, condition_index)
    extrema = _new_extrema(line_input)
    trace: list[PowerTracePoint] = []
    offset = 0.0
    step_index = 0
    while offset <= run_parameters.car_spacing_m + 1e-9:
        positions = position_vehicles(
            offset_m=offset,
            ascent_length_m=line_input.ascent.developed_length_m,
            descent_length_m=line_input.descent.developed_length_m,
            drive_pulley_diameter_mm=line_input.general.drive_pulley_diameter_mm,
            car_spacing_m=run_parameters.vehicle_spacing_m,
            vehicles_per_group=run_parameters.vehicles_per_group,
            intra_group_spacing_m=run_parameters.car_spacing_m,
        )
        vehicles = _vehicles_in_spans(line_input, positions.ascent, positions.descent)
        step = _solve_step(
            line_input,
            load_preparation.ascent_loads_da_n,
            load_preparation.descent_loads_da_n,
            vehicles,
            motion,
            direction,
            tension_multiplier,
            reference_index=(hypothesis_index - 1) * 6 + condition_index,
            anchored_reference_ground_length=anchored_reference_ground_length,
        )
        _memorize(
            line_input,
            step,
            vehicles,
            load_preparation.ascent_loads_da_n,
            load_preparation.descent_loads_da_n,
            reference_index=(hypothesis_index - 1) * 6 + condition_index,
            extrema=extrema,
        )
        ascent_tension, descent_tension, motive = _trace_values(line_input, step, direction)
        trace.append(
            PowerTracePoint(
                offset_m=offset,
                ascent_tension_da_n=ascent_tension,
                descent_tension_da_n=descent_tension,
                motive_force_da_n=motive,
            )
        )
        _update_trace_extrema(extrema, ascent_tension, descent_tension, motive)
        step_index += 1
        offset = step_index * run_parameters.step_m
    return _DirectionResult(trace=trace, extrema=extrema)


def _solve_step(
    line_input: LineInput,
    ascent_loads: list[float],
    descent_loads: list[float],
    vehicles: _VehiclesInSpans,
    motion: _MotionParameters,
    direction: str,
    tension_multiplier: float,
    reference_index: int,
    anchored_reference_ground_length: float | None = None,
) -> _StepState:
    t_start = line_input.general.rated_tension_da_n * tension_multiplier
    if anchored_reference_ground_length is None:
        return _solve_tension_balance(line_input, ascent_loads, descent_loads, vehicles, motion, direction, t_start)
    previous_delta = 0.0
    tension_step = t_start / 10.0
    for iteration in range(50):
        step = _solve_tension_balance(line_input, ascent_loads, descent_loads, vehicles, motion, direction, t_start)
        _update_span_developments(line_input, step, vehicles, ascent_loads, descent_loads, reference_index)
        ground_length = _ground_rope_length(line_input, step, line_input.verification.parameters.local_temperature_c)
        delta = anchored_reference_ground_length - ground_length
        if abs(delta) <= 0.001:
            return step
        if iteration == 0:
            adjustment = tension_step
        else:
            denominator = previous_delta - delta
            adjustment = tension_step / 2.0 if abs(denominator) < 1e-12 else tension_step * delta / denominator
        t_start += adjustment
        tension_step = adjustment
        previous_delta = delta
    return step


def _solve_tension_balance(
    line_input: LineInput,
    ascent_loads: list[float],
    descent_loads: list[float],
    vehicles: _VehiclesInSpans,
    motion: _MotionParameters,
    direction: str,
    t_start: float,
) -> _StepState:
    step = _new_step_state(line_input)
    ns = len(line_input.ascent.supports)
    nd = len(line_input.descent.supports)
    ksd_ascent, ksd_descent = (1.0, -1.0) if direction == "forward" else (-1.0, 1.0)
    start_ascent = t_start / 2.0
    start_descent = t_start / 2.0
    for _iteration in range(101):
        step.tcm[1][0] = start_ascent
        step.tcm[2][0] = start_descent
        for branch_index, ksd in ((1, ksd_ascent), (2, ksd_descent)):
            branch = line_input.ascent if branch_index == 1 else line_input.descent
            loads = ascent_loads if branch_index == 1 else descent_loads
            for span_index in range(1, len(branch.supports)):
                calc = _campata(line_input, branch_index, span_index, branch, loads, vehicles, step, motion, ksd)
                step.tcv[branch_index][span_index] = calc.valley_tension
                step.tcm[branch_index][span_index] = calc.mountain_tension
                step.tp[branch_index][span_index] = calc.support_tension
                step.dp[branch_index][span_index] = calc.deviation_rad
                step.pp[branch_index][span_index] = calc.pressure_da_n
                step.friction[branch_index][span_index] = calc.friction_da_n
                step.av[branch_index][span_index] = calc.valley_angle_rad
                step.am[branch_index][span_index] = calc.mountain_angle_rad
                step.hc[branch_index][span_index] = calc.horizontal_component
        if line_input.general.drive_station == "VALLE" and line_input.general.tension_station == "MONTE":
            ascent_delta = step.tcm[1][ns - 1] - t_start / 2.0
            descent_delta = step.tcm[2][nd - 1] - t_start / 2.0
            if abs(step.tcm[1][ns - 1] + step.tcm[2][nd - 1] - t_start) <= 0.1:
                return step
            start_ascent -= ascent_delta
            start_descent -= descent_delta
            continue
        if line_input.general.drive_station == "VALLE" and line_input.general.tension_station == "VALLE":
            delta = step.tcm[1][ns - 1] - step.tcm[2][nd - 1]
            if abs(delta) <= 0.1:
                return step
            start_ascent -= delta / 2.0
            start_descent += delta / 2.0
            continue
        return step
    return step


def _campata(
    line_input: LineInput,
    branch_index: int,
    span_index: int,
    branch: LineBranch,
    loads: list[float],
    vehicles: _VehiclesInSpans,
    step: _StepState,
    motion: _MotionParameters,
    ksd: float,
) -> _SpanCalculation:
    span = branch.spans[span_index - 1]
    horizontal = float(span.horizontal_distance_m)
    vertical = float(span.height_difference_m)
    alpha = atan(vertical / horizontal)
    chord = sqrt(horizontal**2 + vertical**2)
    vehicle_load = loads[span_index - 1] if span_index - 1 < len(loads) else 0.0
    vehicle_count = vehicles.counts[branch_index][span_index]
    tk = step.tcm[branch_index][span_index - 1]
    a_value = line_input.general.rope_unit_weight_da_n_m * chord * horizontal / 2.0
    b_value = sum(vehicle_load * bo for bo in vehicles.bo[branch_index][span_index])
    attr = _friction_value(line_input, branch_index, span_index, condition_index_from_motion(motion))
    absolute_friction = 0.0
    support_tension = tk
    pressure = 0.0
    poriz = 0.0
    deviation = 0.0
    beta = 0.0
    tv = tk
    hk = tk
    for _iteration in range(100):
        c_value = tk * chord
        x_value = (a_value + b_value) / c_value
        epsilon = atan(x_value / sqrt(1.0 - x_value**2))
        beta = epsilon - alpha
        hk, nv, tv = _tenscampata(line_input, tk, beta, chord, vehicle_load, vehicle_count, motion.acceleration, ksd)
        rollers = _roller_count(line_input, branch_index, span_index)
        support_code = branch.supports[span_index - 1].code
        roller_weight = line_input.general.compression_roller_weight if "R" in support_code else line_input.general.support_roller_weight
        roller_mass = rollers * roller_weight / (VBA_G * 10.0)
        support_tension = step.tcm[branch_index][span_index - 1] + roller_mass / 2.0 * motion.acceleration * ksd
        horizontal_angle = _horizontal_deviation_angle(support_code)
        if _friction_is_absolute(line_input, branch_index, span_index):
            support_tension += attr / 2.0 * ksd
        else:
            support_tension += (abs(pressure) + abs(poriz)) * attr / 2.0 * ksd
        deviation = step.am[branch_index][span_index - 1] + beta
        pressure = 2.0 * support_tension * sin(deviation / 2.0)
        poriz = 2.0 * support_tension * sin(horizontal_angle / 2.0)
        if span_index <= 1:
            break
        if _friction_is_absolute(line_input, branch_index, span_index):
            tk = support_tension + attr / 2.0 * ksd + roller_mass / 2.0 * motion.acceleration * ksd
            absolute_friction = attr
            hk, nv, tv = _tenscampata(line_input, tk, beta, chord, vehicle_load, vehicle_count, motion.acceleration, ksd)
            break
        next_friction = (abs(pressure) + abs(poriz)) * attr
        tk = support_tension + next_friction / 2.0 * ksd + roller_mass / 2.0 * motion.acceleration * ksd
        if abs(absolute_friction - abs(pressure * attr)) < 0.1:
            absolute_friction = next_friction
            break
        absolute_friction = next_friction
    hk, nv, tv = _tenscampata(line_input, tk, beta, chord, vehicle_load, vehicle_count, motion.acceleration, ksd)
    mountain_tension = tv
    if span_index == 1 and condition_index_from_motion(motion) != 1:
        if line_input.general.acceleration_beam_force > 0:
            mountain_tension += line_input.general.acceleration_beam_force / 10.0 * ksd
        station = line_input.general.station_deviation_angle
        if 0 < station <= 0.5:
            mountain_tension += mountain_tension * 2.0 * sin(station / 2.0) * line_input.general.friction_running_percent / 100.0 * ksd
    return _SpanCalculation(
        valley_tension=tk,
        mountain_tension=mountain_tension,
        support_tension=support_tension,
        deviation_rad=deviation,
        pressure_da_n=pressure,
        friction_da_n=absolute_friction,
        valley_angle_rad=beta,
        mountain_angle_rad=atan(nv / hk),
        horizontal_component=hk,
    )


def _tenscampata(
    line_input: LineInput,
    tk: float,
    beta: float,
    chord: float,
    vehicle_load: float,
    vehicle_count: int,
    acceleration: float,
    ksd: float,
) -> tuple[float, float, float]:
    hk = tk * cos(beta)
    nk = tk * sin(beta)
    vertical = line_input.general.rope_unit_weight_da_n_m * chord + vehicle_load * vehicle_count - nk
    tv = sqrt(vertical**2 + hk**2)
    span_mass = (line_input.general.rope_unit_weight_da_n_m * chord + vehicle_load * vehicle_count) / VBA_G
    tv += span_mass * acceleration * ksd / 10.0
    return hk, vertical, tv


def _memorize(
    line_input: LineInput,
    step: _StepState,
    vehicles: _VehiclesInSpans,
    ascent_loads: list[float],
    descent_loads: list[float],
    reference_index: int,
    extrema: _Extrema,
) -> None:
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            _maxmin(extrema.maxtp, extrema.mintp, branch_index, span_index, step.tp[branch_index][span_index])
            _maxmin(extrema.maxps, extrema.minps, branch_index, span_index, step.pp[branch_index][span_index])
            _maxmin(extrema.maxdev, extrema.mindev, branch_index, span_index, step.dp[branch_index][span_index])
            sq = abs(abs(step.pp[1][span_index]) - abs(step.pp[2][span_index]))
            if sq > extrema.maxsq[span_index]:
                extrema.maxsq[span_index] = sq
            friction = abs(step.friction[branch_index][span_index])
            _maxmin(extrema.maxattr, extrema.minattr, branch_index, span_index, friction)
            _maxmin(extrema.maxtv, extrema.mintv, branch_index, span_index, step.tcv[branch_index][span_index])
            _maxmin(extrema.maxtm, extrema.mintm, branch_index, span_index, step.tcm[branch_index][span_index])
            _maxmin(extrema.maxav, extrema.minav, branch_index, span_index, step.av[branch_index][span_index])
            _maxmin(extrema.maxam, extrema.minam, branch_index, span_index, step.am[branch_index][span_index])
    _update_span_developments(line_input, step, vehicles, ascent_loads, descent_loads, reference_index)
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            _maxmin(extrema.fmmax, extrema.fmmin, branch_index, span_index, step.fm[branch_index][span_index])
    _update_stroke(line_input, step, extrema)
    extrema.ground_length_last = _ground_rope_length(line_input, step, line_input.verification.parameters.local_temperature_c)


def _update_span_developments(
    line_input: LineInput,
    step: _StepState,
    vehicles: _VehiclesInSpans,
    ascent_loads: list[float],
    descent_loads: list[float],
    reference_index: int,
) -> None:
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            horizontal = float(branch.spans[span_index - 1].horizontal_distance_m)
            vertical = float(branch.spans[span_index - 1].height_difference_m)
            chord = sqrt(horizontal**2 + vertical**2)
            alpha = atan(vertical / horizontal)
            mean_tension = (step.tcv[branch_index][span_index] + step.tcm[branch_index][span_index]) / 2.0
            fm = line_input.general.rope_unit_weight_da_n_m * chord**2 / (8.0 * mean_tension)
            sag_development = 8.0 / 3.0 * fm**2 * cos(alpha) ** 2 / chord
            if vehicles.counts[branch_index][span_index] > 0 and reference_index > 1:
                loads = ascent_loads if branch_index == 1 else descent_loads
                fm, sag_development = _loaded_sag(
                    line_input,
                    branch_index,
                    span_index,
                    loads[span_index - 1] if span_index - 1 < len(loads) else 0.0,
                    vehicles,
                    step,
                    fm,
                    sag_development,
                )
            step.fm[branch_index][span_index] = fm
            step.sag_development[branch_index][span_index] = sag_development
            step.elastic_extension[branch_index][span_index] = mean_tension * chord / (
                line_input.general.rope_modulus_n_mm2 / 10.0 * line_input.general.rope_metallic_area_mm2
            )


def _ground_rope_length(line_input: LineInput, step: _StepState, delta_temperature: float) -> float:
    total_development = 0.0
    total_elastic = 0.0
    total_thermal = 0.0
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            horizontal = float(branch.spans[span_index - 1].horizontal_distance_m)
            vertical = float(branch.spans[span_index - 1].height_difference_m)
            chord = sqrt(horizontal**2 + vertical**2)
            span_development = chord + step.sag_development[branch_index][span_index]
            mean_tension = (step.tcv[branch_index][span_index] + step.tcm[branch_index][span_index]) / 2.0
            total_development += span_development
            total_elastic += span_development * 10.0 * mean_tension / line_input.general.rope_modulus_n_mm2 / line_input.general.rope_metallic_area_mm2
            total_thermal += span_development * 0.000012 * delta_temperature
    return total_development - total_elastic - total_thermal


def _loaded_sag(
    line_input: LineInput,
    branch_index: int,
    span_index: int,
    load: float,
    vehicles: _VehiclesInSpans,
    step: _StepState,
    bare_fm: float,
    bare_development: float,
) -> tuple[float, float]:
    branch = line_input.ascent if branch_index == 1 else line_input.descent
    span = branch.spans[span_index - 1]
    horizontal = float(span.horizontal_distance_m)
    vertical = float(span.height_difference_m)
    alpha = atan(vertical / horizontal)
    chord = sqrt(horizontal**2 + vertical**2)
    h_component = step.hc[branch_index][span_index]
    bo_values = vehicles.bo[branch_index][span_index]
    fv = [load * (horizontal - bo) * bo / (horizontal * h_component) for bo in bo_values]
    fvc = []
    for bo, value in zip(bo_values, fv):
        if bo <= horizontal / 2.0:
            fvc.append(value * horizontal / (2.0 * (horizontal - bo)))
        else:
            fvc.append(value * horizontal / (2.0 * bo))
    fm = bare_fm + sum(fvc)
    bare_ff = _bare_sag_at_loads(horizontal, alpha, bare_fm, bo_values)
    fvt: list[float] = []
    for index, bo in enumerate(bo_values):
        value = fv[index]
        for previous in range(index):
            value += fv[previous] / bo_values[previous] * bo
        for following in range(index + 1, len(bo_values)):
            value += fv[following] / (horizontal - bo_values[following]) * (horizontal - bo)
        fvt.append(value)
    total = 0.0
    skj = 0.0
    for index in range(len(bo_values) + 1):
        if index == 0:
            li = horizontal - bo_values[index]
            di = li * tan(alpha) - fvt[index] - bare_ff[index]
        elif index < len(bo_values):
            li = bo_values[index - 1] - bo_values[index]
            di = ((horizontal - bo_values[index]) * tan(alpha) - fvt[index] - bare_ff[index]) - (
                (horizontal - bo_values[index - 1]) * tan(alpha) - fvt[index - 1] - bare_ff[index - 1]
            )
        else:
            li = bo_values[index - 1]
            di = li * tan(alpha) + fvt[index - 1] + bare_ff[index - 1]
        segment_alpha = atan(di / li)
        segment_chord = sqrt(li**2 + di**2)
        total += segment_chord
        f_segment = line_input.general.rope_unit_weight_da_n_m * li**2 / (8.0 * h_component * cos(segment_alpha))
        skj += 8.0 / 3.0 * f_segment**2 * cos(segment_alpha) ** 2 / segment_chord
    development = total + skj - chord
    return fm, max(bare_development, development)


def _bare_sag_at_loads(horizontal: float, alpha: float, sag: float, bo_values: list[float]) -> list[float]:
    values: list[float] = []
    for bo in bo_values:
        x_value = horizontal - bo
        y_value = x_value * tan(alpha) - 4.0 * sag / horizontal**2 * x_value * (horizontal - x_value)
        values.append(x_value * tan(alpha) - y_value)
    return values

def _update_stroke(line_input: LineInput, step: _StepState, extrema: _Extrema) -> None:
    sag_sum = 0.0
    elastic_sum = 0.0
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            sag_sum += step.sag_development[branch_index][span_index]
            elastic_sum += step.elastic_extension[branch_index][span_index]
    stroke = elastic_sum - sag_sum
    if stroke > extrema.stroke_max:
        extrema.stroke_max = stroke
    if stroke < extrema.stroke_min:
        extrema.stroke_min = stroke
    extrema.ground_length_last = _ground_rope_length(line_input, step, 0.0)


def _ground_rope_length(line_input: LineInput, step: _StepState, delta_temperature_c: float) -> float:
    total_development = 0.0
    elastic_extension = 0.0
    thermal_extension = 0.0
    for branch_index, branch in ((1, line_input.ascent), (2, line_input.descent)):
        for span_index in range(1, len(branch.supports)):
            horizontal = float(branch.spans[span_index - 1].horizontal_distance_m)
            vertical = float(branch.spans[span_index - 1].height_difference_m)
            chord = sqrt(horizontal**2 + vertical**2)
            span_development = chord + step.sag_development[branch_index][span_index]
            mean_tension = (step.tcv[branch_index][span_index] + step.tcm[branch_index][span_index]) / 2.0
            total_development += span_development
            elastic_extension += (
                span_development
                * 10.0
                * mean_tension
                / line_input.general.rope_modulus_n_mm2
                / line_input.general.rope_metallic_area_mm2
            )
            thermal_extension += 0.000012 * span_development * delta_temperature_c
    return total_development - elastic_extension - thermal_extension


def _trace_values(line_input: LineInput, step: _StepState, direction: str) -> tuple[float, float, float]:
    if line_input.general.drive_station == "MONTE":
        ascent = step.tcm[1][len(line_input.ascent.supports) - 1]
        descent = step.tcm[2][len(line_input.descent.supports) - 1]
        motive = ascent - descent if direction == "forward" else descent - ascent
        return ascent, descent, motive
    ascent = step.tcv[1][1]
    descent = step.tcv[2][1]
    motive = descent - ascent if direction == "forward" else ascent - descent
    return ascent, descent, motive


def _update_trace_extrema(extrema: _Extrema, ascent: float, descent: float, motive: float) -> None:
    extrema.motive_max = max(extrema.motive_max, motive)
    extrema.motive_min = min(extrema.motive_min, motive)
    if descent:
        extrema.max_slip = max(extrema.max_slip, ascent / descent if ascent > descent else descent / ascent)
    extrema.ascent_tension_max = max(extrema.ascent_tension_max, ascent)
    extrema.descent_tension_max = max(extrema.descent_tension_max, descent)
    extrema.ascent_tension_min = min(extrema.ascent_tension_min, ascent)
    extrema.descent_tension_min = min(extrema.descent_tension_min, descent)


def _hydraulic_stroke_reference(result: _DirectionResult) -> float:
    return (result.extrema.stroke_max + result.extrema.stroke_min) / 4.0


def _line_results(
    line_input: LineInput,
    family: ResultFamily,
    reference: ResultCaseRef,
    extrema: _Extrema,
    flag_average: int,
) -> LineResultCase:
    return LineResultCase(
        family=family,
        reference=reference,
        source="Native iterative solver",
        flag_average=flag_average,
        ascent=_line_branch_results(line_input.ascent, 1, extrema, flag_average, "Ascent Branch Results"),
        descent=_line_branch_results(line_input.descent, 2, extrema, flag_average, "Descent Branch Results"),
    )


def _power_summary(
    line_input: LineInput,
    family: ResultFamily,
    reference: ResultCaseRef,
    forward: _DirectionResult,
    reverse: _DirectionResult,
    condition_index: int,
) -> PowerSummaryCase:
    rows = [
        _power_summary_row(line_input, reference, forward, condition_index, "Forward"),
        _power_summary_row(line_input, reference, reverse, condition_index, "Reverse"),
    ]
    return PowerSummaryCase(family=family, rows=rows, source="Native iterative solver")


def _power_summary_row(
    line_input: LineInput,
    reference: ResultCaseRef,
    direction: _DirectionResult,
    condition_index: int,
    direction_label: str,
) -> PowerSummaryRow:
    motion = _motion_parameters(line_input, condition_index)
    motive_values = [point.motive_force_da_n for point in direction.trace]
    mean_force = sum(motive_values) / len(motive_values) if motive_values else 0.0
    drive_inertia = line_input.general.drive_equivalent_weight / 10.0 * motion.acceleration
    motor_stress = mean_force + drive_inertia
    mean_efficiency = line_input.general.drive_efficiency if motor_stress > 0 else 1.0 / line_input.general.drive_efficiency
    mean_power = motor_stress * line_input.general.speed_m_s / mean_efficiency / 100.0
    max_motive = max(motive_values, key=abs) if motive_values else 0.0
    max_stress = max_motive + drive_inertia
    return PowerSummaryRow(
        direction=direction_label,
        case_index=reference.store_index,
        load_hypothesis=line_input.verification.hypotheses[reference.hypothesis_index - 1].label,
        plant_condition=line_input.verification.conditions[reference.condition_index - 1].label,
        mean_force_da_n=mean_force,
        max_force_da_n=max_stress,
        motor_inertia_da_n=drive_inertia,
        motor_stress_da_n=motor_stress,
        efficiency=mean_efficiency,
        power_kw=mean_power,
        slip=direction.extrema.max_slip,
        length_m=direction.extrema.stroke_max / 2.0,
        total_tension_da_n=(
            (direction.extrema.ascent_tension_max + direction.extrema.ascent_tension_min) / 2.0
            + (direction.extrema.descent_tension_max + direction.extrema.descent_tension_min) / 2.0
        ),
    )


def _max_min_case(line_result: LineResultCase) -> MaxMinCase:
    return MaxMinCase(
        family=line_result.family,
        ascent_rows=_max_min_rows(line_result.ascent),
        descent_rows=_max_min_rows(line_result.descent),
        source="Native iterative solver",
    )


def _combined_max_min_case(cases: list[IterativeCaseResult]) -> MaxMinCase:
    family = cases[0].family
    return MaxMinCase(
        family=family,
        ascent_rows=_combined_max_min_rows([case.forward_line_results.ascent for case in cases]),
        descent_rows=_combined_max_min_rows([case.forward_line_results.descent for case in cases]),
        source="Native iterative solver",
    )


def _combined_max_min_rows(branches: list[LineResultBranch]) -> list[MaxMinRow]:
    rows: list[MaxMinRow] = []
    if not branches:
        return rows
    row_count = len(branches[0].rows)
    index = 0
    while index < row_count:
        max_span = _select_numeric_row((branch.rows[index] for branch in branches), "valley_tension_da_n", max)
        min_span = _select_numeric_row((branch.rows[index + 1] for branch in branches), "valley_tension_da_n", min)
        rows.append(_max_min_span_row(max_span, "max"))
        rows.append(_max_min_span_row(min_span, "min"))
        index += 2
        if index >= row_count or not branches[0].rows[index].support_code:
            continue
        max_support = _select_numeric_row((branch.rows[index] for branch in branches), "support_tension_da_n", max)
        min_support = _select_numeric_row((branch.rows[index + 1] for branch in branches), "support_tension_da_n", min)
        rows.append(_max_min_support_row(max_support, "max"))
        rows.append(_max_min_support_row(min_support, "min"))
        index += 2
    return rows


def _select_numeric_row(rows, field_name: str, selector) -> LineResultRow:
    numeric_rows = [row for row in rows if isinstance(getattr(row, field_name), float)]
    if not numeric_rows:
        return LineResultRow()
    selected_value = selector(getattr(row, field_name) for row in numeric_rows)
    for row in numeric_rows:
        if getattr(row, field_name) == selected_value:
            return row
    return numeric_rows[0]


def _max_min_span_row(row: LineResultRow, state: str) -> MaxMinRow:
    return MaxMinRow(
        span_start=row.span_start,
        span_end=row.span_end,
        span_state=state,
        tension_da_n=row.valley_tension_da_n,
        sag_m=row.sag_m,
        valley_angle_deg=row.valley_angle_deg,
        mountain_angle_deg=row.mountain_angle_deg,
    )


def _max_min_rows(branch: LineResultBranch) -> list[MaxMinRow]:
    rows: list[MaxMinRow] = []
    index = 0
    while index < len(branch.rows):
        span_max = branch.rows[index]
        span_min = branch.rows[index + 1] if index + 1 < len(branch.rows) else LineResultRow()
        rows.append(
            MaxMinRow(
                span_start=span_max.span_start,
                span_end=span_max.span_end,
                span_state="max",
                tension_da_n=span_max.valley_tension_da_n,
                sag_m=span_max.sag_m,
                valley_angle_deg=span_max.valley_angle_deg,
                mountain_angle_deg=span_max.mountain_angle_deg,
            )
        )
        rows.append(
            MaxMinRow(
                span_state="min",
                tension_da_n=span_min.valley_tension_da_n,
                sag_m=span_min.sag_m,
                valley_angle_deg=span_min.valley_angle_deg,
                mountain_angle_deg=span_min.mountain_angle_deg,
            )
        )
        index += 2
        if index >= len(branch.rows) or not branch.rows[index].support_code:
            continue
        support_max = branch.rows[index]
        support_min = branch.rows[index + 1] if index + 1 < len(branch.rows) else LineResultRow()
        rows.append(_max_min_support_row(support_max, "max"))
        rows.append(_max_min_support_row(support_min, "min"))
        index += 2
    return rows


def _max_min_support_row(row: LineResultRow, state: str) -> MaxMinRow:
    return MaxMinRow(
        support_code=row.support_code,
        support_state=state,
        support_tension_da_n=row.support_tension_da_n,
        deviation_deg=row.total_deviation_deg,
        pressure_da_n=row.pressure_da_n,
        friction_da_n=row.friction_da_n,
        roller_count=row.roller_count,
        unit_deviation_deg=row.unit_deviation_deg,
        unit_pressure_da_n=row.unit_pressure_da_n,
    )


def _line_branch_results(branch: LineBranch, branch_index: int, extrema: _Extrema, flag_average: int, name: str) -> LineResultBranch:
    rows: list[LineResultRow] = []
    support_count = len(branch.supports)
    for j in range(2, support_count + 1):
        rows.append(
            LineResultRow(
                span_start=branch.supports[j - 2].code,
                span_end=branch.supports[j - 1].code,
                valley_tension_da_n=extrema.maxtv[branch_index][j - 1],
                mountain_tension_da_n=extrema.maxtm[branch_index][j - 1],
                sag_m=extrema.fmmax[branch_index][j - 1],
                valley_angle_deg=extrema.maxav[branch_index][j - 1] * 180.0 / VBA_PI,
                mountain_angle_deg=extrema.maxam[branch_index][j - 1] * 180.0 / VBA_PI,
            )
        )
        if flag_average == 0:
            rows.append(
                LineResultRow(
                    valley_tension_da_n=extrema.mintv[branch_index][j - 1],
                    mountain_tension_da_n=extrema.mintm[branch_index][j - 1],
                    sag_m=extrema.fmmin[branch_index][j - 1],
                    valley_angle_deg=extrema.minav[branch_index][j - 1] * 180.0 / VBA_PI,
                    mountain_angle_deg=extrema.minam[branch_index][j - 1] * 180.0 / VBA_PI,
                )
            )
        if j < support_count:
            support_code = branch.supports[j - 1].code
            rollers = float(branch.supports[j - 1].roller_quantity or 0.0)
            rows.append(
                LineResultRow(
                    support_code=support_code,
                    support_tension_da_n=extrema.maxtp[branch_index][j],
                    total_deviation_deg=extrema.maxdev[branch_index][j] * 180.0 / VBA_PI,
                    pressure_da_n=extrema.maxps[branch_index][j],
                    friction_da_n=extrema.maxattr[branch_index][j],
                    roller_count=rollers,
                    unit_deviation_deg=extrema.maxdev[branch_index][j] * 180.0 / VBA_PI / rollers if rollers else "",
                    unit_pressure_da_n=extrema.maxps[branch_index][j] / rollers if rollers else "",
                )
            )
            if flag_average == 0:
                rows.append(
                    LineResultRow(
                        support_tension_da_n=extrema.mintp[branch_index][j],
                        total_deviation_deg=extrema.mindev[branch_index][j] * 180.0 / VBA_PI,
                        pressure_da_n=extrema.minps[branch_index][j],
                        friction_da_n=extrema.minattr[branch_index][j],
                        roller_count=rollers,
                        unit_deviation_deg=extrema.mindev[branch_index][j] * 180.0 / VBA_PI / rollers if rollers else "",
                        unit_pressure_da_n=extrema.minps[branch_index][j] / rollers if rollers else "",
                    )
                )
    return LineResultBranch(name=name, rows=rows)


def _vehicles_in_spans(line_input: LineInput, ascent_positions: list[float], descent_positions: list[float]) -> _VehiclesInSpans:
    max_supports = max(len(line_input.ascent.supports), len(line_input.descent.supports)) + 1
    counts = [[0 for _ in range(max_supports)] for _branch in range(3)]
    bo = [[[] for _ in range(max_supports)] for _branch in range(3)]
    for branch_index, branch, positions in (
        (1, line_input.ascent, ascent_positions),
        (2, line_input.descent, descent_positions),
    ):
        for span_index in range(1, len(branch.supports)):
            alpha = atan(float(branch.spans[span_index - 1].height_difference_m) / float(branch.spans[span_index - 1].horizontal_distance_m))
            start = branch.support_progressive_m[span_index - 1]
            end = branch.support_progressive_m[span_index]
            for position in positions:
                if position > start and position < end:
                    counts[branch_index][span_index] += 1
                    bo[branch_index][span_index].append((end - position) * cos(alpha))
    return _VehiclesInSpans(counts=counts, bo=bo)


def _motion_parameters(line_input: LineInput, condition_index: int) -> _MotionParameters:
    acceleration_by_condition = {
        1: 0.0,
        2: 0.0,
        3: line_input.general.acceleration_m_s2,
        4: -line_input.general.deceleration_1_m_s2,
        5: -line_input.general.deceleration_2_m_s2,
        6: -line_input.general.deceleration_3_m_s2,
    }
    friction_by_condition = {
        1: 0.0,
        2: line_input.general.friction_running_percent,
        3: line_input.general.friction_running_percent,
        4: line_input.general.friction_braking_percent,
        5: line_input.general.friction_braking_percent,
        6: line_input.general.friction_braking_percent,
    }
    return _MotionParameters(
        acceleration=acceleration_by_condition[condition_index],
        friction_percent=friction_by_condition[condition_index],
    )


def condition_index_from_motion(motion: _MotionParameters) -> int:
    if motion.acceleration == 0.0 and motion.friction_percent == 0.0:
        return 1
    if motion.acceleration == 0.0:
        return 2
    if motion.acceleration > 0.0:
        return 3
    if motion.friction_percent:
        return 4
    return 2


def _friction_value(line_input: LineInput, branch_index: int, span_index: int, condition_index: int) -> float:
    row_index = span_index - 2
    rows = line_input.ascent_friction if branch_index == 1 else line_input.descent_friction
    if row_index >= 0 and row_index < len(rows) and rows[row_index].friction_type:
        value = rows[row_index].braking_value if condition_index > 3 else rows[row_index].steady_value
        return value * _roller_count(line_input, branch_index, span_index) if rows[row_index].friction_type.upper() == "A" else value / 100.0
    return line_input.general.friction_braking_percent / 100.0 if condition_index > 3 else line_input.general.friction_running_percent / 100.0


def _friction_is_absolute(line_input: LineInput, branch_index: int, span_index: int) -> bool:
    row_index = span_index - 2
    rows = line_input.ascent_friction if branch_index == 1 else line_input.descent_friction
    return row_index >= 0 and row_index < len(rows) and rows[row_index].friction_type.upper() == "A"


def _horizontal_deviation_angle(support_code: str) -> float:
    marker = "/C"
    if marker not in support_code:
        return 0.0
    value = support_code.split(marker, 1)[1]
    try:
        return float(value) * VBA_PI / 180.0
    except ValueError:
        return 0.0


def _roller_count(line_input: LineInput, branch_index: int, support_index: int) -> float:
    branch = line_input.ascent if branch_index == 1 else line_input.descent
    if support_index - 1 >= len(branch.supports):
        return 0.0
    return float(branch.supports[support_index - 1].roller_quantity or 0.0)


def _maxmin(max_array: list[list[float]], min_array: list[list[float]], branch: int, index: int, value: float) -> None:
    if value > max_array[branch][index]:
        max_array[branch][index] = value
    if value < min_array[branch][index]:
        min_array[branch][index] = value


def _new_step_state(line_input: LineInput) -> _StepState:
    size = max(len(line_input.ascent.supports), len(line_input.descent.supports)) + 2
    return _StepState(
        tcv=_zeros(size),
        tcm=_zeros(size),
        tp=_zeros(size),
        dp=_zeros(size),
        pp=_zeros(size),
        friction=_zeros(size),
        av=_zeros(size),
        am=_zeros(size),
        hc=_zeros(size),
        fm=_zeros(size),
        sag_development=_zeros(size),
        elastic_extension=_zeros(size),
    )


def _new_extrema(line_input: LineInput) -> _Extrema:
    size = max(len(line_input.ascent.supports), len(line_input.descent.supports)) + 2
    return _Extrema(
        maxps=_filled(size, -999999.0),
        minps=_filled(size, 999999.0),
        fmmax=_filled(size, -999999.0),
        fmmin=_filled(size, 999999.0),
        maxav=_filled(size, -999999.0),
        minav=_filled(size, 999999.0),
        maxam=_filled(size, -999999.0),
        minam=_filled(size, 999999.0),
        maxtv=_filled(size, -999999.0),
        mintv=_filled(size, 999999.0),
        maxtm=_filled(size, -999999.0),
        mintm=_filled(size, 999999.0),
        maxdev=_filled(size, -999999.0),
        mindev=_filled(size, 999999.0),
        maxattr=_filled(size, -999999.0),
        minattr=_filled(size, 999999.0),
        maxtp=_filled(size, -999999.0),
        mintp=_filled(size, 999999.0),
        maxsq=[-999999.0 for _ in range(size)],
    )


def _zeros(size: int) -> list[list[float]]:
    return [[0.0 for _ in range(size)] for _branch in range(3)]


def _filled(size: int, value: float) -> list[list[float]]:
    return [[value for _ in range(size)] for _branch in range(3)]
