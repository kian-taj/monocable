import pytest

from winsif_mon.services.solver import (
    SolverMode,
    build_line_input,
    derive_run_parameters,
    position_vehicles,
    prepare_loads,
    run_solver_setup,
)
from winsif_mon.services.iterative_solver import run_iterative_solver
from winsif_mon.domain.friction import load_friction_assignment_state


def test_solver_line_input_matches_f05_default_run_parameters():
    line_input = build_line_input()

    run_parameters = derive_run_parameters(line_input)

    assert line_input.ascent.developed_length_m == pytest.approx(950.5830399669411)
    assert line_input.descent.developed_length_m == pytest.approx(950.5830399669411)
    assert run_parameters.rope_loop_length_m == pytest.approx(1913.7300799338823)
    assert run_parameters.total_cars == 128.0
    assert run_parameters.car_spacing_m == 15.0
    assert run_parameters.step_m == 1.5
    assert len(line_input.ascent_friction) == 19
    assert line_input.ascent_friction[0].tower_code == "AV"
    assert line_input.ascent_friction[0].steady_value == 3.0


def test_vehicle_positions_match_vba_starting_pattern_for_default_offset():
    line_input = build_line_input()
    run_parameters = derive_run_parameters(line_input)

    positions = position_vehicles(
        offset_m=0.0,
        ascent_length_m=line_input.ascent.developed_length_m,
        descent_length_m=line_input.descent.developed_length_m,
        drive_pulley_diameter_mm=line_input.general.drive_pulley_diameter_mm,
        car_spacing_m=run_parameters.vehicle_spacing_m,
        vehicles_per_group=run_parameters.vehicles_per_group,
        intra_group_spacing_m=run_parameters.car_spacing_m,
    )

    assert positions.ascent[:5] == [0.0, 15.0, 30.0, 45.0, 60.0]
    assert positions.descent[:3] == pytest.approx([8.718, 23.718, 38.718])
    assert len(positions.ascent) == 64
    assert len(positions.descent) == 63


def test_solver_prepares_standard_and_custom_loads():
    line_input = build_line_input()

    standard = prepare_loads(line_input, hypothesis_index=1, condition_index=2)
    bare = prepare_loads(line_input, hypothesis_index=5, condition_index=1)

    assert standard.ascent_loads_da_n[:3] == pytest.approx([245.25, 245.25, 245.25])
    assert standard.descent_loads_da_n[:3] == pytest.approx([88.29, 88.29, 88.29])
    assert all(value == 0.0 for value in bare.ascent_loads_da_n)
    assert all(value == 0.0 for value in bare.descent_loads_da_n)


def test_solver_setup_runs_active_matrix_cases_without_full_tension_loop():
    line_input = build_line_input()

    result = run_solver_setup(SolverMode.NORMAL, line_input)

    assert result.parity_complete is False
    assert len(result.active_cases) == 2
    assert len(result.prepared_loads) == 2
    assert result.run_parameters.rope_loop_length_m == pytest.approx(1913.7300799338823)
    assert result.native_iterative_result is not None
    assert len(result.native_iterative_result.cases) == 2


def test_solver_uses_edited_f07_friction_state():
    friction_state = load_friction_assignment_state()
    friction_state.ascent_rows[0].steady_value = 4.5

    line_input = build_line_input(friction_state=friction_state)

    assert line_input.ascent_friction[0].steady_value == 4.5


def test_native_iterative_solver_runs_default_normal_cases():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.NORMAL, line_input)

    assert result.parity_complete is False
    assert len(result.cases) == 2
    assert result.cases[0].reference.hypothesis_index == 1
    assert result.cases[0].reference.condition_index == 2
    assert len(result.cases[0].forward_trace) == 11
    assert len(result.cases[0].reverse_trace) == 11
    assert result.cases[0].forward_line_results.source == "Native iterative solver"
    assert result.cases[0].power_trace.source == "Native iterative solver"
    assert result.cases[0].power_summary.source == "Native iterative solver"
    assert result.max_min_results[result.cases[0].family].source == "Native iterative solver"


def test_native_iterative_solver_runs_variable_tension_families():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.VARIABLE_TENSION, line_input)

    families = [case.family.value for case in result.cases]
    assert families == ["plus_ten", "plus_ten", "minus_ten", "minus_ten"]
    assert result.cases[0].reference.store_index == result.cases[2].reference.store_index
    assert result.cases[0].forward_trace[0].ascent_tension_da_n > result.cases[2].forward_trace[0].ascent_tension_da_n


def test_native_iterative_solver_runs_anchored_family():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.ANCHORED, line_input)

    assert result.mode is SolverMode.ANCHORED
    assert len(result.cases) == 2
    assert result.cases[0].family.value == "normal"
    assert result.cases[0].forward_trace[0].ascent_tension_da_n == pytest.approx(8886.891205614846)
    assert result.cases[0].forward_trace[0].motive_force_da_n == pytest.approx(4066.950650551167)


def test_native_iterative_solver_builds_hydraulic_precheck_matrix():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.HYDRAULIC_PRECHECK, line_input)

    assert result.mode is SolverMode.HYDRAULIC_PRECHECK
    assert result.cases == []
    assert result.hydraulic_reference_values is not None
    assert len(result.hydraulic_reference_values) == 7
    assert all(len(row) == 3 for row in result.hydraulic_reference_values)
    assert result.hydraulic_reference_values[0][0] == pytest.approx(2.7042970035101925)
    assert result.hydraulic_reference_values[4][0] == pytest.approx(0.2523240959520576)


def test_native_iterative_solver_runs_hydraulic_final_family():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.HYDRAULIC_FINAL, line_input)

    assert result.mode is SolverMode.HYDRAULIC_FINAL
    assert len(result.cases) == 2
    assert result.cases[0].family.value == "hydraulic"
    assert result.hydraulic_reference_values is not None


def test_native_iterative_solver_first_trace_matches_store13_golden():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.NORMAL, line_input)
    first = result.cases[0].forward_trace[0]

    assert first.offset_m == pytest.approx(0.0)
    assert first.ascent_tension_da_n == pytest.approx(7471.770942704483, abs=6.0)
    assert first.descent_tension_da_n == pytest.approx(11369.716244164345, abs=6.0)
    assert first.motive_force_da_n == pytest.approx(3897.945301459862, abs=2.0)


def test_native_iterative_solver_first_span_matches_store06_golden():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.NORMAL, line_input)
    first_span = result.cases[0].forward_line_results.ascent.rows[0]

    assert first_span.span_start == "Valle"
    assert first_span.span_end == "AV"
    assert first_span.valley_tension_da_n == pytest.approx(7530.783599668728, abs=6.0)
    assert first_span.mountain_tension_da_n == pytest.approx(7530.783599668727, abs=6.0)
    assert first_span.sag_m == pytest.approx(0.025423470370434634, abs=1e-4)


def test_native_iterative_solver_max_min_aggregates_active_cases():
    line_input = build_line_input()

    result = run_iterative_solver(SolverMode.NORMAL, line_input)
    max_min = result.max_min_results[next(iter(result.max_min_results))]

    assert len(max_min.ascent_rows) == 74
    assert max_min.ascent_rows[0].span_start == "Valle"
    assert max_min.ascent_rows[0].span_end == "AV"
    assert max_min.ascent_rows[0].tension_da_n == pytest.approx(9837.15372805009, abs=6.0)
    assert max_min.ascent_rows[2].support_code == "AV"
    assert max_min.ascent_rows[2].support_tension_da_n == pytest.approx(9838.035785608645, abs=6.0)
