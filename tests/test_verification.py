import pytest

from winsif_mon.domain.geometry import load_geometry_state
from winsif_mon.domain.verification import (
    MatrixState,
    cycle_matrix_state,
    load_custom_load_state,
    load_verification_state,
    reset_matrix,
    validate_custom_loads,
)


def test_f05_verification_defaults_are_loaded_from_workbook():
    state = load_verification_state()

    assert len(state.hypotheses) == 15
    assert len(state.conditions) == 6
    assert state.hypotheses[0].label == "Ascent laden / Desc. empty"
    assert state.conditions[0].label == "Standstill plant"
    assert state.matrix[0][1] is MatrixState.NORMAL
    assert state.parameters.rope_loop_length_m == pytest.approx(1913.7300799338823)
    assert state.parameters.total_cars == 128.0
    assert state.parameters.running_speed_m_s == 2.5


def test_matrix_state_cycle_and_reset():
    state = load_verification_state()

    assert cycle_matrix_state(MatrixState.OFF) is MatrixState.NORMAL
    assert cycle_matrix_state(MatrixState.NORMAL) is MatrixState.ALTERNATE
    assert cycle_matrix_state(MatrixState.ALTERNATE) is MatrixState.OFF

    reset_matrix(state)
    assert all(cell is MatrixState.OFF for row in state.matrix for cell in row)


def test_f06_custom_load_defaults_use_geometry_span_names():
    geometry_state = load_geometry_state()
    state = load_custom_load_state(geometry_state=geometry_state)

    assert state.rows[0].ascent_span == "SM-AV"
    assert state.rows[0].descent_span == "SM-AV"
    assert state.rows[1].ascent_span == "AV-R1"
    assert len(state.rows[0].loads) == 20


def test_custom_load_validation_rejects_non_numeric_values():
    state = load_custom_load_state()
    state.rows[0].loads[0] = "bad"

    errors = validate_custom_loads(state)

    assert errors
    assert "not numeric" in errors[0]
