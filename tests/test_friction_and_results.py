import pytest

from winsif_mon.domain.friction import (
    FrictionMode,
    assign_default_friction,
    load_friction_assignment_state,
    validate_friction_state,
)
from winsif_mon.domain.geometry import load_geometry_state
from winsif_mon.domain.line_results import ResultFamily, active_result_cases, load_line_result_case
from winsif_mon.domain.verification import load_verification_state


def test_f07_friction_defaults_match_workbook():
    state = load_friction_assignment_state()

    assert state.settings.default_mode is FrictionMode.PERCENT
    assert state.settings.steady_value == 3.0
    assert state.settings.braking_value == 2.0
    assert len(state.ascent_rows) == 19
    assert len(state.descent_rows) == 19
    assert state.ascent_rows[0].support_code == "AV"
    assert state.ascent_rows[0].mode is FrictionMode.PERCENT
    assert validate_friction_state(state) == []


def test_f07_assign_defaults_uses_internal_supports_only():
    geometry = load_geometry_state()
    state = load_friction_assignment_state()

    rows = assign_default_friction(geometry.ascent_supports, state.settings)

    assert len(rows) == 18
    assert rows[0].support_code == "AV"
    assert rows[-1].support_code == "AM"
    assert all(row.steady_value == 3.0 for row in rows)


def test_f08_store06_parser_reads_default_active_case():
    verification = load_verification_state()
    reference = active_result_cases(verification)[0]

    result = load_line_result_case(ResultFamily.NORMAL, reference, verification)

    assert result.source == "Workbook STORE06 defaults"
    assert result.reference.hypothesis_index == 1
    assert result.reference.condition_index == 2
    assert len(result.ascent.rows) == 74
    assert len(result.descent.rows) == 74
    assert result.ascent.rows[0].span_start == "Valle"
    assert result.ascent.rows[0].span_end == "AV"
    assert result.ascent.rows[0].valley_tension_da_n == pytest.approx(7530.783599668728)
    assert result.ascent.rows[2].support_code == "AV"
    assert result.ascent.rows[2].roller_count == 2.0
