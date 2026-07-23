import pytest

from winsif_mon.domain.labels import load_labels
from winsif_mon.domain.laying_tension import load_laying_tension_state
from winsif_mon.domain.line_results import ResultFamily, active_result_cases
from winsif_mon.domain.max_min_results import load_max_min_case
from winsif_mon.domain.power_summary import load_power_summary_case
from winsif_mon.domain.report import build_general_report
from winsif_mon.domain.verification import load_verification_state
from winsif_mon.domain.verification_plots import load_power_trace_case


def test_f09_power_trace_reads_store13_defaults():
    verification = load_verification_state()
    reference = active_result_cases(verification)[0]

    trace = load_power_trace_case(ResultFamily.NORMAL, reference, verification)

    assert len(trace.forward) == 47
    assert len(trace.reverse) == 28
    assert trace.forward[0].offset_m == 0.0
    assert trace.forward[0].ascent_tension_da_n == pytest.approx(7471.770942704483)
    assert trace.forward[0].motive_force_da_n == pytest.approx(3897.945301459862)


def test_f10_max_min_reads_store05_defaults():
    case = load_max_min_case(ResultFamily.NORMAL)

    assert len(case.ascent_rows) == 74
    assert len(case.descent_rows) == 74
    assert case.ascent_rows[0].span_start == "Valle"
    assert case.ascent_rows[0].span_end == "AV"
    assert case.ascent_rows[0].tension_da_n == pytest.approx(9837.15372805009)
    assert case.ascent_rows[2].support_code == "AV"
    assert case.ascent_rows[2].roller_count == 2.0


def test_f11_power_summary_reads_store13_defaults():
    case = load_power_summary_case(ResultFamily.NORMAL)

    assert len(case.rows) == 4
    assert case.rows[0].direction == "Forward"
    assert case.rows[0].case_index == 2
    assert case.rows[0].plant_condition == "steady state plant"
    assert case.rows[0].power_kw == pytest.approx(113.57244388591899)


def test_f12_report_uses_named_sections_and_sources():
    report = build_general_report(ResultFamily.NORMAL)

    assert "SEGGIOVIA" in report.title
    assert [section.title for section in report.sections] == ["Plant", "Line", "Results"]
    assert report.sections[2].rows[-1][1].startswith("Workbook STORE05")


def test_f13_labels_are_internal_translation_data():
    labels = load_labels()

    assert labels.load_hypotheses[0] == "ASCENT LADEN - DESCENT EMPTY"
    assert labels.plant_conditions[1] == "steady state plant"
    assert labels.branch_names == ("ASCENT BRANCH", "DESCENT BRANCH")


def test_f20_laying_tension_uses_geometry_and_temperature_curve():
    state = load_laying_tension_state()

    assert state.reference_tension_da_n == 25000.0
    assert len(state.ascent_rows) == 19
    assert len(state.descent_rows) == 19
    assert state.temperature_curve[0] == pytest.approx((-20, 25060.0, 25057.5))
