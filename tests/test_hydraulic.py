from winsif_mon.models import HydraulicState
from winsif_mon.services.hydraulic import calculate_hydraulic


def test_hydraulic_totals_non_f_zeta_include_rows_5_6_7():
    state = HydraulicState(
        reference_values=[
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [1.0, 2.0, 3.0],
            [10.0, 20.0, 30.0],
            [100.0, 200.0, 300.0],
        ],
        thermal_delta=30.0,
        extra_stroke_1=0.3,
        extra_stroke_2=0.1,
        ring_length=1000.0,
        zeta="A",
    )

    result = calculate_hydraulic(state)

    assert round(result.thermal_stroke, 2) == 0.18
    assert round(result.total_nominal, 2) == 111.58
    assert round(result.total_plus_10, 2) == 222.58
    assert round(result.total_minus_10, 2) == 333.58


def test_hydraulic_totals_f_zeta_skip_row_5():
    state = HydraulicState(
        reference_values=[
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [0.0, 0.0, 0.0],
            [999.0, 999.0, 999.0],
            [10.0, 20.0, 30.0],
            [100.0, 200.0, 300.0],
        ],
        thermal_delta=0.0,
        extra_stroke_1=0.3,
        extra_stroke_2=0.1,
        zeta="F",
    )

    result = calculate_hydraulic(state)

    assert round(result.total_nominal, 2) == 110.4
    assert round(result.total_plus_10, 2) == 220.4
    assert round(result.total_minus_10, 2) == 330.4
