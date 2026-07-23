from winsif_mon.domain.geometry import (
    convert_spans_to_supports,
    convert_supports_to_spans,
    interpolate_ground_elevation,
    load_geometry_state,
)
from winsif_mon.domain.terrain import load_terrain_profile


def test_f03_support_geometry_defaults_are_loaded_from_workbook():
    state = load_geometry_state()

    first = state.ascent_supports[0]
    assert first.line_number == 1
    assert first.code == "Valle"
    assert first.rope_distance_m == 30.0
    assert round(first.ground_elevation_m, 6) == round(1205.0018354818908, 6)
    assert first.roller_quantity == 0.0
    assert state.descent_supports[0].line_number == 1


def test_f04_span_geometry_defaults_are_loaded_from_workbook():
    state = load_geometry_state()

    assert state.first_progressive_distance_m == 31.02
    assert state.first_rope_height_m == 1208.5699462890625
    assert state.ascent_spans[0].code == "SM-AV"
    assert state.ascent_spans[0].horizontal_distance_m == 2.5
    assert state.ascent_spans[0].height_difference_m == 0.0


def test_supports_convert_to_relative_spans():
    state = load_geometry_state()

    spans = convert_supports_to_spans(state.ascent_supports)

    assert spans[0].code == "Valle-AV"
    assert round(spans[0].horizontal_distance_m, 2) == 3.52
    assert round(spans[0].height_difference_m, 2) == 0.0
    assert spans[1].code == "AV-C1"
    assert round(spans[1].horizontal_distance_m, 2) == 15.0


def test_relative_spans_convert_to_absolute_supports():
    state = load_geometry_state()
    profile = load_terrain_profile()

    supports = convert_spans_to_supports(
        state.ascent_spans[:3],
        state.first_progressive_distance_m,
        state.first_rope_height_m,
        profile,
        "ascent",
    )

    assert supports[0].code == "SM"
    assert supports[1].code == "AV"
    assert round(supports[1].rope_distance_m, 2) == 33.52
    assert round(supports[1].rope_elevation_m, 2) == 1208.57
    assert round(supports[2].rope_distance_m, 2) == 48.52


def test_terrain_interpolation_matches_f03_default_point():
    profile = load_terrain_profile()

    elevation = interpolate_ground_elevation(profile, "ascent", 30.0)

    assert round(elevation, 6) == round(1205.0018354818908, 6)
