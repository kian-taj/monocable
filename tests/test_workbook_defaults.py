from winsif_mon.domain.f01_general import load_f01_defaults, load_rope_catalog
from winsif_mon.domain.terrain import load_terrain_profile


def test_f01_defaults_are_loaded_from_workbook():
    data = load_f01_defaults()

    assert data.value("plant_description") == 'SEGGIOVIA "PAIN DELLE GUIDE - CORNO CANETO"'
    assert data.value("carrying_rope_outer_diameter") == 30.0
    assert data.value("working_speed") == 2.5
    assert data.units["working_speed"] == "[m/s]"


def test_rope_catalog_loads_named_rows():
    rows = load_rope_catalog("ROPE_1")
    selected = next(row for row in rows if row.description == "SEALE 114" and row.diameter == 30.0)

    assert selected.metallic_section == 355.5
    assert selected.breaking_load_1960 == 697.0
    assert selected.minimum_failure_1960 == 599.0


def test_f02_terrain_defaults_are_loaded_from_workbook():
    profile = load_terrain_profile()

    assert len(profile.ascent) >= 32
    assert len(profile.descent) >= 32
    assert profile.ascent[0].progressive_distance_m == 0.0
    assert profile.ascent[0].ground_elevation_m == 1205.47
    assert profile.descent[-1].progressive_distance_m == 1200.0
