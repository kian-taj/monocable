from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import (
    QButtonGroup,
    QComboBox,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.f01_general import (
    GENERAL_FIELD_SPECS,
    RopeCatalogRow,
    load_f01_defaults,
    load_rope_catalog,
)
from winsif_mon.ui.controls import set_editor_value, value_editor


class F01Page(QWidget):
    """Faithful form-style implementation of the workbook `F01` page."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.data = load_f01_defaults()
        self.carrying_catalog = load_rope_catalog("ROPE_1")
        self.tension_catalog = load_rope_catalog("ROPE_2")
        self.signal_catalog = load_rope_catalog("ROPE_3")
        self.editors: dict[str, QWidget] = {}

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        self._content_layout = QVBoxLayout(content)
        self._content_layout.setSpacing(8)
        self._build_general()
        self._build_rope_section(
            title="CARRYING-HAULING ROPE",
            combo_name="carrying",
            catalog=self.carrying_catalog,
            default_diameter=self.data.value("carrying_rope_outer_diameter"),
            setup=self._setup_carrying_rope,
            names=[
                "carrying_rope_description",
                "carrying_rope_outer_diameter",
                "carrying_rope_metallic_section",
                "carrying_rope_unit_weight",
                "carrying_rope_outer_wire_diameter",
                "carrying_rope_min_break_load",
                "carrying_rope_modulus_elasticity",
            ],
            grades=True,
        )
        self._build_rope_section(
            title="TENSION ROPE",
            combo_name="tension",
            catalog=self.tension_catalog,
            default_diameter=self.data.value("tension_rope_outer_diameter"),
            setup=self._setup_tension_rope,
            names=[
                "tension_rope_description",
                "tension_rope_outer_diameter",
                "tension_rope_metallic_section",
                "tension_rope_unit_weight",
                "tension_rope_max_wire_diameter",
                "tension_rope_min_break_load",
                "tension_rope_branches",
            ],
            grades=False,
        )
        self._build_rope_section(
            title="SIGNAL CABLE",
            combo_name="signal",
            catalog=self.signal_catalog,
            default_diameter=self.data.value("signal_cable_outer_diameter"),
            setup=self._setup_signal_cable,
            names=[
                "signal_cable_description",
                "signal_cable_outer_diameter",
                "signal_cable_metallic_section",
                "signal_cable_unit_weight",
                "signal_cable_ice_sleeve_thickness",
                "signal_cable_min_break_load",
            ],
            grades=False,
        )
        self._build_technical_data()
        self._content_layout.addStretch(1)
        scroll.setWidget(content)
        outer.addWidget(scroll)

    def _build_general(self) -> None:
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)
        grid.addWidget(_section_header("GENERAL"), 0, 0, 1, 5)
        for row, name in enumerate(("plant_description", "plant_location", "committee_details"), start=1):
            self._add_field_row(grid, row, name, wide=True)
        self._content_layout.addLayout(grid)

    def _build_rope_section(
        self,
        title: str,
        combo_name: str,
        catalog: list[RopeCatalogRow],
        default_diameter: Any,
        setup,
        names: list[str],
        grades: bool,
    ) -> None:
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)
        header = _section_header(title)
        grid.addWidget(header, 0, 0, 1, 2)
        combo = QComboBox()
        for row in catalog:
            combo.addItem(f"{row.diameter:g}", row)
        selected = _find_diameter(catalog, default_diameter)
        if selected >= 0:
            combo.setCurrentIndex(selected)
        grid.addWidget(combo, 0, 2, 1, 2)
        setup_button = QPushButton("Setup")
        setup_button.clicked.connect(lambda: setup(combo.currentData()))
        grid.addWidget(setup_button, 0, 4)
        if grades:
            grid.addWidget(_radio_row(["1770 N/mm2", "1960 N/mm2", "2160 N/mm2"], checked=1), 0, 5, 1, 4)
        for row, name in enumerate(names, start=1):
            self._add_field_row(grid, row, name)
        self._content_layout.addLayout(grid)

    def _build_technical_data(self) -> None:
        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)
        grid.addWidget(_section_header("TECHNICAL DATA"), 0, 0, 1, 8)
        grid.addWidget(QLabel("PLANT TYPE:"), 1, 0)
        grid.addWidget(_radio_row(["FIXED GRIP", "DETACHABLE CHAIRLIFT", "GONDOLA ROPEWAY"], checked=0), 1, 1, 1, 4)
        grid.addWidget(_radio_row(["PULSED-MOVEMENT AERIAL"], checked=0), 1, 5, 1, 2)

        option_rows = [
            ("Direction of rotation  (Clockwise - Anticlockwise)", ["CLOCKWISE", "ANTICLOCKWISE"], 0),
            ("Power unit position (Downstream- Upstream)", ["TOP", "BOTTOM"], 1),
            ("Tightener position  (Downstream- Upstream)", ["TOP", "BOTTOM"], 1),
            ("Type (gravity  -  hydraulic) of the  tightener", ["GRAVITY", "HYDRAUL"], 1),
            ("Carrying-tension rope loop  (under tension - anchored fixed)", ["IN TENS.", "FIXED GRIP"], 0),
        ]
        for offset, (label, options, checked) in enumerate(option_rows, start=2):
            grid.addWidget(_label(label), offset, 0, 1, 2)
            grid.addWidget(_radio_row(options, checked=checked), offset, 2, 1, 3)

        grid.addWidget(_label("Load distribution on the line:"), 7, 0, 1, 2)
        grid.addWidget(_radio_row(["ONLY ASCENT", "ASCENT AND DESCENT"], checked=1), 7, 2, 1, 4)

        tech_names = [
            "car_gap_ascent",
            "rated_tightener_tension",
            "carrying_capacity_per_hour",
            "working_speed",
            "persons_per_car",
            "person_weight",
            "empty_car_weight",
            "laden_car_weight",
            "startup_acceleration",
            "stopping_deceleration_1",
            "stopping_deceleration_2",
            "stopping_deceleration_3",
            "power_unit_equivalent_weight",
            "power_unit_efficiency",
            "friction_running",
            "friction_braking",
            "station_rope_deviation_angle",
            "acceleration_beam_force",
            "driving_pulley_diameter",
            "snub_pulley_diameter",
            "rope_distance",
            "car_type",
            "clamps_per_car",
            "empty_car_crosswind_surface",
            "laden_car_crosswind_surface",
            "running_crosswind_thrust",
            "out_of_service_crosswind_thrust",
            "roller_type",
            "support_roller_weight",
            "compression_roller_weight",
            "support_roller_diameter",
            "compression_roller_diameter",
            "max_support_roller_deviation",
            "max_compression_roller_deviation",
            "max_support_roller_load",
            "max_compression_roller_load",
            "double_acting_roller_type",
            "double_acting_roller_diameter",
            "double_acting_roller_weight",
            "double_acting_roller_max_load",
            "vehicle_vertical_height",
            "vehicle_width",
            "vehicle_admissible_inclination",
            "pulley_envelopment_corner",
            "pulley_rope_friction_coefficient",
            "valley_station_time",
            "mountain_station_time",
        ]
        for i, name in enumerate(tech_names, start=8):
            self._add_field_row(grid, i, name)
        self._content_layout.addLayout(grid)

    def _add_field_row(self, grid: QGridLayout, row: int, name: str, wide: bool = False) -> None:
        spec = next(spec for spec in GENERAL_FIELD_SPECS if spec.name == name)
        grid.addWidget(_label(spec.label), row, 0, 1, 2)
        unit = self.data.units.get(name, "")
        editor = value_editor(
            self.data.value(name),
            unit,
            lambda value, n=name: self.data.set_value(n, value),
        )
        editor.setObjectName(name)
        self.editors[name] = editor
        span = 5 if wide else 2
        grid.addWidget(editor, row, 2, 1, span)

    def _setup_carrying_rope(self, row: RopeCatalogRow) -> None:
        self._apply_rope(row, "carrying_rope", include_modulus=True)
        self._set_field("carrying_rope_min_break_load", row.breaking_load_1960)

    def _setup_tension_rope(self, row: RopeCatalogRow) -> None:
        self._apply_rope(row, "tension_rope")
        self._set_field("tension_rope_min_break_load", row.breaking_load_1960)

    def _setup_signal_cable(self, row: RopeCatalogRow) -> None:
        self._set_field("signal_cable_description", row.description)
        self._set_field("signal_cable_outer_diameter", row.diameter)
        self._set_field("signal_cable_metallic_section", row.metallic_section)
        self._set_field("signal_cable_unit_weight", row.unit_weight)
        self._set_field("signal_cable_min_break_load", row.breaking_load_1960)

    def _apply_rope(self, row: RopeCatalogRow, prefix: str, include_modulus: bool = False) -> None:
        self._set_field(f"{prefix}_description", row.description)
        self._set_field(f"{prefix}_outer_diameter", row.diameter)
        self._set_field(f"{prefix}_metallic_section", row.metallic_section)
        self._set_field(f"{prefix}_unit_weight", row.unit_weight)
        wire_name = f"{prefix}_outer_wire_diameter" if prefix == "carrying_rope" else f"{prefix}_max_wire_diameter"
        self._set_field(wire_name, row.wire_diameter)

    def _set_field(self, name: str, value: Any) -> None:
        self.data.set_value(name, value)
        if name in self.editors:
            set_editor_value(self.editors[name], value)


def _section_header(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("sectionHeader")
    return label


def _label(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("fieldLabel")
    return label

def _radio_row(options: list[str], checked: int = 0) -> QWidget:
    widget = QWidget()
    layout = QHBoxLayout(widget)
    layout.setContentsMargins(0, 0, 0, 0)
    group = QButtonGroup(widget)
    for index, option in enumerate(options):
        button = QRadioButton(option)
        group.addButton(button)
        layout.addWidget(button)
        if index == checked:
            button.setChecked(True)
    layout.addStretch(1)
    return widget


def _find_diameter(catalog: list[RopeCatalogRow], diameter: Any) -> int:
    try:
        target = float(diameter)
    except (TypeError, ValueError):
        return -1
    for index, row in enumerate(catalog):
        if row.diameter == target:
            return index
    return -1
