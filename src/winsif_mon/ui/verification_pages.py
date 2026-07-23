from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QDoubleSpinBox,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.geometry import GeometryState
from winsif_mon.domain.f01_general import load_f01_defaults
from winsif_mon.domain.verification import (
    CustomLoadState,
    MatrixState,
    PlantRunParameters,
    SpanLoadRow,
    VerificationState,
    build_calculation_setup,
    cycle_matrix_state,
    load_custom_load_state,
    load_verification_state,
    reset_matrix,
)


class LineVerificationPage(QWidget):
    def __init__(
        self,
        verification_state: VerificationState | None = None,
        custom_load_state: CustomLoadState | None = None,
        actions: list[QPushButton] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.verification_state = verification_state or load_verification_state()
        self.custom_load_state = custom_load_state or load_custom_load_state()
        self._loading = False

        layout = QVBoxLayout(self)
        heading = QLabel("Line Verification")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        if actions:
            action_row = QHBoxLayout()
            for action in actions:
                action_row.addWidget(action)
            action_row.addStretch(1)
            layout.addLayout(action_row)

        layout.addLayout(self._parameter_grid(self.verification_state.parameters))

        matrix_header = QLabel("Load Hypotheses And Plant Conditions")
        matrix_header.setObjectName("sectionHeader")
        layout.addWidget(matrix_header)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self._button("Reset Matrix", self.reset_matrix))
        toolbar.addWidget(self._button("Validate Setup", self.validate_setup))
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        self.matrix_table = QTableWidget(
            len(self.verification_state.hypotheses),
            len(self.verification_state.conditions) + 1,
        )
        self.matrix_table.setObjectName("verification_matrix")
        self.matrix_table.setAlternatingRowColors(True)
        self.matrix_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.matrix_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.matrix_table.cellDoubleClicked.connect(self.cycle_cell)
        self.matrix_table.setHorizontalHeaderLabels(
            ["Load hypothesis"] + [condition.label for condition in self.verification_state.conditions]
        )
        self.matrix_table.setColumnWidth(0, 270)
        for col in range(1, self.matrix_table.columnCount()):
            self.matrix_table.setColumnWidth(col, 116)
        layout.addWidget(self.matrix_table, 1)
        self.refresh_matrix()

    def reset_matrix(self) -> None:
        reset_matrix(self.verification_state)
        self.refresh_matrix()

    def cycle_cell(self, row: int, col: int) -> None:
        if col == 0:
            return
        self.verification_state.matrix[row][col - 1] = cycle_matrix_state(
            self.verification_state.matrix[row][col - 1]
        )
        self.refresh_matrix()

    def validate_setup(self) -> None:
        try:
            build_calculation_setup(self.verification_state, self.custom_load_state)
        except ValueError as exc:
            QMessageBox.warning(self, "Invalid custom loads", str(exc))
            return
        QMessageBox.information(self, "Calculation setup", "Inputs are valid for the calculation service boundary.")

    def refresh_matrix(self) -> None:
        self._loading = True
        for row, hypothesis in enumerate(self.verification_state.hypotheses):
            label = QTableWidgetItem(f"{hypothesis.index}. {hypothesis.label}")
            label.setFlags(label.flags() & ~Qt.ItemIsEditable)
            self.matrix_table.setItem(row, 0, label)
            for col, state in enumerate(self.verification_state.matrix[row], start=1):
                item = QTableWidgetItem(state.value)
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(_state_color(state))
                self.matrix_table.setItem(row, col, item)
        self._loading = False

    def _parameter_grid(self, parameters: PlantRunParameters) -> QGridLayout:
        grid = QGridLayout()
        values = [
            ("Rope loop length", parameters.rope_loop_length_m, "m"),
            ("Total cars", parameters.total_cars, "n"),
            ("Distance between cars", parameters.car_spacing_m, "m"),
            ("Running speed", parameters.running_speed_m_s, "m/s"),
            ("Carrying capacity", parameters.carrying_capacity_p_h, "p/h"),
            ("Step advancement", parameters.step_advancement_m, "m"),
            ("Advancement steps", parameters.advancement_steps, "n"),
            ("Tightener precision", parameters.tightener_precision_percent, "%"),
            ("Local temperature", parameters.local_temperature_c, "deg C"),
        ]
        for index, (label, value, unit) in enumerate(values):
            row = index // 3
            col = (index % 3) * 2
            grid.addWidget(QLabel(label), row, col)
            editor = _readonly_double(value, unit)
            editor.setObjectName(label.lower().replace(" ", "_"))
            grid.addWidget(editor, row, col + 1)
        return grid

    def _button(self, text: str, slot) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button


class CustomLoadCasesPage(QWidget):
    def __init__(
        self,
        custom_load_state: CustomLoadState | None = None,
        geometry_state: GeometryState | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.custom_load_state = custom_load_state or load_custom_load_state(geometry_state=geometry_state)
        self._loading = False

        layout = QVBoxLayout(self)
        heading = QLabel("Custom Load Cases")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        controls = QHBoxLayout()
        self.load_kind = QComboBox()
        for label, value in self._load_options():
            self.load_kind.addItem(label, value)
        self.load_kind.currentIndexChanged.connect(self._sync_selected_load_value)
        controls.addWidget(QLabel("Vehicle load"))
        controls.addWidget(self.load_kind)

        self.custom_value = QDoubleSpinBox()
        self.custom_value.setRange(-1_000_000.0, 1_000_000.0)
        self.custom_value.setDecimals(3)
        self.custom_value.setSuffix(" kg")
        controls.addWidget(self.custom_value)
        self._sync_selected_load_value()

        self.scope_group = QButtonGroup(self)
        self.selected_span = QRadioButton("Selected span")
        self.all_spans = QRadioButton("All spans")
        self.selected_span.setChecked(True)
        self.scope_group.addButton(self.selected_span)
        self.scope_group.addButton(self.all_spans)
        controls.addWidget(self.selected_span)
        controls.addWidget(self.all_spans)

        self.load_target = QComboBox()
        self.load_target.addItem("Selected cell", "cell")
        self.load_target.addItem("Selected hypothesis", "hypothesis")
        self.load_target.addItem("All load cases", "all")
        controls.addWidget(QLabel("Target"))
        controls.addWidget(self.load_target)

        controls.addWidget(self._button("Apply Load", self.apply_load))
        controls.addWidget(self._button("Reset Selected Span", self.reset_selected_span))
        controls.addWidget(self._button("Reset All Loads", self.reset_all_loads))
        controls.addStretch(1)
        layout.addLayout(controls)

        self.table = QTableWidget(len(self.custom_load_state.rows), 22)
        self.table.setObjectName("custom_load_cases")
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setHorizontalHeaderLabels(_custom_load_headers())
        self.table.cellChanged.connect(lambda _row, _col: self._save_to_state())
        self.table.setColumnWidth(0, 120)
        self.table.setColumnWidth(1, 120)
        for col in range(2, 22):
            self.table.setColumnWidth(col, 86)
        layout.addWidget(self.table, 1)
        self.refresh_table()

    def refresh_table(self) -> None:
        self._loading = True
        self.table.setRowCount(max(len(self.custom_load_state.rows), 24))
        for row in range(self.table.rowCount()):
            if row < len(self.custom_load_state.rows):
                source = self.custom_load_state.rows[row]
                values = [source.ascent_span, source.descent_span] + source.loads
            else:
                values = ["", ""] + [""] * 20
            for col, value in enumerate(values):
                item = QTableWidgetItem(_format(value))
                if col >= 2:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, col, item)
        self._loading = False

    def apply_load(self) -> None:
        value = self.load_kind.currentData()
        if value is None:
            value = self.custom_value.value()
        rows = range(self.table.rowCount()) if self.all_spans.isChecked() else [max(self.table.currentRow(), 0)]
        col = self.table.currentColumn()
        for row in rows:
            for target_col in self._target_columns(col):
                self.table.setItem(row, target_col, QTableWidgetItem(_format(value)))
        self._save_to_state()

    def reset_selected_span(self) -> None:
        row = self.table.currentRow()
        if row < 0:
            row = 0
        for col in range(2, self.table.columnCount()):
            self.table.setItem(row, col, QTableWidgetItem(""))
        self._save_to_state()

    def reset_all_loads(self) -> None:
        for row in range(self.table.rowCount()):
            for col in range(2, self.table.columnCount()):
                self.table.setItem(row, col, QTableWidgetItem(""))
        self._save_to_state()

    def _target_columns(self, current_col: int) -> list[int]:
        target = self.load_target.currentData()
        if target == "all":
            return list(range(2, self.table.columnCount()))
        if current_col < 2:
            current_col = 2
        if target == "hypothesis":
            start = current_col if current_col % 2 == 0 else current_col - 1
            return [start, start + 1]
        return [current_col]

    def _save_to_state(self) -> None:
        if self._loading:
            return
        rows = []
        for row in range(self.table.rowCount()):
            ascent_span = _cell_text(self.table, row, 0)
            descent_span = _cell_text(self.table, row, 1)
            loads = [_value(_cell_text(self.table, row, col)) for col in range(2, 22)]
            if ascent_span or descent_span or any(value != "" for value in loads):
                rows.append(SpanLoadRow(ascent_span=ascent_span, descent_span=descent_span, loads=loads))
        self.custom_load_state.rows = rows

    def _button(self, text: str, slot: Callable) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _load_options(self) -> list[tuple[str, float | None]]:
        defaults = load_f01_defaults()
        return [
            ("Loaded", float(defaults.value("laden_car_weight") or 0.0)),
            ("Empty", float(defaults.value("empty_car_weight") or 0.0)),
            ("Zero", 0.0),
            ("Custom", None),
        ]

    def _sync_selected_load_value(self) -> None:
        value = self.load_kind.currentData()
        self.custom_value.setReadOnly(value is not None)
        if value is not None:
            self.custom_value.setValue(float(value))


def _readonly_double(value: float, unit: str) -> QDoubleSpinBox:
    editor = QDoubleSpinBox()
    editor.setRange(-1_000_000_000.0, 1_000_000_000.0)
    editor.setDecimals(4)
    editor.setValue(value)
    editor.setSuffix(f" {unit}")
    editor.setReadOnly(True)
    editor.setButtonSymbols(QDoubleSpinBox.NoButtons)
    return editor


def _state_color(state: MatrixState) -> QColor:
    if state is MatrixState.NORMAL:
        return QColor("#dbeafe")
    if state is MatrixState.ALTERNATE:
        return QColor("#fee2e2")
    return QColor("#ffffff")


def _custom_load_headers() -> list[str]:
    headers = ["Ascent span", "Descent span"]
    for index in range(1, 11):
        headers.extend([f"H{index} ascent", f"H{index} descent"])
    return headers


def _cell_text(table: QTableWidget, row: int, col: int) -> str:
    item = table.item(row, col)
    return "" if item is None else item.text().strip()


def _value(text: str) -> float | str:
    if text == "":
        return ""
    try:
        return float(text)
    except ValueError:
        return text


def _format(value) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    return str(value)
