from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.friction import (
    FrictionAssignmentRow,
    FrictionAssignmentState,
    FrictionMode,
    assign_default_friction,
    load_friction_assignment_state,
    reset_friction_from_geometry,
    validate_friction_state,
)
from winsif_mon.domain.geometry import GeometryState, load_geometry_state
from winsif_mon.domain.line_results import (
    LineResultCase,
    LineResultRow,
    ResultCaseRef,
    ResultFamily,
    active_result_cases,
    load_line_result_case,
)
from winsif_mon.domain.verification import VerificationState, load_verification_state


class FrictionAssignmentPage(QWidget):
    def __init__(
        self,
        friction_state: FrictionAssignmentState | None = None,
        geometry_state: GeometryState | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.friction_state = friction_state or load_friction_assignment_state()
        self.geometry_state = geometry_state or load_geometry_state()
        self._loading = False

        layout = QVBoxLayout(self)
        heading = QLabel("Friction Assignment")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        layout.addLayout(self._settings_layout())
        layout.addLayout(self._actions_layout())

        self.ascent_table = self._table("ascent_friction")
        self.descent_table = self._table("descent_friction")
        tabs = QTabWidget()
        tabs.addTab(self.ascent_table, "Ascent Branch Frictions")
        tabs.addTab(self.descent_table, "Descent Branch Frictions")
        layout.addWidget(tabs, 1)
        self.refresh_tables()

    def assign_ascent_defaults(self) -> None:
        self._save_to_state()
        self.friction_state.ascent_rows = assign_default_friction(
            self.geometry_state.ascent_supports,
            self.friction_state.settings,
        )
        self.refresh_tables()

    def assign_descent_defaults(self) -> None:
        self._save_to_state()
        supports = [support for support in self.geometry_state.descent_supports if support.code]
        if len(supports) < 2:
            supports = [support for support in self.geometry_state.ascent_supports if support.code]
        self.friction_state.descent_rows = assign_default_friction(supports, self.friction_state.settings)
        self.refresh_tables()

    def reset_from_geometry(self) -> None:
        self._save_to_state()
        reset_friction_from_geometry(self.friction_state, self.geometry_state)
        self.refresh_tables()

    def validate_inputs(self) -> bool:
        self._save_to_state()
        errors = validate_friction_state(self.friction_state)
        if errors:
            QMessageBox.warning(self, "Invalid friction assignments", "\n".join(errors))
            return False
        QMessageBox.information(self, "Friction assignments", "Friction inputs are valid.")
        return True

    def refresh_tables(self) -> None:
        self._loading = True
        self._fill_table(self.ascent_table, self.friction_state.ascent_rows)
        self._fill_table(self.descent_table, self.friction_state.descent_rows)
        self._loading = False

    def _settings_layout(self) -> QFormLayout:
        form = QFormLayout()
        self.default_mode = QComboBox()
        self.default_mode.setObjectName("friction_default_mode")
        for mode in FrictionMode:
            self.default_mode.addItem(mode.label, mode)
        self.default_mode.setCurrentIndex(0 if self.friction_state.settings.default_mode is FrictionMode.PERCENT else 1)
        self.default_mode.currentIndexChanged.connect(self._save_settings)
        form.addRow("Default friction type", self.default_mode)

        self.steady_default = _double_editor(self.friction_state.settings.steady_value, "%")
        self.steady_default.setObjectName("friction_default_steady")
        self.steady_default.valueChanged.connect(lambda _value: self._save_settings())
        form.addRow("Steady state default", self.steady_default)

        self.braking_default = _double_editor(self.friction_state.settings.braking_value, "%")
        self.braking_default.setObjectName("friction_default_braking")
        self.braking_default.valueChanged.connect(lambda _value: self._save_settings())
        form.addRow("Braking default", self.braking_default)
        return form

    def _actions_layout(self) -> QHBoxLayout:
        actions = QHBoxLayout()
        actions.addWidget(_button("Assign Defaults To Ascent", self.assign_ascent_defaults))
        actions.addWidget(_button("Assign Defaults To Descent", self.assign_descent_defaults))
        actions.addWidget(_button("Reset From Support Geometry", self.reset_from_geometry))
        actions.addWidget(_button("Validate", self.validate_inputs))
        actions.addStretch(1)
        return actions

    def _table(self, name: str) -> QTableWidget:
        table = QTableWidget(0, 4)
        table.setObjectName(name)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setHorizontalHeaderLabels(["Support", "Friction type", "Steady", "Braking"])
        table.cellChanged.connect(lambda _row, _col: self._save_to_state())
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        return table

    def _fill_table(self, table: QTableWidget, rows: list[FrictionAssignmentRow]) -> None:
        table.setRowCount(max(len(rows), 1))
        for row_index in range(table.rowCount()):
            source = rows[row_index] if row_index < len(rows) else FrictionAssignmentRow("", FrictionMode.PERCENT, 0.0, 0.0)
            table.setItem(row_index, 0, QTableWidgetItem(source.support_code))
            mode = QComboBox()
            for friction_mode in FrictionMode:
                mode.addItem(friction_mode.label, friction_mode)
            mode.setCurrentIndex(0 if source.mode is FrictionMode.PERCENT else 1)
            mode.currentIndexChanged.connect(lambda _index, table=table, row=row_index: self._mode_changed(table, row))
            table.setCellWidget(row_index, 1, mode)
            steady = _double_editor(source.steady_value, source.mode.unit)
            braking = _double_editor(source.braking_value, source.mode.unit)
            steady.valueChanged.connect(lambda _value: self._save_to_state())
            braking.valueChanged.connect(lambda _value: self._save_to_state())
            table.setCellWidget(row_index, 2, steady)
            table.setCellWidget(row_index, 3, braking)

    def _save_settings(self) -> None:
        mode = _friction_mode(self.default_mode.currentData())
        self.friction_state.settings.default_mode = mode
        self.friction_state.settings.steady_value = self.steady_default.value()
        self.friction_state.settings.braking_value = self.braking_default.value()
        unit = mode.unit
        self.steady_default.setSuffix(f" {unit}")
        self.braking_default.setSuffix(f" {unit}")

    def _save_to_state(self) -> None:
        if self._loading:
            return
        self._save_settings()
        self.friction_state.ascent_rows = self._rows_from_table(self.ascent_table)
        self.friction_state.descent_rows = self._rows_from_table(self.descent_table)

    def _mode_changed(self, table: QTableWidget, row: int) -> None:
        mode_editor = table.cellWidget(row, 1)
        mode = _friction_mode(mode_editor.currentData() if isinstance(mode_editor, QComboBox) else None)
        unit = mode.unit
        for col in (2, 3):
            editor = table.cellWidget(row, col)
            if isinstance(editor, QDoubleSpinBox):
                editor.setSuffix(f" {unit}")
        self._save_to_state()

    def _rows_from_table(self, table: QTableWidget) -> list[FrictionAssignmentRow]:
        rows: list[FrictionAssignmentRow] = []
        for row in range(table.rowCount()):
            code_item = table.item(row, 0)
            code = code_item.text().strip() if code_item else ""
            if not code:
                continue
            mode_editor = table.cellWidget(row, 1)
            mode = _friction_mode(mode_editor.currentData() if isinstance(mode_editor, QComboBox) else None)
            steady_editor = table.cellWidget(row, 2)
            braking_editor = table.cellWidget(row, 3)
            rows.append(
                FrictionAssignmentRow(
                    support_code=code,
                    mode=mode,
                    steady_value=steady_editor.value() if isinstance(steady_editor, QDoubleSpinBox) else 0.0,
                    braking_value=braking_editor.value() if isinstance(braking_editor, QDoubleSpinBox) else 0.0,
                )
            )
        return rows

    def _section_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("sectionHeader")
        return label


class LineResultsPage(QWidget):
    def __init__(
        self,
        verification_state: VerificationState | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.verification_state = verification_state or load_verification_state()
        self._case_refs = active_result_cases(self.verification_state)
        self._result: LineResultCase | None = None

        layout = QVBoxLayout(self)
        heading = QLabel("Line Results")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        selectors = QHBoxLayout()
        self.family_selector = QComboBox()
        for family in ResultFamily:
            self.family_selector.addItem(family.label, family)
        self.family_selector.currentIndexChanged.connect(self.refresh_results)
        selectors.addWidget(QLabel("Result family"))
        selectors.addWidget(self.family_selector)

        self.case_selector = QComboBox()
        for case_ref in self._case_refs:
            self.case_selector.addItem(case_ref.label, case_ref)
        self.case_selector.currentIndexChanged.connect(self.refresh_results)
        selectors.addWidget(QLabel("Hypothesis and condition"))
        selectors.addWidget(self.case_selector, 1)
        selectors.addStretch(1)
        layout.addLayout(selectors)

        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        self.ascent_table = self._table("ascent_line_results")
        self.descent_table = self._table("descent_line_results")
        tabs = QTabWidget()
        tabs.addTab(self.ascent_table, "Ascent Branch Results")
        tabs.addTab(self.descent_table, "Descent Branch Results")
        layout.addWidget(tabs, 1)
        self.refresh_results()

    def refresh_results(self) -> None:
        family = self.family_selector.currentData()
        if not isinstance(family, ResultFamily):
            family = ResultFamily.NORMAL
        reference = self.case_selector.currentData()
        if not isinstance(reference, ResultCaseRef):
            reference = None
        self._result = load_line_result_case(family, reference, self.verification_state)
        self._fill_table(self.ascent_table, self._result.ascent.rows)
        self._fill_table(self.descent_table, self._result.descent.rows)
        self.summary.setText(
            f"{self._result.family.label} - {self._result.reference.label}. "
            f"Source: {self._result.source}. "
            "These rows mirror F08/STORE06 until the full iterative solver writes native results."
        )

    def _table(self, name: str) -> QTableWidget:
        table = QTableWidget(0, 15)
        table.setObjectName(name)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHorizontalHeaderLabels(
            [
                "Span start",
                "Span end",
                "Valley tension [daN]",
                "Mountain tension [daN]",
                "Sag [m]",
                "Valley angle [deg]",
                "Mountain angle [deg]",
                "Support",
                "Support tension [daN]",
                "Deviation [deg]",
                "Pressure [daN]",
                "Friction [daN]",
                "Rollers [n]",
                "Unit deviation [deg]",
                "Unit pressure [daN]",
            ]
        )
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        table.horizontalHeader().setStretchLastSection(True)
        return table

    def _fill_table(self, table: QTableWidget, rows: list[LineResultRow]) -> None:
        table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            values = [
                row.span_start,
                row.span_end,
                row.valley_tension_da_n,
                row.mountain_tension_da_n,
                row.sag_m,
                row.valley_angle_deg,
                row.mountain_angle_deg,
                row.support_code,
                row.support_tension_da_n,
                row.total_deviation_deg,
                row.pressure_da_n,
                row.friction_da_n,
                row.roller_count,
                row.unit_deviation_deg,
                row.unit_pressure_da_n,
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(_format(value))
                if isinstance(value, (int, float)):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row_index, col, item)

    def _section_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("sectionHeader")
        return label


def _button(text: str, slot) -> QPushButton:
    button = QPushButton(text)
    button.clicked.connect(slot)
    return button


def _double_editor(value: Any, unit: str) -> QDoubleSpinBox:
    editor = QDoubleSpinBox()
    editor.setRange(-1_000_000.0, 1_000_000.0)
    editor.setDecimals(4)
    editor.setValue(float(value or 0.0))
    editor.setSuffix(f" {unit}")
    return editor


def _friction_mode(value: Any) -> FrictionMode:
    if value == FrictionMode.ABSOLUTE:
        return FrictionMode.ABSOLUTE
    return FrictionMode.PERCENT


def _format(value: Any) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)
