from __future__ import annotations

from dataclasses import astuple
from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDoubleSpinBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.geometry import (
    FoundationDrawingRow,
    GeometryState,
    SpanRow,
    SupportRow,
    convert_spans_to_supports,
    convert_supports_to_spans,
    load_geometry_state,
    update_support_ground_from_terrain,
    update_support_rope_elevation,
    update_support_tower_height,
)
from winsif_mon.domain.terrain import TerrainProfile, load_terrain_profile


class SupportGeometryPage(QWidget):
    SUPPORT_HEADERS = [
        "Line",
        "Tower",
        "Rope dist. (m)",
        "Ground elev. (m)",
        "Tower ht. (m)",
        "Rope elev. (m)",
        "Rollers",
    ]
    FOUNDATION_HEADERS = [
        "Rope-base ht. (m)",
        "Incl. (deg)",
        "Protrusion (m)",
        "Plinth ht. (m)",
    ]

    def __init__(
        self,
        state: GeometryState | None = None,
        terrain_profile: TerrainProfile | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.state = state or load_geometry_state()
        self.terrain_profile = terrain_profile or load_terrain_profile()
        self.on_relative_created: Callable[[], None] | None = None
        self._loading = False

        layout = QVBoxLayout(self)
        heading = QLabel("Support Geometry")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self._button("Insert Row", self._insert_row))
        toolbar.addWidget(self._button("Delete Row", self._delete_row))
        toolbar.addWidget(self._button("Create Relative Values", self._create_relative_values))
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        splitter = QSplitter(Qt.Vertical)
        branch_splitter = QSplitter(Qt.Horizontal)
        self.ascent_table = self._support_table("Ascent Branch Supports")
        self.descent_table = self._support_table("Descent Branch Supports")
        branch_splitter.addWidget(_table_panel("Ascent Branch Supports", self.ascent_table))
        branch_splitter.addWidget(_table_panel("Descent Branch Supports", self.descent_table))
        branch_splitter.setSizes([1, 1])

        self.foundation_table = self._foundation_table()
        splitter.addWidget(branch_splitter)
        splitter.addWidget(_table_panel("Support And Foundation Drawing Dimensions", self.foundation_table))
        splitter.setSizes([640, 220])
        layout.addWidget(splitter, 1)

        self.refresh_from_state()

    def refresh_from_state(self) -> None:
        self._loading = True
        _fill_table(self.ascent_table, [astuple(row) for row in self.state.ascent_supports])
        _fill_table(self.descent_table, [astuple(row) for row in self.state.descent_supports])
        _fill_table(self.foundation_table, [astuple(row) for row in self.state.foundation_rows])
        self._loading = False

    def _support_table(self, title: str) -> QTableWidget:
        table = QTableWidget(80, len(self.SUPPORT_HEADERS))
        table.setObjectName(title.lower().replace(" ", "_"))
        table.setHorizontalHeaderLabels(self.SUPPORT_HEADERS)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.cellChanged.connect(lambda row, col, t=table: self._support_cell_changed(t, row, col))
        widths = [54, 72, 104, 112, 96, 100, 66]
        for col, width in enumerate(widths):
            table.setColumnWidth(col, width)
        return table

    def _foundation_table(self) -> QTableWidget:
        table = QTableWidget(80, len(self.FOUNDATION_HEADERS))
        table.setObjectName("support_foundation_drawing_dimensions")
        table.setHorizontalHeaderLabels(self.FOUNDATION_HEADERS)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.horizontalHeader().setStretchLastSection(False)
        for col, width in enumerate([145, 110, 125, 115]):
            table.setColumnWidth(col, width)
        return table

    def _button(self, text: str, slot) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _active_table(self) -> QTableWidget:
        focus = self.focusWidget()
        while focus is not None:
            if focus in {self.ascent_table, self.descent_table, self.foundation_table}:
                return focus
            focus = focus.parentWidget()
        return self.ascent_table

    def _insert_row(self) -> None:
        table = self._active_table()
        row = max(table.currentRow(), 0)
        table.insertRow(row)
        if table is not self.foundation_table:
            table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self._save_to_state()

    def _delete_row(self) -> None:
        table = self._active_table()
        if table.currentRow() >= 0:
            table.removeRow(table.currentRow())
            self._save_to_state()

    def _create_relative_values(self) -> None:
        self._save_to_state()
        self.state.ascent_spans = convert_supports_to_spans(self.state.ascent_supports)
        self.state.descent_spans = convert_supports_to_spans(self.state.descent_supports)
        if self.on_relative_created:
            self.on_relative_created()
        QMessageBox.information(self, "Relative spans", "Span Geometry was updated from Support Geometry.")

    def _support_cell_changed(self, table: QTableWidget, row: int, col: int) -> None:
        if self._loading:
            return
        self._save_to_state()
        rows = self.state.ascent_supports if table is self.ascent_table else self.state.descent_supports
        if row >= len(rows):
            return
        branch = "ascent" if table is self.ascent_table else "descent"
        if col in {2, 3}:
            rows[row] = update_support_ground_from_terrain(rows[row], self.terrain_profile, branch)
        elif col == 4:
            rows[row] = update_support_rope_elevation(rows[row])
        elif col == 5:
            rows[row] = update_support_tower_height(rows[row])
        self.refresh_from_state()

    def _save_to_state(self) -> None:
        self.state.ascent_supports = _read_support_rows(self.ascent_table)
        self.state.descent_supports = _read_support_rows(self.descent_table)
        self.state.foundation_rows = _read_foundation_rows(self.foundation_table)


class SpanGeometryPage(QWidget):
    SPAN_HEADERS = ["Span No.", "Span code", "Horizontal distance (m)", "Height difference (m)"]

    def __init__(
        self,
        state: GeometryState | None = None,
        terrain_profile: TerrainProfile | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.state = state or load_geometry_state()
        self.terrain_profile = terrain_profile or load_terrain_profile()
        self.on_absolute_created: Callable[[], None] | None = None

        layout = QVBoxLayout(self)
        heading = QLabel("Span Geometry")
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        top = QHBoxLayout()
        top.addWidget(QLabel("First progressive distance"))
        self.first_distance = _meter_spin(self.state.first_progressive_distance_m)
        top.addWidget(self.first_distance)
        top.addWidget(QLabel("First height: rope bottom"))
        self.first_height = _meter_spin(self.state.first_rope_height_m)
        top.addWidget(self.first_height)
        top.addStretch(1)
        layout.addLayout(top)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self._button("Insert Row", self._insert_row))
        toolbar.addWidget(self._button("Delete Row", self._delete_row))
        toolbar.addWidget(self._button("Find Absolute Values", self._find_absolute_values))
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        splitter = QSplitter(Qt.Horizontal)
        self.ascent_table = self._span_table("Ascent Branch Spans")
        self.descent_table = self._span_table("Descent Branch Spans")
        splitter.addWidget(_table_panel("Ascent Branch Spans", self.ascent_table))
        splitter.addWidget(_table_panel("Descent Branch Spans", self.descent_table))
        splitter.setSizes([1, 1])
        layout.addWidget(splitter, 1)

        self.refresh_from_state()

    def refresh_from_state(self) -> None:
        self.first_distance.setValue(self.state.first_progressive_distance_m)
        self.first_height.setValue(self.state.first_rope_height_m)
        _fill_table(self.ascent_table, [astuple(row) for row in self.state.ascent_spans])
        _fill_table(self.descent_table, [astuple(row) for row in self.state.descent_spans])

    def _span_table(self, title: str) -> QTableWidget:
        table = QTableWidget(80, len(self.SPAN_HEADERS))
        table.setObjectName(title.lower().replace(" ", "_"))
        table.setHorizontalHeaderLabels(self.SPAN_HEADERS)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setColumnWidth(0, 80)
        table.setColumnWidth(1, 130)
        table.setColumnWidth(2, 170)
        table.setColumnWidth(3, 160)
        return table

    def _button(self, text: str, slot) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _active_table(self) -> QTableWidget:
        focus = self.focusWidget()
        while focus is not None:
            if focus in {self.ascent_table, self.descent_table}:
                return focus
            focus = focus.parentWidget()
        return self.ascent_table

    def _insert_row(self) -> None:
        table = self._active_table()
        row = max(table.currentRow(), 0)
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self._save_to_state()

    def _delete_row(self) -> None:
        table = self._active_table()
        if table.currentRow() >= 0:
            table.removeRow(table.currentRow())
            self._save_to_state()

    def _find_absolute_values(self) -> None:
        self._save_to_state()
        self.state.ascent_supports = convert_spans_to_supports(
            self.state.ascent_spans,
            self.state.first_progressive_distance_m,
            self.state.first_rope_height_m,
            self.terrain_profile,
            "ascent",
        )
        self.state.descent_supports = convert_spans_to_supports(
            self.state.descent_spans,
            self.state.first_progressive_distance_m,
            self.state.first_rope_height_m,
            self.terrain_profile,
            "descent",
        )
        if self.on_absolute_created:
            self.on_absolute_created()
        QMessageBox.information(self, "Absolute supports", "Support Geometry was updated from Span Geometry.")

    def _save_to_state(self) -> None:
        self.state.first_progressive_distance_m = self.first_distance.value()
        self.state.first_rope_height_m = self.first_height.value()
        self.state.ascent_spans = _read_span_rows(self.ascent_table)
        self.state.descent_spans = _read_span_rows(self.descent_table)


def _table_panel(title: str, table: QTableWidget) -> QWidget:
    panel = QWidget()
    layout = QVBoxLayout(panel)
    layout.setContentsMargins(0, 0, 0, 0)
    label = QLabel(title)
    label.setObjectName("sectionHeader")
    layout.addWidget(label)
    layout.addWidget(table)
    return panel


def _fill_table(table: QTableWidget, rows: list[tuple]) -> None:
    table.blockSignals(True)
    table.setRowCount(max(len(rows), 24))
    for row_index in range(table.rowCount()):
        values = rows[row_index] if row_index < len(rows) else [""] * table.columnCount()
        for col_index in range(table.columnCount()):
            value = values[col_index] if col_index < len(values) else ""
            item = QTableWidgetItem(_format(value))
            if _is_number(value):
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            table.setItem(row_index, col_index, item)
    table.blockSignals(False)


def _read_support_rows(table: QTableWidget) -> list[SupportRow]:
    rows: list[SupportRow] = []
    for row in range(table.rowCount()):
        values = [_cell_text(table, row, col) for col in range(table.columnCount())]
        if not any(values[1:]):
            continue
        rows.append(
            SupportRow(
                line_number=int(_number(values[0]) or len(rows) + 1),
                code=values[1],
                rope_distance_m=_value(values[2]),
                ground_elevation_m=_value(values[3]),
                tower_height_m=_value(values[4]),
                rope_elevation_m=_value(values[5]),
                roller_quantity=_value(values[6]),
            )
        )
    return rows


def _read_foundation_rows(table: QTableWidget) -> list[FoundationDrawingRow]:
    rows: list[FoundationDrawingRow] = []
    for row in range(table.rowCount()):
        values = [_value(_cell_text(table, row, col)) for col in range(table.columnCount())]
        if not any(value != "" for value in values):
            continue
        rows.append(FoundationDrawingRow(*values))
    return rows


def _read_span_rows(table: QTableWidget) -> list[SpanRow]:
    rows: list[SpanRow] = []
    for row in range(table.rowCount()):
        values = [_cell_text(table, row, col) for col in range(table.columnCount())]
        if not any(values[1:]):
            continue
        rows.append(
            SpanRow(
                span_number=int(_number(values[0]) or len(rows) + 1),
                code=values[1],
                horizontal_distance_m=_value(values[2]),
                height_difference_m=_value(values[3]),
            )
        )
    return rows


def _meter_spin(value: float) -> QDoubleSpinBox:
    editor = QDoubleSpinBox()
    editor.setRange(-1_000_000.0, 1_000_000.0)
    editor.setDecimals(4)
    editor.setValue(value)
    editor.setSuffix(" m")
    return editor


def _cell_text(table: QTableWidget, row: int, col: int) -> str:
    item = table.item(row, col)
    return "" if item is None else item.text().strip()


def _value(text: str) -> float | str:
    if text == "":
        return ""
    if _is_number(text):
        return float(text)
    return text


def _number(value) -> float:
    if value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _is_number(value) -> bool:
    if value == "":
        return False
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _format(value) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    return str(value)
