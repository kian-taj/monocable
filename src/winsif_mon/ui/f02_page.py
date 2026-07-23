from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtCharts import QChart, QChartView, QLineSeries, QValueAxis
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.terrain import TerrainPoint, load_terrain_profile


class F02Page(QWidget):
    """Workbook-backed terrain profile page (`F02`)."""

    HEADERS = [
        "Stake No.",
        "Code",
        "Distance (m)",
        "Elevation (m)",
        "Left slope (%)",
        "Right slope (%)",
    ]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.profile = load_terrain_profile()

        layout = QVBoxLayout(self)
        header = QLabel("Terrain Profile")
        header.setObjectName("pageTitle")
        layout.addWidget(header)

        toolbar = QHBoxLayout()
        toolbar.addWidget(self._button("Insert row", self._insert_row))
        toolbar.addWidget(self._button("Delete row", self._delete_row))
        toolbar.addWidget(self._button("Clear selected branch", self._clear_branch))
        toolbar.addWidget(self._button("Refresh plot", self._refresh_chart))
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        self.chart = QChart()
        self.chart.setTitle("Ground profile")
        self.chart.legend().setVisible(True)
        self.chart_view = QChartView(self.chart)
        self.chart_view.setMinimumHeight(460)
        layout.addWidget(self.chart_view)

        splitter = QSplitter(Qt.Horizontal)
        self.ascent_table = self._make_table("ASCENT BRANCH", self.profile.ascent)
        self.descent_table = self._make_table("DESCENT BRANCH", self.profile.descent)
        splitter.addWidget(self.ascent_table)
        splitter.addWidget(self.descent_table)
        splitter.setSizes([1, 1])
        layout.addWidget(splitter, 1)
        self._refresh_chart()

    def _make_table(self, title: str, rows: list[TerrainPoint]) -> QTableWidget:
        table = QTableWidget(max(len(rows), 24), len(self.HEADERS))
        table.setObjectName(title.lower().replace(" ", "_"))
        table.setHorizontalHeaderLabels(self.HEADERS)
        table.setAlternatingRowColors(True)
        table.setProperty("branchTitle", title)
        for index, point in enumerate(rows):
            self._set_point(table, index, point)
        table.resizeColumnsToContents()
        return table

    def _set_point(self, table: QTableWidget, row: int, point: TerrainPoint) -> None:
        values = [
            point.stake_number,
            point.stake_code,
            point.progressive_distance_m,
            point.ground_elevation_m,
            point.left_slope_percent,
            point.right_slope_percent,
        ]
        for col, value in enumerate(values):
            item = QTableWidgetItem(_format(value))
            if col != 1:
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            table.setItem(row, col, item)

    def _button(self, text: str, slot) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _active_table(self) -> QTableWidget:
        focus = self.focusWidget()
        while focus is not None:
            if focus is self.ascent_table or focus is self.descent_table:
                return focus
            focus = focus.parentWidget()
        return self.ascent_table

    def _insert_row(self) -> None:
        table = self._active_table()
        row = max(table.currentRow(), 0)
        table.insertRow(row)
        table.setItem(row, 0, QTableWidgetItem(str(row + 1)))

    def _delete_row(self) -> None:
        table = self._active_table()
        row = table.currentRow()
        if row >= 0:
            table.removeRow(row)

    def _clear_branch(self) -> None:
        table = self._active_table()
        table.setRowCount(24)
        for row in range(table.rowCount()):
            for col in range(table.columnCount()):
                table.setItem(row, col, QTableWidgetItem(""))
        self._refresh_chart()

    def _refresh_chart(self) -> None:
        self.chart.removeAllSeries()
        for axis in self.chart.axes():
            self.chart.removeAxis(axis)
        ascent = self._series_from_table(self.ascent_table, "Ascent branch")
        descent = self._series_from_table(self.descent_table, "Descent branch")
        for series in (ascent, descent):
            if series.count():
                self.chart.addSeries(series)

        axis_x = QValueAxis()
        axis_x.setTitleText("Distance (m)")
        axis_y = QValueAxis()
        axis_y.setTitleText("Elevation (m)")
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        for series in self.chart.series():
            series.attachAxis(axis_x)
            series.attachAxis(axis_y)
        points = [point for series in (ascent, descent) for point in series.points()]
        if points:
            min_x = min(point.x() for point in points)
            max_x = max(point.x() for point in points)
            min_y = min(point.y() for point in points)
            max_y = max(point.y() for point in points)
            axis_x.setRange(min_x, max_x)
            axis_y.setRange(min_y, max_y)

    def _series_from_table(self, table: QTableWidget, name: str) -> QLineSeries:
        series = QLineSeries()
        series.setName(name)
        for row in range(table.rowCount()):
            distance = _number_item(table.item(row, 2))
            elevation = _number_item(table.item(row, 3))
            if distance is None or elevation is None:
                continue
            series.append(distance, elevation)
        return series


def _format(value) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    return str(value)


def _number_item(item: QTableWidgetItem | None) -> float | None:
    if item is None:
        return None
    text = item.text().strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None
