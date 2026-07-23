from __future__ import annotations

from typing import Any

from PySide6.QtCharts import QChart, QChartView, QLineSeries, QValueAxis
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.domain.laying_tension import LayingTensionState, load_laying_tension_state
from winsif_mon.domain.line_results import ResultFamily, active_result_cases
from winsif_mon.domain.max_min_results import MaxMinRow, load_max_min_case
from winsif_mon.domain.power_summary import PowerSummaryRow, load_power_summary_case
from winsif_mon.domain.report import build_general_report
from winsif_mon.domain.verification import VerificationState, load_verification_state
from winsif_mon.domain.verification_plots import PowerTracePoint, load_power_trace_case


class VerificationPlotsPage(QWidget):
    def __init__(self, verification_state: VerificationState | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.verification_state = verification_state or load_verification_state()
        self._case_refs = active_result_cases(self.verification_state)

        layout = QVBoxLayout(self)
        layout.addWidget(_title("Verification Plots"))
        layout.addLayout(self._selectors())
        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        layout.addWidget(self.summary)

        self.chart = QChart()
        self.chart.legend().setVisible(True)
        self.chart_view = QChartView(self.chart)
        self.chart_view.setMinimumHeight(430)
        layout.addWidget(self.chart_view)

        self.table = _table("verification_plot_trace", _trace_headers())
        layout.addWidget(self.table, 1)
        self.refresh()

    def refresh(self) -> None:
        family = _family(self.family_selector)
        reference = self.case_selector.currentData() if self.case_selector.count() else None
        case = load_power_trace_case(family, reference, self.verification_state)
        self.summary.setText(f"{case.family.label} - {case.reference.label}. Source: {case.source}.")
        self._refresh_chart(case.forward, case.reverse)
        self._fill_table(case.forward, case.reverse)

    def _selectors(self) -> QHBoxLayout:
        row = QHBoxLayout()
        self.family_selector = _family_selector(self.refresh)
        self.case_selector = QComboBox()
        for ref in self._case_refs:
            self.case_selector.addItem(ref.label, ref)
        self.case_selector.currentIndexChanged.connect(lambda _index: self.refresh())
        row.addWidget(QLabel("Result family"))
        row.addWidget(self.family_selector)
        row.addWidget(QLabel("Hypothesis and condition"))
        row.addWidget(self.case_selector, 1)
        return row

    def _refresh_chart(self, forward: list[PowerTracePoint], reverse: list[PowerTracePoint]) -> None:
        self.chart.removeAllSeries()
        for axis in self.chart.axes():
            self.chart.removeAxis(axis)
        series_defs = [
            ("Forward ascent tension", forward, "ascent_tension_da_n"),
            ("Forward descent tension", forward, "descent_tension_da_n"),
            ("Forward motive force", forward, "motive_force_da_n"),
            ("Reverse motive force", reverse, "motive_force_da_n"),
        ]
        all_points = []
        for name, points, attr in series_defs:
            series = QLineSeries()
            series.setName(name)
            for point in points:
                x = point.offset_m
                y = float(getattr(point, attr))
                series.append(x, y)
                all_points.append((x, y))
            if series.count():
                self.chart.addSeries(series)
        axis_x = QValueAxis()
        axis_x.setTitleText("Offset [m]")
        axis_y = QValueAxis()
        axis_y.setTitleText("Force / tension [daN]")
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        for series in self.chart.series():
            series.attachAxis(axis_x)
            series.attachAxis(axis_y)
        if all_points:
            axis_x.setRange(min(x for x, _y in all_points), max(x for x, _y in all_points))
            axis_y.setRange(min(y for _x, y in all_points), max(y for _x, y in all_points))

    def _fill_table(self, forward: list[PowerTracePoint], reverse: list[PowerTracePoint]) -> None:
        rows = []
        for point in forward:
            rows.append(["Forward", point.offset_m, point.ascent_tension_da_n, point.descent_tension_da_n, point.motive_force_da_n])
        for point in reverse:
            rows.append(["Reverse", point.offset_m, point.ascent_tension_da_n, point.descent_tension_da_n, point.motive_force_da_n])
        _fill(self.table, rows)


class MaxMinResultsPage(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addWidget(_title("Maximum And Minimum Results"))
        selectors = QHBoxLayout()
        self.family_selector = _family_selector(self.refresh)
        selectors.addWidget(QLabel("Result family"))
        selectors.addWidget(self.family_selector)
        selectors.addStretch(1)
        layout.addLayout(selectors)
        self.summary = QLabel("")
        layout.addWidget(self.summary)
        self.ascent_table = _table("ascent_max_min_results", _max_min_headers())
        self.descent_table = _table("descent_max_min_results", _max_min_headers())
        layout.addWidget(_section("Ascent Branch Max/Min"))
        layout.addWidget(self.ascent_table, 1)
        layout.addWidget(_section("Descent Branch Max/Min"))
        layout.addWidget(self.descent_table, 1)
        self.refresh()

    def refresh(self) -> None:
        case = load_max_min_case(_family(self.family_selector))
        self.summary.setText(f"{case.family.label}. Source: {case.source}. External DXF/viewer tools are not available.")
        _fill(self.ascent_table, [_max_min_values(row) for row in case.ascent_rows])
        _fill(self.descent_table, [_max_min_values(row) for row in case.descent_rows])


class PowerSummaryPage(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addWidget(_title("Power Summary"))
        selectors = QHBoxLayout()
        self.family_selector = _family_selector(self.refresh)
        selectors.addWidget(QLabel("Result family"))
        selectors.addWidget(self.family_selector)
        selectors.addStretch(1)
        layout.addLayout(selectors)
        self.summary = QLabel("")
        layout.addWidget(self.summary)
        self.forward_table = _table("forward_power_summary", _power_headers())
        self.reverse_table = _table("reverse_power_summary", _power_headers())
        layout.addWidget(_section("Forward Running"))
        layout.addWidget(self.forward_table, 1)
        layout.addWidget(_section("Reverse Running"))
        layout.addWidget(self.reverse_table, 1)
        self.refresh()

    def refresh(self) -> None:
        case = load_power_summary_case(_family(self.family_selector))
        self.summary.setText(f"{case.family.label}. Source: {case.source}.")
        forward = [row for row in case.rows if row.direction == "Forward"]
        reverse = [row for row in case.rows if row.direction == "Reverse"]
        _fill(self.forward_table, [_power_values(row) for row in forward])
        _fill(self.reverse_table, [_power_values(row) for row in reverse])


class GeneralReportPage(QWidget):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.addWidget(_title("General Report"))
        selectors = QHBoxLayout()
        self.family_selector = _family_selector(self.refresh)
        selectors.addWidget(QLabel("Result family"))
        selectors.addWidget(self.family_selector)
        selectors.addStretch(1)
        layout.addLayout(selectors)
        self.report = QTextEdit()
        self.report.setObjectName("general_report_text")
        self.report.setReadOnly(True)
        layout.addWidget(self.report, 1)
        self.refresh()

    def refresh(self) -> None:
        report = build_general_report(_family(self.family_selector))
        lines = [report.title, "", f"Source: {report.source}", ""]
        for section in report.sections:
            lines.append(section.title)
            for label, value in section.rows:
                lines.append(f"  {label}: {value}")
            lines.append("")
        self.report.setPlainText("\n".join(lines))


class LayingTensionPage(QWidget):
    def __init__(self, state: LayingTensionState | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state or load_laying_tension_state()
        layout = QVBoxLayout(self)
        layout.addWidget(_title("Laying Tension"))
        self.summary = QLabel(
            f"Ambient temperature: {self.state.ambient_temperature_c:g} deg C. "
            f"Reference tension: {self.state.reference_tension_da_n:g} daN."
        )
        layout.addWidget(self.summary)
        self.chart = QChart()
        self.chart.legend().setVisible(True)
        self.chart_view = QChartView(self.chart)
        self.chart_view.setMinimumHeight(360)
        layout.addWidget(self.chart_view)
        self.ascent_table = _table("ascent_laying_tension", _laying_headers())
        self.descent_table = _table("descent_laying_tension", _laying_headers())
        layout.addWidget(_section("Ascent Laying Spans"))
        layout.addWidget(self.ascent_table, 1)
        layout.addWidget(_section("Descent Laying Spans"))
        layout.addWidget(self.descent_table, 1)
        self.refresh()

    def refresh(self) -> None:
        self._refresh_chart()
        _fill(self.ascent_table, [[r.span_start, r.span_end, r.horizontal_distance_m, r.height_difference_m, r.laying_tension_da_n] for r in self.state.ascent_rows])
        _fill(self.descent_table, [[r.span_start, r.span_end, r.horizontal_distance_m, r.height_difference_m, r.laying_tension_da_n] for r in self.state.descent_rows])

    def _refresh_chart(self) -> None:
        self.chart.removeAllSeries()
        for axis in self.chart.axes():
            self.chart.removeAxis(axis)
        first = QLineSeries()
        first.setName("First span")
        last = QLineSeries()
        last.setName("Last span")
        for temperature, first_value, last_value in self.state.temperature_curve:
            first.append(temperature, first_value)
            last.append(temperature, last_value)
        self.chart.addSeries(first)
        self.chart.addSeries(last)
        axis_x = QValueAxis()
        axis_x.setTitleText("Temperature [deg C]")
        axis_y = QValueAxis()
        axis_y.setTitleText("Laying tension [daN]")
        self.chart.addAxis(axis_x, Qt.AlignBottom)
        self.chart.addAxis(axis_y, Qt.AlignLeft)
        first.attachAxis(axis_x)
        first.attachAxis(axis_y)
        last.attachAxis(axis_x)
        last.attachAxis(axis_y)
        axis_x.setRange(-20, 30)
        values = [value for _temperature, first_value, last_value in self.state.temperature_curve for value in (first_value, last_value)]
        axis_y.setRange(min(values), max(values))


def _family_selector(slot) -> QComboBox:
    selector = QComboBox()
    for family in ResultFamily:
        selector.addItem(family.label, family)
    selector.currentIndexChanged.connect(lambda _index: slot())
    return selector


def _family(selector: QComboBox) -> ResultFamily:
    value = selector.currentData()
    return value if isinstance(value, ResultFamily) else ResultFamily.NORMAL


def _title(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("pageTitle")
    return label


def _section(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("sectionHeader")
    return label


def _table(name: str, headers: list[str]) -> QTableWidget:
    table = QTableWidget(0, len(headers))
    table.setObjectName(name)
    table.setAlternatingRowColors(True)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setEditTriggers(QAbstractItemView.NoEditTriggers)
    table.setHorizontalHeaderLabels(headers)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
    table.horizontalHeader().setStretchLastSection(True)
    return table


def _fill(table: QTableWidget, rows: list[list[Any]]) -> None:
    table.setRowCount(len(rows))
    for row_index, row_values in enumerate(rows):
        for col, value in enumerate(row_values):
            item = QTableWidgetItem(_format(value))
            if isinstance(value, (int, float)):
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            table.setItem(row_index, col, item)


def _trace_headers() -> list[str]:
    return ["Direction", "Offset [m]", "Ascent tension [daN]", "Descent tension [daN]", "Motive force [daN]"]


def _max_min_headers() -> list[str]:
    return [
        "Span start",
        "Span end",
        "Span state",
        "Tension [daN]",
        "Sag [m]",
        "Valley angle [deg]",
        "Mountain angle [deg]",
        "Support",
        "Support state",
        "Support tension [daN]",
        "Deviation [deg]",
        "Pressure [daN]",
        "Friction [daN]",
        "Rollers [n]",
        "Unit deviation [deg]",
        "Unit pressure [daN]",
        "Test",
    ]


def _max_min_values(row: MaxMinRow) -> list[Any]:
    return [
        row.span_start,
        row.span_end,
        row.span_state,
        row.tension_da_n,
        row.sag_m,
        row.valley_angle_deg,
        row.mountain_angle_deg,
        row.support_code,
        row.support_state,
        row.support_tension_da_n,
        row.deviation_deg,
        row.pressure_da_n,
        row.friction_da_n,
        row.roller_count,
        row.unit_deviation_deg,
        row.unit_pressure_da_n,
        row.test,
    ]


def _power_headers() -> list[str]:
    return [
        "Case",
        "Load hypothesis",
        "Plant condition",
        "Mean force [daN]",
        "Max force [daN]",
        "Motor inertia [daN]",
        "Motor stress [daN]",
        "Efficiency [n]",
        "Power [kW]",
        "Slip [n]",
        "Length [m]",
        "Total tension [daN]",
    ]


def _power_values(row: PowerSummaryRow) -> list[Any]:
    return [
        row.case_index,
        row.load_hypothesis,
        row.plant_condition,
        row.mean_force_da_n,
        row.max_force_da_n,
        row.motor_inertia_da_n,
        row.motor_stress_da_n,
        row.efficiency,
        row.power_kw,
        row.slip,
        row.length_m,
        row.total_tension_da_n,
    ]


def _laying_headers() -> list[str]:
    return ["Span start", "Span end", "Horizontal distance [m]", "Height difference [m]", "Laying tension [daN]"]


def _format(value: Any) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        return f"{value:.4f}".rstrip("0").rstrip(".")
    return str(value)
