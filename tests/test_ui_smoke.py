import os

from PySide6.QtWidgets import QApplication, QDoubleSpinBox, QSpinBox

from winsif_mon.ui.geometry_pages import SpanGeometryPage, SupportGeometryPage
from winsif_mon.ui.line_tools_pages import FrictionAssignmentPage, LineResultsPage
from winsif_mon.ui.main_window import MainWindow
from winsif_mon.ui.result_pages import (
    GeneralReportPage,
    LayingTensionPage,
    MaxMinResultsPage,
    PowerSummaryPage,
    VerificationPlotsPage,
)
from winsif_mon.ui.verification_pages import CustomLoadCasesPage, LineVerificationPage
from winsif_mon.ui.workbook_table_page import WorkbookTablePage


def test_main_window_uses_structured_geometry_pages_and_unit_suffixes():
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    support = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), SupportGeometryPage)
    )
    span = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), SpanGeometryPage)
    )
    working_speed = window.findChild(QDoubleSpinBox, "working_speed") or window.findChild(QSpinBox, "working_speed")
    line_verification = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), LineVerificationPage)
    )
    custom_loads = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), CustomLoadCasesPage)
    )
    friction = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), FrictionAssignmentPage)
    )
    line_results = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), LineResultsPage)
    )
    verification_plots = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), VerificationPlotsPage)
    )
    max_min = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), MaxMinResultsPage)
    )
    power_summary = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), PowerSummaryPage)
    )
    report = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), GeneralReportPage)
    )
    laying = next(
        window._stack.widget(index)
        for index in range(window._stack.count())
        if isinstance(window._stack.widget(index), LayingTensionPage)
    )

    assert support.ascent_table.columnCount() == 7
    assert support.descent_table.columnCount() == 7
    assert support.foundation_table.columnCount() == 4
    assert span.ascent_table.columnCount() == 4
    assert span.descent_table.columnCount() == 4
    assert working_speed is not None
    assert working_speed.prefix() == ""
    assert working_speed.suffix() == " [m/s]"
    assert line_verification.matrix_table.rowCount() == 15
    assert line_verification.matrix_table.columnCount() == 7
    assert custom_loads.table.columnCount() == 22
    assert friction.ascent_table.columnCount() == 4
    assert friction.descent_table.columnCount() == 4
    assert friction.steady_default.suffix() == " %"
    assert line_results.ascent_table.columnCount() == 15
    assert line_results.descent_table.columnCount() == 15
    assert line_results.ascent_table.item(0, 0).text() == "Valle"
    assert verification_plots.table.columnCount() == 5
    assert verification_plots.chart.series()
    assert max_min.ascent_table.columnCount() == 17
    assert power_summary.forward_table.columnCount() == 12
    assert "Source:" in report.report.toPlainText()
    assert laying.ascent_table.columnCount() == 5
    assert laying.chart.series()
    assert not any(
        isinstance(window._stack.widget(index), WorkbookTablePage)
        and window._nav.item(index).text()
        in {
            "Verification Plots",
            "Maximum And Minimum Results",
            "Power Summary",
            "General Report",
            "Texts And Translations",
            "Laying Tension",
        }
        for index in range(window._stack.count())
    )
    before = line_verification.matrix_table.item(0, 1).text()
    line_verification.cycle_cell(0, 1)
    after = line_verification.matrix_table.item(0, 1).text()
    assert before != after

    app.processEvents()
