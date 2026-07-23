from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.inventory import build_inventory
from winsif_mon.models import CalculationMode, ProjectState
from winsif_mon.services.integrations import discover_integrations
from winsif_mon.services.solver import SolverMode, build_line_input, mode_from_calculation_mode, run_solver_setup
from winsif_mon.domain.friction import load_friction_assignment_state, validate_friction_state
from winsif_mon.domain.geometry import load_geometry_state
from winsif_mon.domain.laying_tension import load_laying_tension_state
from winsif_mon.domain.terrain import load_terrain_profile
from winsif_mon.domain.verification import (
    build_calculation_setup,
    load_custom_load_state,
    load_verification_state,
)
from winsif_mon.ui.f01_page import F01Page
from winsif_mon.ui.f02_page import F02Page
from winsif_mon.ui.geometry_pages import SpanGeometryPage, SupportGeometryPage
from winsif_mon.ui.hydraulic_dialog import HydraulicDialog
from winsif_mon.ui.line_tools_pages import FrictionAssignmentPage, LineResultsPage
from winsif_mon.ui.pages import PlaceholderPage
from winsif_mon.ui.result_pages import (
    GeneralReportPage,
    LayingTensionPage,
    MaxMinResultsPage,
    PowerSummaryPage,
    VerificationPlotsPage,
)
from winsif_mon.ui.theme import LIGHT_THEME
from winsif_mon.ui.verification_pages import CustomLoadCasesPage, LineVerificationPage
from winsif_mon.ui.workbook_table_page import WorkbookTablePage


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.state = ProjectState()
        self.geometry_state = load_geometry_state()
        self.terrain_profile = load_terrain_profile()
        self.verification_state = load_verification_state()
        self.custom_load_state = load_custom_load_state(geometry_state=self.geometry_state)
        self.friction_state = load_friction_assignment_state()
        self.last_solver_result = None
        self.setWindowTitle("WinSIF MON - PySide6 Migration")
        self.resize(1180, 760)

        root = QWidget()
        layout = QHBoxLayout(root)
        self._nav = QListWidget()
        self._nav.setFixedWidth(280)
        self._stack = QStackedWidget()
        layout.addWidget(self._nav)
        layout.addWidget(self._stack, 1)
        self.setCentralWidget(root)

        self._build_pages()
        self._nav.currentRowChanged.connect(self._stack.setCurrentIndex)
        self._nav.setCurrentRow(0)
        self._apply_style()

    def _build_pages(self) -> None:
        self._add_page(
            "Start",
            PlaceholderPage(
                "Start",
                "Workbook startup, language selection, activation compatibility, and project entry point.",
                [self._button("Run Inventory", self._show_inventory)],
            ),
        )
        self._add_page(
            "Navigation",
            PlaceholderPage(
                "Navigation",
                "Navigation hub mirroring the Excel menu sheet. Each migrated sheet is available from the left navigation.",
            ),
        )
        self._add_page("General Plant Data", F01Page())
        self._add_page("Terrain Profile", F02Page())
        support_page = SupportGeometryPage(self.geometry_state, self.terrain_profile)
        span_page = SpanGeometryPage(self.geometry_state, self.terrain_profile)
        support_page.on_relative_created = span_page.refresh_from_state
        span_page.on_absolute_created = support_page.refresh_from_state
        self._add_page("Support Geometry", support_page)
        self._add_page("Span Geometry", span_page)

        normal = self._button("Calculate Normal", lambda: self._calculation_stub(CalculationMode.NORMAL))
        anchored = self._button("Calculate Anchored", lambda: self._calculation_stub(CalculationMode.ANCHORED))
        variable = self._button(
            "Calculate Tension Variation",
            lambda: self._calculation_stub(CalculationMode.VARIABLE_TENSION),
        )
        hydraulic = self._button("Calculate Hydraulic", self._open_hydraulic)
        self._add_page(
            "Line Verification",
            LineVerificationPage(
                self.verification_state,
                self.custom_load_state,
                actions=[normal, anchored, variable, hydraulic],
            ),
        )
        self._add_page(
            "Custom Load Cases",
            CustomLoadCasesPage(self.custom_load_state, self.geometry_state),
        )
        self.friction_page = FrictionAssignmentPage(self.friction_state, self.geometry_state)
        self.line_results_page = LineResultsPage(self.verification_state)
        self._add_page("Friction Assignment", self.friction_page)
        self._add_page("Line Results", self.line_results_page)
        self._add_page("Verification Plots", VerificationPlotsPage(self.verification_state))
        self._add_page("Maximum And Minimum Results", MaxMinResultsPage())
        self._add_page("Power Summary", PowerSummaryPage())
        self._add_page("General Report", GeneralReportPage())
        self._add_page("Laying Tension", LayingTensionPage(load_laying_tension_state(self.geometry_state)))
        self._add_page(
            "Integrations",
            PlaceholderPage(
                "External Integrations",
                self._integration_text(),
            ),
        )

    def _add_page(self, name: str, page: QWidget) -> None:
        self._nav.addItem(QListWidgetItem(name))
        self._stack.addWidget(page)

    def _button(self, text: str, slot) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(slot)
        return button

    def _open_hydraulic(self) -> None:
        result = self._run_solver_preparation(SolverMode.HYDRAULIC_PRECHECK)
        if result is None:
            return
        if result.native_iterative_result and result.native_iterative_result.hydraulic_reference_values:
            self.state.hydraulic.reference_values = result.native_iterative_result.hydraulic_reference_values
            self.state.hydraulic.ring_length = result.run_parameters.rope_loop_length_m
            self.state.hydraulic.zeta = "F"
        self.state.calculation_mode = CalculationMode.HYDRAULIC_PRECHECK
        dialog = HydraulicDialog(self.state.hydraulic, self)
        if dialog.exec():
            result = self._run_solver_preparation(SolverMode.HYDRAULIC_FINAL)
            if result is None:
                return
            self.state.calculation_mode = CalculationMode.HYDRAULIC_FINAL
            QMessageBox.information(
                self,
                "Hydraulic calculation",
                "Hydraulic parameters accepted. Native hydraulic result objects were prepared; workbook stores "
                "remain the golden output until hydraulic parity tests pass.",
            )

    def _calculation_stub(self, mode: CalculationMode) -> None:
        result = self._run_solver_preparation(mode_from_calculation_mode(mode))
        if result is None:
            return
        self.state.calculation_mode = mode
        QMessageBox.information(
            self,
            "Solver preparation complete",
            (
                f"{mode.value} prepared.\n"
                f"Active cases: {len(result.active_cases)}\n"
                f"Rope loop length: {result.run_parameters.rope_loop_length_m:.3f} m\n"
                f"Total cars: {result.run_parameters.total_cars:.0f}\n"
                f"Step: {result.run_parameters.step_m:.3f} m\n\n"
                f"{result.status}"
            ),
        )

    def _validate_calculation_setup(self) -> bool:
        try:
            build_calculation_setup(self.verification_state, self.custom_load_state)
        except ValueError as exc:
            QMessageBox.warning(self, "Invalid calculation setup", str(exc))
            return False
        friction_errors = validate_friction_state(self.friction_state)
        if friction_errors:
            QMessageBox.warning(self, "Invalid friction assignments", "\n".join(friction_errors))
            return False
        return True

    def _run_solver_preparation(self, mode: SolverMode):
        if not self._validate_calculation_setup():
            return None
        line_input = build_line_input(
            terrain=self.terrain_profile,
            geometry=self.geometry_state,
            verification=self.verification_state,
            custom_loads=self.custom_load_state,
            friction_state=self.friction_state,
        )
        self.last_solver_result = run_solver_setup(mode, line_input)
        self.line_results_page.refresh_results()
        return self.last_solver_result

    def _show_inventory(self) -> None:
        inventory = build_inventory()
        summary = inventory["summary"]
        text = (
            f"Components: {summary['component_count']}\n"
            f"Procedures: {summary['procedure_count']}\n"
            f"Referenced sheets: {', '.join(summary['referenced_sheets'])}"
        )
        QMessageBox.information(self, "VBA inventory", text)

    def _integration_text(self) -> str:
        integrations = discover_integrations(Path.cwd())
        lines = []
        for integration in integrations:
            status = "available" if integration.available else "missing"
            lines.append(
                f"{integration.name}: {status}\n"
                f"Original: {integration.original_target}\n"
                f"Replacement: {integration.replacement_behavior}"
            )
        return "\n\n".join(lines)

    def _apply_style(self) -> None:
        self.setStyleSheet(LIGHT_THEME)
