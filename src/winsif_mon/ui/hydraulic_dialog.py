from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.models import HydraulicState
from winsif_mon.services.hydraulic import calculate_hydraulic, format_vba_number


class HydraulicDialog(QDialog):
    """PySide6 version of the exported `idraulic_input` UserForm."""

    def __init__(self, state: HydraulicState, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.state = state
        self.setWindowTitle(
            "Verifica della linea per guasto al tenditore idraulico: parametri di riferimento"
        )
        self.setMinimumSize(760, 520)

        self._result_labels: dict[str, QLabel] = {}
        self._thermal = QLineEdit(format_vba_number(state.thermal_delta))
        self._extra_1 = QLineEdit(format_vba_number(state.extra_stroke_1))
        self._extra_2 = QLineEdit(format_vba_number(state.extra_stroke_2))
        self._cylinder = QLineEdit()
        self._cylinder.setReadOnly(True)
        self._opt_nominal = QRadioButton("Calcolo corse con riferimento alla fune nuda")
        self._opt_unloaded = QRadioButton("Calcolo corse con riferimento alla fune scarica")
        self._opt_nominal.setChecked(True)

        layout = QVBoxLayout(self)
        layout.addWidget(self._build_reference_group())
        layout.addWidget(self._build_input_group())
        layout.addWidget(self._build_total_group())

        buttons = QDialogButtonBox()
        start = QPushButton("Esegui la verifica")
        stop = QPushButton("Esci")
        buttons.addButton(start, QDialogButtonBox.AcceptRole)
        buttons.addButton(stop, QDialogButtonBox.RejectRole)
        start.clicked.connect(self.accept)
        stop.clicked.connect(self.reject)
        layout.addWidget(buttons)

        for editor in (self._thermal, self._extra_1, self._extra_2):
            editor.editingFinished.connect(self._recalculate)
            editor.returnPressed.connect(self._recalculate)
        self._recalculate()

    def _build_reference_group(self) -> QGroupBox:
        group = QGroupBox("Corse di riferimento")
        grid = QGridLayout(group)
        headers = ["", "Nominal", "+ 10%", "- 10%"]
        for col, text in enumerate(headers):
            label = QLabel(text)
            label.setAlignment(Qt.AlignCenter)
            grid.addWidget(label, 0, col)

        row_names = [
            "1 - Corsa di riferimento a fune nuda",
            "2 - Variazione per carico",
            "3 - Variazione per temperatura",
            "4 - Variazione per tensione",
            "5 - Riferimento tensioni",
            "6 - Riferimento sviluppi",
            "7 - Riferimento finale",
        ]
        for row, name in enumerate(row_names, start=1):
            grid.addWidget(QLabel(name), row, 0)
            for col in range(3):
                value = self.state.reference_values[row - 1][col]
                label = QLabel(format_vba_number(value))
                label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                grid.addWidget(label, row, col + 1)
        return group

    def _build_input_group(self) -> QGroupBox:
        group = QGroupBox("Parametri")
        grid = QGridLayout(group)
        grid.addWidget(QLabel("Massima escursione della temperatura C"), 0, 0)
        grid.addWidget(self._thermal, 0, 1)
        grid.addWidget(QLabel("Corsa termica"), 0, 2)
        self._result_labels["thermal"] = QLabel()
        grid.addWidget(self._result_labels["thermal"], 0, 3)

        grid.addWidget(QLabel("Incremento corsa 1"), 1, 0)
        grid.addWidget(self._extra_1, 1, 1)
        grid.addWidget(QLabel("Incremento corsa 2"), 2, 0)
        grid.addWidget(self._extra_2, 2, 1)
        grid.addWidget(QLabel("Corsa cilindro"), 3, 0)
        grid.addWidget(self._cylinder, 3, 1)

        options = QHBoxLayout()
        options.addWidget(self._opt_nominal)
        options.addWidget(self._opt_unloaded)
        grid.addLayout(options, 4, 0, 1, 4)
        return group

    def _build_total_group(self) -> QGroupBox:
        group = QGroupBox("Totali")
        grid = QGridLayout(group)
        for row, key, text in (
            (0, "nominal", "Nominal"),
            (1, "plus", "+ 10%"),
            (2, "minus", "- 10%"),
            (3, "max", "Massima corsa"),
        ):
            grid.addWidget(QLabel(text), row, 0)
            label = QLabel()
            label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self._result_labels[key] = label
            grid.addWidget(label, row, 1)
        return group

    def _recalculate(self) -> None:
        self.state.thermal_delta = _parse_float(self._thermal.text(), self.state.thermal_delta)
        self.state.extra_stroke_1 = _parse_float(self._extra_1.text(), self.state.extra_stroke_1)
        self.state.extra_stroke_2 = _parse_float(self._extra_2.text(), self.state.extra_stroke_2)
        result = calculate_hydraulic(self.state)
        self._result_labels["thermal"].setText(format_vba_number(result.thermal_stroke))
        self._result_labels["nominal"].setText(format_vba_number(result.total_nominal))
        self._result_labels["plus"].setText(format_vba_number(result.total_plus_10))
        self._result_labels["minus"].setText(format_vba_number(result.total_minus_10))
        self._result_labels["max"].setText(format_vba_number(result.max_stroke))
        self._cylinder.setText(format_vba_number(result.cylinder_stroke))


def _parse_float(text: str, fallback: float) -> float:
    try:
        return float(text.replace(",", "."))
    except ValueError:
        return fallback
