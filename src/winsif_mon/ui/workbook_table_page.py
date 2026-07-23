from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from winsif_mon.workbook import WorkbookReader


class WorkbookTablePage(QWidget):
    """Read-only workbook sheet view used for migrated output/report pages."""

    def __init__(
        self,
        title: str,
        sheet_name: str,
        max_rows: int = 120,
        max_columns: int | None = None,
        workbook_path: str = "MONOFUNI.xls",
        editable: bool = False,
        actions: list[QPushButton] | None = None,
    ) -> None:
        super().__init__()
        self.title = title
        self.sheet_name = sheet_name
        self.max_rows = max_rows
        self.max_columns = max_columns
        self.workbook_path = workbook_path
        self.editable = editable

        layout = QVBoxLayout(self)
        heading = QLabel(title)
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)

        self.status = QLabel("")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)

        if actions:
            action_row = QHBoxLayout()
            for action in actions:
                action_row.addWidget(action)
            action_row.addStretch(1)
            layout.addLayout(action_row)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        if not editable:
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table, 1)

        refresh = QPushButton("Reload workbook defaults")
        refresh.clicked.connect(self.load_sheet)
        layout.addWidget(refresh)

        self.load_sheet()

    def load_sheet(self) -> None:
        reader = WorkbookReader(self.workbook_path)
        rows = reader.non_empty_rows(self.sheet_name, max_row=self.max_rows)
        visible_columns = self._column_count(rows)
        self.table.setRowCount(len(rows))
        self.table.setColumnCount(visible_columns)
        self.table.setHorizontalHeaderLabels([_excel_column(index + 1) for index in range(visible_columns)])
        self.table.setVerticalHeaderLabels([str(row[0].row) for row in rows])

        for table_row, cells in enumerate(rows):
            row_values: dict[int, Any] = {cell.col: cell.value for cell in cells}
            for table_col in range(visible_columns):
                value = row_values.get(table_col + 1, "")
                item = QTableWidgetItem(_format_value(value))
                if isinstance(value, (int, float)):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(table_row, table_col, item)

        self.table.resizeColumnsToContents()
        self.status.setText(
            f"Showing non-empty rows from Excel sheet {self.sheet_name}. "
            "STORE05, STORE06, and STORE13 remain internal calculation stores and are intentionally hidden."
        )

    def _column_count(self, rows: list[list[Any]]) -> int:
        if self.max_columns is not None:
            return self.max_columns
        highest = 1
        for row in rows:
            for cell in row:
                highest = max(highest, cell.col)
        return highest


def _excel_column(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _format_value(value: Any) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    return str(value)
