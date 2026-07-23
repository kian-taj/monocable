from __future__ import annotations

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


class PlaceholderPage(QWidget):
    def __init__(self, title: str, description: str, actions: list[QPushButton] | None = None) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        heading = QLabel(title)
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)
        text = QLabel(description)
        text.setWordWrap(True)
        layout.addWidget(text)
        if actions:
            row = QHBoxLayout()
            for action in actions:
                row.addWidget(action)
            row.addStretch(1)
            layout.addLayout(row)
        layout.addStretch(1)


class SheetLikePage(QWidget):
    def __init__(self, title: str, rows: int = 24, columns: int = 9) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        heading = QLabel(title)
        heading.setObjectName("pageTitle")
        layout.addWidget(heading)
        table = QTableWidget(rows, columns)
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QAbstractItemView.SelectItems)
        for row in range(rows):
            for col in range(columns):
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                table.setItem(row, col, item)
        layout.addWidget(table)
