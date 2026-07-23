from __future__ import annotations


LIGHT_THEME = """
QMainWindow, QWidget {
    background: #f6f7f9;
    color: #18202a;
    font-size: 13px;
}
QListWidget {
    background: #ffffff;
    color: #263241;
    border: 1px solid #d8dee8;
    border-radius: 8px;
    padding: 6px;
}
QListWidget::item {
    padding: 9px 10px;
    border-radius: 6px;
}
QListWidget::item:selected {
    background: #dceafe;
    color: #113a67;
}
QLabel#pageTitle {
    font-size: 22px;
    font-weight: 700;
    color: #111827;
    margin-bottom: 8px;
}
QLabel#sectionHeader {
    background: #e8edf4;
    color: #18202a;
    font-size: 16px;
    font-weight: 700;
    padding: 7px 10px;
    border-radius: 6px;
    margin-top: 10px;
}
QLabel#fieldLabel {
    color: #273444;
    padding: 4px 6px;
}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    background: #ffffff;
    color: #111827;
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    padding: 4px 6px;
    min-height: 24px;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {
    border: 1px solid #2f80ed;
}
QRadioButton {
    background: #eef2f7;
    color: #1f2937;
    border-radius: 5px;
    padding: 4px 10px;
    min-height: 22px;
}
QPushButton {
    padding: 7px 12px;
    border-radius: 5px;
    background: #2f80ed;
    color: #ffffff;
    border: 0;
}
QPushButton:hover {
    background: #246fd3;
}
QPushButton:disabled {
    background: #aab6c5;
}
QTableWidget {
    background: #ffffff;
    alternate-background-color: #f8fafc;
    gridline-color: #d8dee8;
    border: 1px solid #d8dee8;
    border-radius: 6px;
}
QHeaderView::section {
    background: #eef2f7;
    color: #243244;
    border: 0;
    border-right: 1px solid #d8dee8;
    border-bottom: 1px solid #d8dee8;
    padding: 5px 6px;
    font-weight: 600;
}
QSplitter::handle {
    background: #d8dee8;
}
"""
