from __future__ import annotations

from typing import Any, Callable

from PySide6.QtWidgets import QDoubleSpinBox, QLineEdit, QSpinBox, QWidget


def value_editor(value: Any, unit: str, on_change: Callable[[Any], None]) -> QWidget:
    """Create an editor that stores values by field name and shows units inside the control."""
    if _is_number(value):
        number = float(value)
        if number.is_integer() and unit == "[n]":
            editor = QSpinBox()
            editor.setRange(-1_000_000, 1_000_000)
            editor.setValue(int(number))
            _set_suffix(editor, unit)
            editor.valueChanged.connect(on_change)
            return editor

        editor = QDoubleSpinBox()
        editor.setRange(-1_000_000_000.0, 1_000_000_000.0)
        editor.setDecimals(4)
        editor.setSingleStep(_step_for(number))
        editor.setValue(number)
        _set_suffix(editor, unit)
        editor.valueChanged.connect(on_change)
        return editor

    editor = QLineEdit(str(value))
    editor.editingFinished.connect(lambda: on_change(editor.text()))
    return editor


def set_editor_value(editor: QWidget, value: Any) -> None:
    if isinstance(editor, QSpinBox) and _is_number(value):
        editor.setValue(int(float(value)))
    elif isinstance(editor, QDoubleSpinBox) and _is_number(value):
        editor.setValue(float(value))
    elif isinstance(editor, QLineEdit):
        editor.setText(_format_value(value))


def _set_suffix(editor: QSpinBox | QDoubleSpinBox, unit: str) -> None:
    if unit:
        editor.setSuffix(f" {unit}")


def _is_number(value: Any) -> bool:
    if value == "":
        return False
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _step_for(value: float) -> float:
    if abs(value) >= 1000:
        return 100.0
    if abs(value) >= 100:
        return 10.0
    if abs(value) >= 10:
        return 1.0
    return 0.1


def _format_value(value: Any) -> str:
    if value == "":
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:g}"
    return str(value)
