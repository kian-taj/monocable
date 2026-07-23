from __future__ import annotations

import math
from collections.abc import Iterable

from PySide6.QtCharts import QValueAxis


def padded_range(values: Iterable[float], padding_ratio: float = 0.05) -> tuple[float, float]:
    numbers = [float(value) for value in values]
    if not numbers:
        return 0.0, 1.0
    minimum = min(numbers)
    maximum = max(numbers)
    if math.isclose(minimum, maximum):
        padding = max(abs(minimum) * padding_ratio, 1.0)
    else:
        padding = (maximum - minimum) * padding_ratio
    return minimum - padding, maximum + padding


def apply_nice_y_axis(axis: QValueAxis, values: Iterable[float], tick_count: int = 6) -> None:
    minimum, maximum = padded_range(values, padding_ratio=0.08)
    nice_minimum, nice_maximum, step = nice_range(minimum, maximum, tick_count)
    axis.setRange(nice_minimum, nice_maximum)
    axis.setTickCount(max(tick_count, 2))
    axis.setLabelFormat(_label_format(step))


def nice_range(minimum: float, maximum: float, tick_count: int = 6) -> tuple[float, float, float]:
    if tick_count < 2:
        tick_count = 2
    if math.isclose(minimum, maximum):
        minimum -= 1.0
        maximum += 1.0
    raw_step = (maximum - minimum) / (tick_count - 1)
    step = _nice_number(raw_step)
    nice_minimum = math.floor(minimum / step) * step
    nice_maximum = math.ceil(maximum / step) * step
    return nice_minimum, nice_maximum, step


def _nice_number(value: float) -> float:
    if value <= 0:
        return 1.0
    exponent = math.floor(math.log10(value))
    fraction = value / 10**exponent
    if fraction <= 1:
        nice_fraction = 1
    elif fraction <= 2:
        nice_fraction = 2
    elif fraction <= 5:
        nice_fraction = 5
    else:
        nice_fraction = 10
    return nice_fraction * 10**exponent


def _label_format(step: float) -> str:
    if step >= 1:
        return "%.0f"
    decimals = min(max(int(math.ceil(-math.log10(step))), 1), 4)
    return f"%.{decimals}f"
