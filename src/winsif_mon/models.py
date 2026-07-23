from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class CalculationMode(str, Enum):
    NORMAL = "normal"
    HYDRAULIC_PRECHECK = "hydraulic_precheck"
    HYDRAULIC_FINAL = "hydraulic_final"
    ANCHORED = "anchored"
    VARIABLE_TENSION = "variable_tension"


@dataclass(slots=True)
class ProjectState:
    """Application state that replaces workbook-global variables incrementally."""

    language: str = "it"
    workbook_path: str = "MONOFUNI.xls"
    activation_enabled: bool = False
    calculation_mode: CalculationMode | None = None
    hydraulic: "HydraulicState" = field(default_factory=lambda: HydraulicState())


@dataclass(slots=True)
class HydraulicState:
    """Subset of the original `corsaz` hydraulic dialog state."""

    reference_values: list[list[float]] = field(
        default_factory=lambda: [[0.0 for _ in range(3)] for _ in range(7)]
    )
    thermal_delta: float = 30.0
    extra_stroke_1: float = 0.3
    extra_stroke_2: float = 0.1
    ring_length: float = 0.0
    zeta: str = ""

    @property
    def thermal_stroke(self) -> float:
        return 0.000012 * self.ring_length / 2.0 * self.thermal_delta


@dataclass(frozen=True, slots=True)
class ExternalIntegration:
    name: str
    original_target: str
    replacement_behavior: str
    available: bool = False
