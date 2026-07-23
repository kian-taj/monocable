from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from winsif_mon.workbook import WorkbookReader


@dataclass(frozen=True, slots=True)
class WorkbookLabels:
    load_hypotheses: list[str]
    plant_conditions: list[str]
    branch_names: tuple[str, str]
    hydraulic_labels: list[str]


def load_labels(workbook_path: Path | str = "MONOFUNI.xls", language: str = "en") -> WorkbookLabels:
    reader = WorkbookReader(workbook_path)
    offset = 59 if language.lower().startswith("it") else 0
    return WorkbookLabels(
        load_hypotheses=[str(reader.value("F13", row + offset, 1)).strip() for row in range(1, 16)],
        plant_conditions=[str(reader.value("F13", row + offset, 5)).strip() for row in range(1, 7)],
        branch_names=(
            str(reader.value("F13", 1 + offset, 8)).strip() or "Ascent branch",
            str(reader.value("F13", 2 + offset, 8)).strip() or "Descent branch",
        ),
        hydraulic_labels=[str(reader.value("F13", row + offset, 1)).strip() for row in range(30, 47)],
    )
