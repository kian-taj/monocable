from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from winsif_mon.resources import default_workbook_path, resource_path


@dataclass(frozen=True, slots=True)
class SheetCell:
    row: int
    col: int
    value: Any


class WorkbookReader:
    """Small `.xls` reader used to migrate workbook-backed defaults."""

    def __init__(self, path: Path | str = "MONOFUNI.xls") -> None:
        self.path = default_workbook_path() if str(path) == "MONOFUNI.xls" else resource_path(path)
        if not self.path.exists():
            raise RuntimeError(
                f"Workbook data file not found: {self.path}. "
                "For PyInstaller builds, include MONOFUNI.xls as bundled data."
            )
        try:
            import xlrd
        except ModuleNotFoundError as exc:
            raise RuntimeError("xlrd is required to read bundled MONOFUNI.xls defaults.") from exc
        self._xlrd = xlrd
        self._book = xlrd.open_workbook(self.path, formatting_info=True)

    def sheet_names(self) -> list[str]:
        return self._book.sheet_names()

    def value(self, sheet_name: str, row: int, col: int) -> Any:
        """Return a 1-based Excel cell value."""
        sheet = self._book.sheet_by_name(sheet_name)
        return sheet.cell_value(row - 1, col - 1)

    def non_empty_rows(self, sheet_name: str, max_row: int | None = None) -> list[list[SheetCell]]:
        sheet = self._book.sheet_by_name(sheet_name)
        stop = min(sheet.nrows, max_row or sheet.nrows)
        rows: list[list[SheetCell]] = []
        for r in range(stop):
            cells: list[SheetCell] = []
            for c in range(sheet.ncols):
                value = sheet.cell_value(r, c)
                if value != "":
                    cells.append(SheetCell(r + 1, c + 1, value))
            if cells:
                rows.append(cells)
        return rows
