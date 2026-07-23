from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_WORKBOOK_NAME = "MONOFUNI.xls"
DEFAULT_VBA_EXPORT_NAME = "VBA_Export_20260720_120646"


def bundled_root() -> Path:
    """Return the runtime data root for source runs and PyInstaller bundles."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return PROJECT_ROOT


def resource_path(relative_path: str | Path) -> Path:
    path = Path(relative_path)
    if path.is_absolute():
        return path
    return bundled_root() / path


def default_workbook_path() -> Path:
    return resource_path(DEFAULT_WORKBOOK_NAME)


def default_vba_export_dir() -> Path:
    return resource_path(DEFAULT_VBA_EXPORT_NAME)
