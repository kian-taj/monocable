from __future__ import annotations

import sys
from pathlib import Path

import pytest

from winsif_mon.resources import default_workbook_path, resource_path
from winsif_mon.workbook import WorkbookReader


def test_default_workbook_path_uses_pyinstaller_meipass(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)

    assert default_workbook_path() == tmp_path / "MONOFUNI.xls"
    assert resource_path("other.dat") == tmp_path / "other.dat"


def test_workbook_reader_reports_missing_bundled_workbook(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(tmp_path), raising=False)

    with pytest.raises(RuntimeError, match="include MONOFUNI.xls as bundled data"):
        WorkbookReader()
