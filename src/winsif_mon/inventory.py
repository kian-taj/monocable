from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .resources import default_vba_export_dir, default_workbook_path
from .vba_parser import scan_export


DEFAULT_EXPORT_DIR = default_vba_export_dir()
DEFAULT_WORKBOOK = default_workbook_path()


def inspect_workbook(path: Path) -> dict[str, Any]:
    """Return workbook metadata when xlrd is available.

    Legacy `.xls` parsing is optional so the PySide6 shell can run without Excel
    extraction dependencies. Values/formulas/layout extraction will be expanded
    in later migration phases.
    """
    if not path.exists():
        return {"path": str(path), "exists": False, "sheets": []}
    try:
        import xlrd
    except ModuleNotFoundError:
        return {
            "path": str(path),
            "exists": True,
            "sheets": [],
            "warning": "Install xlrd to inspect legacy .xls sheet metadata.",
        }

    book = xlrd.open_workbook(path, formatting_info=True)
    sheets = [
        {
            "name": sheet.name,
            "rows": sheet.nrows,
            "cols": sheet.ncols,
        }
        for sheet in book.sheets()
    ]
    return {"path": str(path), "exists": True, "sheets": sheets}


def build_inventory(export_dir: Path | None = None, workbook: Path | None = None) -> dict[str, Any]:
    export_dir = export_dir or default_vba_export_dir()
    workbook = workbook or default_workbook_path()
    components = scan_export(export_dir)
    return {
        "export_dir": str(export_dir),
        "workbook": inspect_workbook(workbook),
        "components": [asdict(component) | {"path": str(component.path)} for component in components],
        "summary": {
            "component_count": len(components),
            "procedure_count": sum(len(component.procedures) for component in components),
            "referenced_sheets": sorted({sheet for component in components for sheet in component.sheet_refs}),
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inventory MONOFUNI workbook/VBA migration inputs.")
    parser.add_argument("--export-dir", type=Path)
    parser.add_argument("--workbook", type=Path)
    parser.add_argument("--json", action="store_true", help="Print JSON instead of a short text summary.")
    args = parser.parse_args(argv)

    inventory = build_inventory(args.export_dir, args.workbook)
    if args.json:
        print(json.dumps(inventory, indent=2, ensure_ascii=False))
        return 0

    summary = inventory["summary"]
    print(f"Components: {summary['component_count']}")
    print(f"Procedures: {summary['procedure_count']}")
    print("Referenced sheets: " + ", ".join(summary["referenced_sheets"]))
    workbook = inventory["workbook"]
    print(f"Workbook: {workbook['path']}")
    if workbook.get("warning"):
        print(f"Workbook warning: {workbook['warning']}")
    else:
        print(f"Workbook sheets: {len(workbook.get('sheets', []))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
