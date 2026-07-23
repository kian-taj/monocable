from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


PROC_RE = re.compile(
    r"^\s*(?:Private\s+|Public\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)",
    re.I | re.M,
)
SHEET_RE = re.compile(r'Sheets\("([^"]+)"\)', re.I)


@dataclass(frozen=True, slots=True)
class VbaComponent:
    name: str
    kind: str
    path: Path
    procedures: tuple[str, ...]
    sheet_refs: tuple[str, ...]
    line_count: int


def scan_component(path: Path, kind: str) -> VbaComponent:
    text = path.read_text(encoding="latin-1", errors="replace")
    procedures = tuple(match.group(1) for match in PROC_RE.finditer(text))
    sheet_refs = tuple(sorted(set(SHEET_RE.findall(text))))
    return VbaComponent(
        name=path.stem,
        kind=kind,
        path=path,
        procedures=procedures,
        sheet_refs=sheet_refs,
        line_count=len(text.splitlines()),
    )


def scan_export(export_dir: Path) -> list[VbaComponent]:
    groups = [
        ("module", export_dir / "Modules", "*.bas"),
        ("sheet", export_dir / "ExcelObjects", "*.cls"),
        ("class", export_dir / "ClassModules", "*.cls"),
        ("form", export_dir / "Forms", "*.frm"),
    ]
    components: list[VbaComponent] = []
    for kind, folder, pattern in groups:
        if not folder.exists():
            continue
        for path in sorted(folder.glob(pattern)):
            components.append(scan_component(path, kind))
    return components
