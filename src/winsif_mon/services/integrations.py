from __future__ import annotations

from pathlib import Path

from winsif_mon.models import ExternalIntegration


def discover_integrations(root: Path) -> list[ExternalIntegration]:
    dxf_exe = root / "MF_CREA_WINDXF.EXE"
    dwg = root / "MF_AUTOCAD.dwg"
    return [
        ExternalIntegration(
            name="DXF generation",
            original_target=str(dxf_exe),
            replacement_behavior="Run optional generator if supplied; otherwise keep report data in-app.",
            available=dxf_exe.exists(),
        ),
        ExternalIntegration(
            name="AutoCAD drawing",
            original_target=str(dwg),
            replacement_behavior="Open DWG with the operating system default application.",
            available=dwg.exists(),
        ),
    ]
