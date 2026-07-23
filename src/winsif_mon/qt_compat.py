from __future__ import annotations


def require_qt() -> None:
    """Raise a clear error when PySide6 is not installed."""
    try:
        import PySide6  # noqa: F401
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "PySide6 is not installed. Install dependencies with: "
            "python3 -m pip install -r requirements.txt"
        ) from exc
