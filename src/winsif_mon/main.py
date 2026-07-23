from __future__ import annotations

import sys

from .qt_compat import require_qt


def main(argv: list[str] | None = None) -> int:
    """Launch the desktop application."""
    require_qt()
    from PySide6.QtWidgets import QApplication

    from .ui.main_window import MainWindow

    app = QApplication(sys.argv if argv is None else argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
