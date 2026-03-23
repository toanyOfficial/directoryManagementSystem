from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from app.controller.main_controller import MainController
from app.services.dry_run_analyzer import DryRunAnalyzer
from app.services.excel_initializer import ExcelInitializer
from app.ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)

    window = MainWindow()
    initializer = ExcelInitializer()
    analyzer = DryRunAnalyzer()
    MainController(window, initializer, analyzer)

    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
