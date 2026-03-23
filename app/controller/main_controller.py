from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QFileDialog

from app.services.excel_initializer import ExcelInitializer
from app.ui.main_window import MainWindow


class MainController:
    """Connects UI actions to stage-1 application logic."""

    def __init__(self, view: MainWindow, excel_initializer: ExcelInitializer) -> None:
        self.view = view
        self.excel_initializer = excel_initializer
        self.selected_excel_path: Path | None = None

        self._connect_signals()
        self._log("프로그램이 시작되었습니다.")

    def _connect_signals(self) -> None:
        self.view.create_excel_button.clicked.connect(self.create_excel)
        self.view.select_excel_button.clicked.connect(self.select_excel)
        self.view.exit_button.clicked.connect(self.view.close)

    def create_excel(self) -> None:
        current_directory = Path.cwd()
        self._log(f"엑셀 생성을 요청했습니다. 대상 폴더: {current_directory}")
        result = self.excel_initializer.create_template(current_directory)
        self._log(result.message)

        if result.success and result.path is not None:
            self.selected_excel_path = result.path
            self.view.set_selected_path(result.path)

    def select_excel(self) -> None:
        start_directory = str(Path.cwd())
        selected_file, _ = QFileDialog.getOpenFileName(
            self.view,
            "엑셀 파일 선택",
            start_directory,
            "Excel Files (*.xlsx)",
        )

        if not selected_file:
            self._log("엑셀 선택이 취소되었습니다.")
            return

        path = Path(selected_file)
        if not path.exists():
            self._log(f"선택한 파일을 찾을 수 없습니다: {path}")
            return

        self.selected_excel_path = path.resolve()
        self.view.set_selected_path(self.selected_excel_path)
        self._log(f"엑셀 파일을 선택했습니다: {self.selected_excel_path}")

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.view.append_log(f"[{timestamp}] {message}")
