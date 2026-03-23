from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QFileDialog

from app.services.apply_service import ApplyService
from app.services.dry_run_analyzer import DryRunAnalyzer
from app.services.excel_initializer import ExcelInitializer
from app.ui.main_window import MainWindow


class MainController:
    """Connects UI actions to stage-3 application logic."""

    def __init__(
        self,
        view: MainWindow,
        excel_initializer: ExcelInitializer,
        dry_run_analyzer: DryRunAnalyzer,
        apply_service: ApplyService,
    ) -> None:
        self.view = view
        self.excel_initializer = excel_initializer
        self.dry_run_analyzer = dry_run_analyzer
        self.apply_service = apply_service
        self.selected_excel_path: Path | None = None

        self._connect_signals()
        self.view.set_status_message("대기")
        self._log("프로그램이 시작되었습니다.")

    def _connect_signals(self) -> None:
        self.view.create_excel_button.clicked.connect(self.create_excel)
        self.view.select_excel_button.clicked.connect(self.select_excel)
        self.view.dry_run_button.clicked.connect(self.run_dry_run)
        self.view.apply_button.clicked.connect(self.apply_changes)
        self.view.exit_button.clicked.connect(self.view.close)

    def create_excel(self) -> None:
        current_directory = Path.cwd()
        self._log(f"엑셀 생성을 요청했습니다. 대상 폴더: {current_directory}")
        result = self.excel_initializer.create_template(current_directory)
        self._log(result.message)

        if result.success and result.path is not None:
            self.selected_excel_path = result.path
            self.view.set_selected_path(result.path)
            self.view.clear_analysis_result()
            self.view.set_status_message("엑셀 생성 완료")
        else:
            self.view.set_status_message("엑셀 생성 실패")

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
            self.view.set_status_message("엑셀 선택 취소")
            return

        path = Path(selected_file)
        if not path.exists():
            self._log(f"선택한 파일을 찾을 수 없습니다: {path}")
            self.view.set_status_message("엑셀 선택 실패")
            return

        self.selected_excel_path = path.resolve()
        self.view.set_selected_path(self.selected_excel_path)
        self.view.clear_analysis_result()
        self.view.set_status_message("엑셀 선택 완료")
        self._log(f"엑셀 파일을 선택했습니다: {self.selected_excel_path}")

    def run_dry_run(self) -> None:
        if self.selected_excel_path is None:
            self._log("dry-run을 시작할 수 없습니다. 먼저 엑셀 파일을 선택하세요.")
            self.view.clear_analysis_result()
            self.view.set_status_message("dry-run 대기")
            return

        self.view.set_status_message("dry-run 실행 중")
        self._log(f"dry-run을 시작합니다: {self.selected_excel_path}")
        result = self.dry_run_analyzer.analyze(self.selected_excel_path)
        self.view.display_analysis_result(result)

        if not result.success:
            self.view.set_status_message("dry-run 실패")
            self._log(f"dry-run 치명적 오류: {result.fatal_error}")
            return

        self.view.set_status_message("dry-run 완료")
        self._log(f"총 row 수: {result.total_rows}")
        self._log(f"유효 row 수: {result.valid_rows}")
        self._log(f"오류 row 수: {result.error_rows}")
        self._log(f"생성 예정 개수: {result.create_count}")
        self._log(f"삭제 후보 개수: {result.delete_count}")
        self._log(f"위험 폴더 개수: {result.danger_count}")
        self._log(f"최종 판정: {'가능' if result.is_applicable else '불가'}")

    def apply_changes(self) -> None:
        if self.selected_excel_path is None:
            self._log("apply를 시작할 수 없습니다. 먼저 엑셀 파일을 선택하세요.")
            self.view.set_status_message("apply 대기")
            return

        self.view.set_status_message("apply 실행 중")
        self._log(f"apply를 시작합니다: {self.selected_excel_path}")
        result = self.apply_service.apply(self.selected_excel_path)

        self.view.set_status_message(result.status_message)
        if result.dry_run_result is not None:
            self.view.display_analysis_result(result.dry_run_result)

        if result.backup_path is not None:
            self._log(f"엑셀 백업 생성: {result.backup_path}")

        if result.log_path is not None:
            self._log(f"로그 파일 경로: {result.log_path}")

        for folder in result.created_folders:
            self._log(f"생성 완료: {folder}")

        for folder in result.deleted_folders:
            self._log(f"삭제 완료: {folder}")

        for error in result.errors:
            self._log(f"오류: {error}")

        self._log(result.message)

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.view.append_log(f"[{timestamp}] {message}")
