from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPlainTextEdit,
    QPushButton,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app.services.dry_run_analyzer import DryRunResult, RowError


class MainWindow(QMainWindow):
    """Main view for the stage-4 application."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("엑셀 기반 디렉토리 관리 도구")
        self.resize(1000, 760)

        central_widget = QWidget(self)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        self.create_excel_button = QPushButton("엑셀 생성")
        self.select_excel_button = QPushButton("엑셀 선택")
        self.select_root_button = QPushButton("루트 선택")
        self.dry_run_button = QPushButton("미리보기 (dry-run)")
        self.apply_button = QPushButton("적용 (apply)")
        self.exit_button = QPushButton("종료")

        self.create_excel_button.setMinimumWidth(110)
        self.select_excel_button.setMinimumWidth(110)
        self.select_root_button.setMinimumWidth(110)
        self.dry_run_button.setMinimumWidth(150)
        self.apply_button.setMinimumWidth(120)

        button_layout.addWidget(self.create_excel_button)
        button_layout.addWidget(self.select_excel_button)
        button_layout.addWidget(self.select_root_button)
        button_layout.addWidget(self.dry_run_button)
        button_layout.addWidget(self.apply_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.exit_button)

        path_group = QGroupBox("입력 경로")
        path_layout = QGridLayout(path_group)
        path_layout.setHorizontalSpacing(8)
        path_layout.setVerticalSpacing(8)

        path_label = QLabel("엑셀 파일")
        self.selected_path_edit = QLineEdit()
        self.selected_path_edit.setReadOnly(True)
        self.selected_path_edit.setPlaceholderText("아직 선택된 엑셀 파일이 없습니다.")
        self.selected_path_edit.setClearButtonEnabled(False)

        root_label = QLabel("루트 디렉토리")
        self.root_directory_edit = QLineEdit()
        self.root_directory_edit.setReadOnly(True)
        self.root_directory_edit.setPlaceholderText("미선택 시 엑셀 파일이 있는 폴더를 기본 루트로 사용합니다.")

        path_layout.addWidget(path_label, 0, 0)
        path_layout.addWidget(self.selected_path_edit, 0, 1)
        path_layout.addWidget(root_label, 1, 0)
        path_layout.addWidget(self.root_directory_edit, 1, 1)

        status_group = QGroupBox("진행 상태")
        status_layout = QVBoxLayout(status_group)
        self.status_message_label = QLabel("대기")
        self.status_message_label.setWordWrap(True)
        status_layout.addWidget(self.status_message_label)

        summary_group = QGroupBox("분석 요약")
        summary_layout = QGridLayout(summary_group)
        summary_layout.setHorizontalSpacing(24)
        summary_layout.setVerticalSpacing(8)

        self.total_rows_value = self._create_value_label("-")
        self.valid_rows_value = self._create_value_label("-")
        self.error_rows_value = self._create_value_label("-")
        self.create_count_value = self._create_value_label("-")
        self.delete_count_value = self._create_value_label("-")
        self.danger_count_value = self._create_value_label("-")
        self.final_status_value = self._create_value_label("대기")

        summary_layout.addWidget(QLabel("총 row 수"), 0, 0)
        summary_layout.addWidget(self.total_rows_value, 0, 1)
        summary_layout.addWidget(QLabel("유효 row 수"), 0, 2)
        summary_layout.addWidget(self.valid_rows_value, 0, 3)
        summary_layout.addWidget(QLabel("오류 row 수"), 0, 4)
        summary_layout.addWidget(self.error_rows_value, 0, 5)
        summary_layout.addWidget(QLabel("생성 예정 개수"), 1, 0)
        summary_layout.addWidget(self.create_count_value, 1, 1)
        summary_layout.addWidget(QLabel("삭제 후보 개수"), 1, 2)
        summary_layout.addWidget(self.delete_count_value, 1, 3)
        summary_layout.addWidget(QLabel("위험 폴더 개수"), 1, 4)
        summary_layout.addWidget(self.danger_count_value, 1, 5)
        summary_layout.addWidget(QLabel("최종 판정"), 2, 0)
        summary_layout.addWidget(self.final_status_value, 2, 1, 1, 5)

        results_group = QGroupBox("dry-run 결과")
        results_layout = QVBoxLayout(results_group)
        self.result_tabs = QTabWidget()
        self.create_candidates_output = self._create_result_box()
        self.delete_candidates_output = self._create_result_box()
        self.danger_folders_output = self._create_result_box()
        self.row_errors_output = self._create_result_box()
        self.result_tabs.addTab(self.create_candidates_output, "생성 예정")
        self.result_tabs.addTab(self.delete_candidates_output, "삭제 후보")
        self.result_tabs.addTab(self.danger_folders_output, "위험 폴더")
        self.result_tabs.addTab(self.row_errors_output, "row 오류")
        results_layout.addWidget(self.result_tabs)

        log_group = QGroupBox("실행 로그")
        log_layout = QVBoxLayout(log_group)
        self.log_output = self._create_result_box()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumBlockCount(2000)
        log_layout.addWidget(self.log_output)

        lower_layout = QHBoxLayout()
        lower_layout.setSpacing(12)
        lower_layout.addWidget(results_group, stretch=3)
        lower_layout.addWidget(log_group, stretch=2)

        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)

        main_layout.addLayout(button_layout)
        main_layout.addWidget(path_group)
        main_layout.addWidget(status_group)
        main_layout.addWidget(separator)
        main_layout.addWidget(summary_group)
        main_layout.addLayout(lower_layout, stretch=1)

        self.setCentralWidget(central_widget)
        self.clear_analysis_result()

    def set_selected_path(self, path: Path | None) -> None:
        self.selected_path_edit.setText(str(path) if path else "")

    def append_log(self, message: str) -> None:
        self.log_output.appendPlainText(message)
        scrollbar = self.log_output.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def set_root_directory(self, path: Path | None) -> None:
        self.root_directory_edit.setText(str(path) if path else "")

    def clear_analysis_result(self) -> None:
        self.total_rows_value.setText("-")
        self.valid_rows_value.setText("-")
        self.error_rows_value.setText("-")
        self.create_count_value.setText("-")
        self.delete_count_value.setText("-")
        self.danger_count_value.setText("-")
        self.final_status_value.setText("대기")
        self.create_candidates_output.setPlainText("아직 dry-run 결과가 없습니다.")
        self.delete_candidates_output.setPlainText("아직 dry-run 결과가 없습니다.")
        self.danger_folders_output.setPlainText("아직 dry-run 결과가 없습니다.")
        self.row_errors_output.setPlainText("아직 dry-run 결과가 없습니다.")

    def display_analysis_result(self, result: DryRunResult) -> None:
        if not result.success:
            self.clear_analysis_result()
            self.final_status_value.setText("불가")
            self.row_errors_output.setPlainText(result.fatal_error or "알 수 없는 오류가 발생했습니다.")
            return

        self.total_rows_value.setText(str(result.total_rows))
        self.valid_rows_value.setText(str(result.valid_rows))
        self.error_rows_value.setText(str(result.error_rows))
        self.create_count_value.setText(str(result.create_count))
        self.delete_count_value.setText(str(result.delete_count))
        self.danger_count_value.setText(str(result.danger_count))
        self.final_status_value.setText("가능" if result.is_applicable else "불가")

        self.create_candidates_output.setPlainText(self._format_items(result.create_candidates, "생성 예정 폴더가 없습니다."))
        self.delete_candidates_output.setPlainText(self._format_items(result.delete_candidates, "삭제 후보 폴더가 없습니다."))
        self.danger_folders_output.setPlainText(self._format_items(result.danger_folders, "위험 폴더가 없습니다."))
        self.row_errors_output.setPlainText(self._format_errors(result.row_errors))

    def set_status_message(self, message: str) -> None:
        self.status_message_label.setText(message)

    def _format_items(self, items: list[str], empty_message: str) -> str:
        if not items:
            return empty_message
        return "\n".join(f"- {item}" for item in items)

    def _format_errors(self, row_errors: list[RowError]) -> str:
        if not row_errors:
            return "row 오류가 없습니다."
        return "\n".join(f"- {error.row_number}행: {error.message}" for error in row_errors)

    def _create_result_box(self) -> QPlainTextEdit:
        output = QPlainTextEdit()
        output.setReadOnly(True)
        output.setLineWrapMode(QPlainTextEdit.NoWrap)
        output.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        font = QFont("Consolas")
        font.setStyleHint(QFont.Monospace)
        output.setFont(font)
        return output

    def _create_value_label(self, initial_text: str) -> QLabel:
        label = QLabel(initial_text)
        font = QFont()
        font.setBold(True)
        label.setFont(font)
        return label
