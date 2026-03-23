from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.services.dry_run_analyzer import DryRunResult, RowError


class MainWindow(QMainWindow):
    """Main view for the stage-3 application."""

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
        self.dry_run_button = QPushButton("미리보기 (dry-run)")
        self.apply_button = QPushButton("적용 (apply)")
        self.exit_button = QPushButton("종료")

        button_layout.addWidget(self.create_excel_button)
        button_layout.addWidget(self.select_excel_button)
        button_layout.addWidget(self.dry_run_button)
        button_layout.addWidget(self.apply_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.exit_button)

        path_label = QLabel("선택된 엑셀 파일")
        self.selected_path_edit = QLineEdit()
        self.selected_path_edit.setReadOnly(True)
        self.selected_path_edit.setPlaceholderText("아직 선택된 엑셀 파일이 없습니다.")

        status_group = QGroupBox("진행 상태")
        status_layout = QVBoxLayout(status_group)
        self.status_message_label = QLabel("대기")
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

        log_label = QLabel("로그")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        main_layout.addLayout(button_layout)
        main_layout.addWidget(path_label)
        main_layout.addWidget(self.selected_path_edit)
        main_layout.addWidget(status_group)
        main_layout.addWidget(summary_group)
        main_layout.addWidget(results_group, stretch=1)
        main_layout.addWidget(log_label)
        main_layout.addWidget(self.log_output, stretch=1)

        self.setCentralWidget(central_widget)
        self.clear_analysis_result()

    def set_selected_path(self, path: Path | None) -> None:
        self.selected_path_edit.setText(str(path) if path else "")

    def append_log(self, message: str) -> None:
        self.log_output.append(message)

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
        return output

    def _create_value_label(self, initial_text: str) -> QLabel:
        label = QLabel(initial_text)
        font = QFont()
        font.setBold(True)
        label.setFont(font)
        return label
