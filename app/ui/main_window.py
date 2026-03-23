from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class MainWindow(QMainWindow):
    """Main view for the stage-1 application."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("엑셀 기반 디렉토리 관리 도구")
        self.resize(760, 520)

        central_widget = QWidget(self)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        self.create_excel_button = QPushButton("엑셀 생성")
        self.select_excel_button = QPushButton("엑셀 선택")
        self.exit_button = QPushButton("종료")

        button_layout.addWidget(self.create_excel_button)
        button_layout.addWidget(self.select_excel_button)
        button_layout.addStretch(1)
        button_layout.addWidget(self.exit_button)

        path_label = QLabel("선택된 엑셀 파일")
        self.selected_path_edit = QLineEdit()
        self.selected_path_edit.setReadOnly(True)
        self.selected_path_edit.setPlaceholderText("아직 선택된 엑셀 파일이 없습니다.")

        log_label = QLabel("로그")
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)

        main_layout.addLayout(button_layout)
        main_layout.addWidget(path_label)
        main_layout.addWidget(self.selected_path_edit)
        main_layout.addWidget(log_label)
        main_layout.addWidget(self.log_output, stretch=1)

        self.setCentralWidget(central_widget)

    def set_selected_path(self, path: Path | None) -> None:
        self.selected_path_edit.setText(str(path) if path else "")

    def append_log(self, message: str) -> None:
        self.log_output.append(message)
