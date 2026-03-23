from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QSettings


@dataclass(frozen=True)
class AppSettingsData:
    last_excel_path: Path | None = None
    last_root_directory: Path | None = None


class SettingsService:
    """Persist lightweight user preferences such as the last used paths."""

    def __init__(self) -> None:
        self._settings = QSettings("OpenAI", "DirectoryManagementSystem")

    def load(self) -> AppSettingsData:
        excel_value = self._settings.value("paths/last_excel_path", "", type=str)
        root_value = self._settings.value("paths/last_root_directory", "", type=str)
        return AppSettingsData(
            last_excel_path=Path(excel_value).expanduser() if excel_value else None,
            last_root_directory=Path(root_value).expanduser() if root_value else None,
        )

    def save_last_excel_path(self, path: Path | None) -> None:
        self._settings.setValue("paths/last_excel_path", str(path) if path else "")
        self._settings.sync()

    def save_last_root_directory(self, path: Path | None) -> None:
        self._settings.setValue("paths/last_root_directory", str(path) if path else "")
        self._settings.sync()
