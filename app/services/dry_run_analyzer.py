from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath

from openpyxl import load_workbook

from app.services.excel_schema import EXCEL_HEADERS
from app.utils.path_validator import FolderNameValidator


@dataclass(frozen=True)
class RowError:
    row_number: int
    message: str


@dataclass(frozen=True)
class ParsedRow:
    row_number: int
    path_parts: tuple[str, ...]

    @property
    def relative_path(self) -> Path:
        return Path(*self.path_parts)

    @property
    def display_path(self) -> str:
        return str(PureWindowsPath(*self.path_parts))


@dataclass(frozen=True)
class DryRunResult:
    success: bool
    fatal_error: str | None
    target_root: Path | None
    total_rows: int
    valid_rows: int
    error_rows: int
    create_count: int
    delete_count: int
    danger_count: int
    is_applicable: bool
    create_candidates: list[str] = field(default_factory=list)
    delete_candidates: list[str] = field(default_factory=list)
    danger_folders: list[str] = field(default_factory=list)
    row_errors: list[RowError] = field(default_factory=list)


class DryRunAnalyzer:
    """Analyze Excel rows and compare them with the actual directory structure."""

    def analyze(self, excel_path: Path) -> DryRunResult:
        try:
            resolved_excel_path = excel_path.expanduser().resolve()
        except OSError as exc:
            return self._fatal_result(f"엑셀 경로를 해석할 수 없습니다: {exc}")

        if not resolved_excel_path.exists():
            return self._fatal_result(f"엑셀 파일이 존재하지 않습니다: {resolved_excel_path}")

        if resolved_excel_path.suffix.lower() != ".xlsx":
            return self._fatal_result("지원하지 않는 파일 형식입니다. .xlsx 파일만 선택할 수 있습니다.")

        try:
            workbook = load_workbook(resolved_excel_path, data_only=True)
        except Exception as exc:
            return self._fatal_result(f"엑셀 파일을 읽을 수 없습니다: {exc}")

        try:
            worksheet = workbook.active
            header_values = tuple(
                str(worksheet.cell(row=1, column=index).value or "").strip() for index in range(1, 6)
            )
            if header_values != EXCEL_HEADERS:
                return self._fatal_result(
                    "엑셀 헤더가 올바르지 않습니다. "
                    f"기대값: {', '.join(EXCEL_HEADERS)} / 실제값: {', '.join(header_values)}"
                )

            row_errors: list[RowError] = []
            parsed_rows: list[ParsedRow] = []
            seen_paths: set[Path] = set()
            total_rows = 0

            for row_number in range(2, worksheet.max_row + 1):
                values = [worksheet.cell(row=row_number, column=index).value for index in range(1, 6)]
                normalized_values = [str(value).strip() if value is not None else "" for value in values]

                if not any(normalized_values):
                    continue

                total_rows += 1
                validation_messages = self._validate_row(normalized_values)
                if validation_messages:
                    row_errors.append(RowError(row_number, "; ".join(validation_messages)))
                    continue

                path_parts = tuple(value for value in normalized_values[:4] if value)
                parsed_row = ParsedRow(row_number=row_number, path_parts=path_parts)
                if parsed_row.relative_path in seen_paths:
                    row_errors.append(RowError(row_number, f"중복 구조입니다: {parsed_row.display_path}"))
                    continue

                seen_paths.add(parsed_row.relative_path)
                parsed_rows.append(parsed_row)
        finally:
            workbook.close()

        target_root = resolved_excel_path.parent
        expected_directories = self._build_expected_directories(parsed_rows)
        actual_directories = self._scan_actual_directories(target_root)

        create_candidates = sorted(expected_directories - actual_directories, key=self._sort_key)
        delete_candidates = sorted(actual_directories - expected_directories, key=self._sort_key)
        danger_folders = []
        for relative_path in delete_candidates:
            candidate_path = target_root / relative_path
            try:
                has_children = any(candidate_path.iterdir())
            except OSError:
                has_children = True
            if has_children:
                danger_folders.append(relative_path)

        return DryRunResult(
            success=True,
            fatal_error=None,
            target_root=target_root,
            total_rows=total_rows,
            valid_rows=len(parsed_rows),
            error_rows=len(row_errors),
            create_count=len(create_candidates),
            delete_count=len(delete_candidates),
            danger_count=len(danger_folders),
            is_applicable=(len(row_errors) == 0 and len(danger_folders) == 0),
            create_candidates=[self._format_path(path) for path in create_candidates],
            delete_candidates=[self._format_path(path) for path in delete_candidates],
            danger_folders=[self._format_path(path) for path in danger_folders],
            row_errors=row_errors,
        )

    def _validate_row(self, row_values: list[str]) -> list[str]:
        major, middle, minor, task, _note = row_values
        errors: list[str] = []

        if not major:
            errors.append("대분류는 필수입니다.")

        if not task:
            errors.append("업무는 필수입니다.")

        if not middle and minor:
            errors.append("중간 누락 구조는 허용되지 않습니다.")

        for field_name, value in (
            ("대분류", major),
            ("중분류", middle),
            ("소분류", minor),
            ("업무", task),
        ):
            if not value:
                continue
            validation_result = FolderNameValidator.validate(value)
            if not validation_result.is_valid:
                errors.append(f"{field_name}: {validation_result.message}")

        return errors

    def _build_expected_directories(self, parsed_rows: list[ParsedRow]) -> set[Path]:
        expected_directories: set[Path] = set()
        for parsed_row in parsed_rows:
            current_path = Path()
            for part in parsed_row.path_parts:
                current_path /= part
                expected_directories.add(current_path)
        return expected_directories

    def _scan_actual_directories(self, target_root: Path) -> set[Path]:
        actual_directories: set[Path] = set()
        for path in target_root.rglob("*"):
            if not path.is_dir():
                continue
            actual_directories.add(path.relative_to(target_root))
        return actual_directories

    def _sort_key(self, path: Path) -> tuple[int, str]:
        return (len(path.parts), self._format_path(path))

    def _format_path(self, path: Path) -> str:
        return str(PureWindowsPath(*path.parts))

    def _fatal_result(self, message: str) -> DryRunResult:
        return DryRunResult(
            success=False,
            fatal_error=message,
            target_root=None,
            total_rows=0,
            valid_rows=0,
            error_rows=0,
            create_count=0,
            delete_count=0,
            danger_count=0,
            is_applicable=False,
        )
