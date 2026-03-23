from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

_HEADERS: tuple[str, ...] = ("대분류", "중분류", "소분류", "업무", "비고")
_COLUMN_WIDTHS: dict[str, float] = {
    "A": 20,
    "B": 20,
    "C": 20,
    "D": 28,
    "E": 36,
}
_MAX_EXCEL_ROWS = 1048576


@dataclass(frozen=True)
class ExcelCreationResult:
    success: bool
    message: str
    path: Path | None = None


class ExcelInitializer:
    """Create the stage-1 Excel master file template."""

    def __init__(self, default_filename: str = "directory_master.xlsx") -> None:
        self.default_filename = default_filename

    def create_template(self, directory: Path) -> ExcelCreationResult:
        try:
            target_directory = directory.expanduser().resolve()
        except OSError as exc:
            return ExcelCreationResult(False, f"경로를 해석할 수 없습니다: {exc}")

        if not target_directory.exists():
            return ExcelCreationResult(False, f"대상 경로가 존재하지 않습니다: {target_directory}")

        if not target_directory.is_dir():
            return ExcelCreationResult(False, f"대상 경로가 폴더가 아닙니다: {target_directory}")

        target_path = target_directory / self.default_filename
        if target_path.exists():
            return ExcelCreationResult(False, f"이미 파일이 존재합니다: {target_path}", target_path)

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "원장"

        self._write_headers(worksheet)
        self._apply_column_widths(worksheet)
        self._apply_data_validation(worksheet)
        worksheet.freeze_panes = "A2"

        try:
            workbook.save(target_path)
        except OSError as exc:
            return ExcelCreationResult(False, f"엑셀 파일 저장에 실패했습니다: {exc}")

        return ExcelCreationResult(True, f"엑셀 파일을 생성했습니다: {target_path}", target_path)

    def _write_headers(self, worksheet: Worksheet) -> None:
        header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        header_font = Font(bold=True)
        header_alignment = Alignment(horizontal="center", vertical="center")

        for column_index, header in enumerate(_HEADERS, start=1):
            cell = worksheet.cell(row=1, column=column_index, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

    def _apply_column_widths(self, worksheet: Worksheet) -> None:
        for column_letter, width in _COLUMN_WIDTHS.items():
            worksheet.column_dimensions[column_letter].width = width

    def _apply_data_validation(self, worksheet: Worksheet) -> None:
        for column_letter in ("A", "B", "C", "D"):
            validation = DataValidation(
                type="custom",
                formula1=(
                    f'=AND({column_letter}2<>"",'
                    f'ISERROR(SEARCH(" ",{column_letter}2)),'
                    f'RIGHT({column_letter}2,1)<>".")'
                ),
                allow_blank=True,
                showErrorMessage=True,
                errorTitle="입력 제한",
                error=(
                    "공백 없이 입력하고 마지막 글자에 '.'을 사용할 수 없습니다. "
                    "최종 유효성 검사는 프로그램 기준을 따릅니다."
                ),
                promptTitle="입력 규칙",
                prompt="숫자/한글/영어/언더스코어/점만 사용하는 것을 권장합니다.",
            )
            validation.add(f"{column_letter}2:{column_letter}{_MAX_EXCEL_ROWS}")
            worksheet.add_data_validation(validation)
