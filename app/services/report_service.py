from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.services.dry_run_analyzer import DryRunResult


@dataclass(frozen=True)
class DryRunLogResult:
    success: bool
    message: str
    log_path: Path | None = None


class ReportService:
    """Create text reports for CLI dry-run/apply operations."""

    def write_dry_run_log(self, excel_path: Path, result: DryRunResult) -> DryRunLogResult:
        target_root = result.target_root if result.target_root is not None else excel_path.parent
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_directory = target_root / "logs"
        log_path = log_directory / f"dry_run_{timestamp}.log"

        try:
            log_directory.mkdir(parents=True, exist_ok=True)
            log_path.write_text(self.format_dry_run_report(excel_path, result), encoding="utf-8")
        except OSError as exc:
            return DryRunLogResult(False, f"dry-run 로그 파일 저장 실패: {exc}")

        return DryRunLogResult(True, "dry-run 로그를 저장했습니다.", log_path)

    def format_dry_run_report(self, excel_path: Path, result: DryRunResult) -> str:
        lines = [
            f"엑셀 파일: {excel_path}",
            f"루트 디렉토리: {result.target_root if result.target_root else '-'}",
            f"총 row 수: {result.total_rows}",
            f"유효 row 수: {result.valid_rows}",
            f"오류 row 수: {result.error_rows}",
            f"생성 예정 개수: {result.create_count}",
            f"삭제 후보 개수: {result.delete_count}",
            f"위험 폴더 개수: {result.danger_count}",
            f"최종 판정: {'가능' if result.is_applicable else '불가'}",
            "생성 예정:",
            *self._format_items(result.create_candidates),
            "삭제 후보:",
            *self._format_items(result.delete_candidates),
            "위험 폴더:",
            *self._format_items(result.danger_folders),
            "row 오류:",
            *self._format_errors(result),
        ]
        return "\n".join(lines) + "\n"

    def _format_items(self, items: list[str]) -> list[str]:
        if not items:
            return ["- 없음"]
        return [f"- {item}" for item in items]

    def _format_errors(self, result: DryRunResult) -> list[str]:
        if not result.success:
            return [f"- {result.fatal_error or '알 수 없는 오류'}"]
        if not result.row_errors:
            return ["- 없음"]
        return [f"- {error.row_number}행: {error.message}" for error in result.row_errors]
