from __future__ import annotations

import argparse
from pathlib import Path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="엑셀 기반 디렉토리 구조 관리 CLI",
    )
    parser.add_argument("--init", action="store_true", help="기본 엑셀 파일을 생성합니다.")
    parser.add_argument("--file", type=Path, help="대상 엑셀 파일 경로(.xlsx)")
    parser.add_argument("--root", type=Path, help="비교/적용 대상 루트 디렉토리")
    parser.add_argument("--dry-run", action="store_true", help="분석만 수행하고 로그를 남깁니다.")
    parser.add_argument("--apply", action="store_true", help="실제 반영을 수행합니다.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.dry_run and args.apply:
        parser.error("--dry-run 과 --apply 는 동시에 사용할 수 없습니다.")

    if not any((args.init, args.dry_run, args.apply)):
        parser.print_help()
        return 0

    from app.services.apply_service import ApplyService
    from app.services.dry_run_analyzer import DryRunAnalyzer
    from app.services.excel_initializer import ExcelInitializer
    from app.services.report_service import ReportService

    initializer = ExcelInitializer()
    analyzer = DryRunAnalyzer()
    apply_service = ApplyService(analyzer)
    report_service = ReportService()

    if args.init:
        target_path = args.file if args.file is not None else Path.cwd() / initializer.default_filename
        if target_path.suffix.lower() != ".xlsx":
            target_path = target_path.with_suffix(".xlsx")
        result = initializer.create_template_at(target_path)
        print(result.message)
        return 0 if result.success else 1

    if args.file is None:
        parser.error("--dry-run 또는 --apply 실행 시에는 반드시 --file 이 필요합니다.")

    excel_path = args.file
    root_path = args.root

    if args.dry_run:
        result = analyzer.analyze(excel_path, root_path)
        print(report_service.format_dry_run_report(excel_path, result))
        log_result = report_service.write_dry_run_log(excel_path, result)
        print(log_result.message)
        if log_result.log_path is not None:
            print(f"dry-run 로그 파일: {log_result.log_path}")
        return 0 if result.success else 1

    if args.apply:
        result = apply_service.apply(excel_path, root_path)
        print(_format_apply_output(result))
        return 0 if result.success else 1

    parser.print_help()
    return 0


def _format_apply_output(result) -> str:
    lines = [
        f"결과: {result.message}",
        f"상태: {result.status_message}",
        f"루트 디렉토리: {result.target_root if result.target_root else '-'}",
        f"백업 파일: {result.backup_path if result.backup_path else '-'}",
        f"로그 파일: {result.log_path if result.log_path else '-'}",
        f"갱신된 하이퍼링크 row 수: {result.hyperlink_updated_rows}",
        "실제 생성된 폴더:",
        *_format_items(result.created_folders),
        "실제 삭제된 폴더:",
        *_format_items(result.deleted_folders),
        "오류:",
        *_format_items(result.errors),
        "롤백 시도:",
        *_format_items(result.rollback_actions),
    ]
    return "\n".join(lines)


def _format_items(items: list[str]) -> list[str]:
    if not items:
        return ["- 없음"]
    return [f"- {item}" for item in items]
