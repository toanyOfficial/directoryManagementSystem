from __future__ import annotations

import re
from dataclasses import dataclass

_ALLOWED_NAME_PATTERN = re.compile(r"^[0-9A-Za-z가-힣_.]+$")
_WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
}


@dataclass(frozen=True)
class ValidationResult:
    is_valid: bool
    message: str


class FolderNameValidator:
    """Validate folder names based on the project rules."""

    @staticmethod
    def validate(name: str) -> ValidationResult:
        candidate = name.strip()

        if not candidate:
            return ValidationResult(False, "폴더명은 비어 있을 수 없습니다.")

        if candidate != name:
            return ValidationResult(False, "폴더명 앞뒤 공백은 허용되지 않습니다.")

        if " " in candidate:
            return ValidationResult(False, "폴더명에 공백은 허용되지 않습니다.")

        if candidate.endswith("."):
            return ValidationResult(False, "폴더명 마지막 글자는 '.' 일 수 없습니다.")

        if candidate.upper() in _WINDOWS_RESERVED_NAMES:
            return ValidationResult(False, "Windows 예약어는 폴더명으로 사용할 수 없습니다.")

        if not _ALLOWED_NAME_PATTERN.fullmatch(candidate):
            return ValidationResult(
                False,
                "폴더명에는 숫자, 한글, 영어, '_', '.' 만 사용할 수 있습니다.",
            )

        return ValidationResult(True, "유효한 폴더명입니다.")
