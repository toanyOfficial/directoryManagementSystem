# Directory Management System

엑셀 원장을 기준으로 디렉토리 구조를 분석하고, dry-run과 apply를 통해 실제 폴더 구조를 안전하게 관리하는 Windows용 GUI 도구입니다.

## 1. 프로그램 개요

이 프로그램은 다음 작업을 지원합니다.

- 기본 원장 엑셀 생성
- 엑셀 row를 읽어 목표 디렉토리 구조 분석
- 실제 디렉토리와 비교한 dry-run 결과 표시
- 안전 검증 후 실제 반영(apply)
- `업무` 셀에 상대경로 하이퍼링크 갱신
- 실행 로그 및 엑셀 백업 생성
- 마지막 사용 엑셀 경로 / 루트 디렉토리 저장

기본 GUI는 `PySide6`, 엑셀 처리는 `openpyxl`, exe 패키징은 `PyInstaller`를 사용합니다.

## 2. 설치 및 실행

### GUI 실행

```bash
python -m pip install -r requirements.txt
python -m app.gui_app
```

### CLI 사용

```bash
python -m app.main --help
```

### Windows exe 패키징

```bash
pyinstaller --clean directory_management_system.spec
```

빌드가 끝나면 실행 파일은 `dist/DirectoryManagementSystem/` 아래에 생성됩니다.

## 3. 사용 방법

1. **엑셀 생성**
   - GUI의 `엑셀 생성` 버튼 또는 CLI의 `--init` 으로 기본 원장 파일을 생성합니다.
   - 기본 파일명은 `directory_master.xlsx` 입니다.

2. **엑셀 선택**
   - 분석/적용할 `.xlsx` 파일을 선택합니다.

3. **루트 선택**
   - 실제 비교/적용 대상 루트 폴더를 선택합니다.
   - 선택하지 않으면 엑셀 파일이 있는 폴더를 기본 루트로 사용합니다.

4. **미리보기 (dry-run)**
   - 엑셀 구조와 실제 디렉토리를 비교합니다.
   - 생성 예정 / 삭제 후보 / 위험 폴더 / row 오류를 탭으로 표시합니다.

5. **적용 (apply)**
   - 사전 검증을 다시 수행합니다.
   - 백업 생성 → 폴더 생성 → 하이퍼링크 갱신 → 빈 폴더 삭제 순서로 반영합니다.

## 4. 엑셀 작성 규칙

헤더는 아래 순서를 그대로 사용해야 합니다.

- 대분류
- 중분류
- 소분류
- 업무
- 비고

구조 규칙:

- `대분류` 필수
- `업무` 필수
- `중분류`, `소분류`는 선택
- `소분류`가 있으면 `중분류`도 반드시 있어야 함
- 중간 누락 구조는 오류
- 중복 구조는 오류

예시:

| 대분류 | 중분류 | 소분류 | 업무 | 비고 |
| --- | --- | --- | --- | --- |
| 사업A |  |  | 계획수립 |  |
| 사업A | 운영 |  | 월간보고 |  |
| 사업A | 운영 | 정산 | 비용정리 |  |

## 5. 폴더명 규칙

허용 규칙:

- 숫자
- 한글
- 영문 소문자
- `_`
- `.`

제한 규칙:

- 공백 금지
- Windows 예약어 금지 (`CON`, `PRN`, `AUX`, `NUL`, `COM1` 등)
- 마지막 글자 `.` 금지

엑셀 템플릿에는 영문 소문자 / 공백 금지 규칙을 위한 Data Validation이 포함되며, 프로그램 처리 단계에서도 영문은 소문자 기준으로 해석하고 공백은 제거한 뒤 최종 검증합니다.

## 6. dry-run 설명

dry-run은 엑셀을 실제로 수정하거나 폴더를 바꾸지 않고 결과만 보여주는 분석 단계입니다.

표시 항목:

- 총 row 수
- 유효 row 수
- 오류 row 수
- 생성 예정 개수
- 삭제 후보 개수
- 위험 폴더 개수
- 최종 판정 (가능 / 불가)

목록 항목:

- **생성 예정**: 엑셀 기준으로 있어야 하지만 아직 없는 폴더
- **삭제 후보**: 실제 폴더에는 있지만 엑셀 구조에 없는 폴더
- **위험 폴더**: 삭제 후보이면서 하위 파일/폴더가 있는 폴더
- **row 오류**: 필수값 누락, 이름 규칙 위반, 중간 누락, 중복 구조 등

`logs/`, `backups/`, `_internal/` 폴더는 시스템 관리용 폴더이므로 dry-run 비교와 위험 폴더 판단에서 제외합니다.

## 7. apply 설명

apply는 항상 아래 순서로 진행됩니다.

1. 사전 검증
2. 위험 폴더 확인
3. 엑셀 파일 저장 가능 여부 확인 (열려 있으면 중단)
4. 엑셀 백업 생성
5. 폴더 생성
6. `업무` 셀 하이퍼링크 갱신
7. 빈 폴더만 삭제
8. 로그 파일 기록

하이퍼링크 규칙:

- `업무` 셀에 설정
- 텍스트는 유지
- 엑셀 파일 위치를 기준으로 한 상대경로 사용

## 8. 로그 설명

apply 실행 시 루트 디렉토리 아래 `logs/` 폴더가 생성되며, `apply_YYYYMMDD_HHMMSS.log` 형식의 로그가 저장됩니다.

로그에는 아래 내용이 포함됩니다.

- 실행 시간
- 종료 시간
- 대상 루트
- 백업 파일
- 생성 폴더
- 삭제 폴더
- 오류
- 결과

예시:

```text
실행 시간: 2026-03-23 10:30:00
종료 시간: 2026-03-23 10:30:02
대상 루트: C:\managed_root
백업 파일: C:\managed_root\backups\directory_master_20260323_103000.xlsx
생성 폴더:
- 사업A
- 사업A\운영
- 사업A\운영\월간보고
삭제 폴더:
- 구폴더
오류:
- 없음
결과: 적용이 완료되었습니다.
```

## 9. 주의사항

- 삭제는 **항상 마지막**에만 수행됩니다.
- 삭제 대상은 **완전히 빈 폴더만** 허용됩니다.
- 삭제 후보 중 하나라도 하위 파일/폴더가 있으면 apply는 중단됩니다.
- apply 전에 원본 엑셀 파일이 Excel, 미리보기, 동기화 충돌 상태로 열려 있으면 저장 가능 여부 검사에서 중단됩니다.
- apply 중 오류가 나면 가능한 범위에서 롤백을 시도하지만, 파일 시스템 작업은 운영체제 잠금 상태, Dropbox 동기화 지연, 외부 프로그램 점유, 강제 종료 상황 때문에 완전한 원자성을 보장할 수 없습니다.
- apply 전에 반드시 dry-run으로 결과를 확인하는 것을 권장합니다.

## 10. 파일 구성

- `app/main.py`: 앱 진입점
- `app/gui_app.py`: GUI 실행 진입점
- `app/cli.py`: CLI 실행 진입점
- `app/ui/main_window.py`: 메인 GUI
- `app/controller/main_controller.py`: UI 이벤트 제어
- `app/services/excel_initializer.py`: 기본 엑셀 생성
- `app/services/dry_run_analyzer.py`: dry-run 분석
- `app/services/apply_service.py`: 실제 반영(apply)
- `app/services/settings_service.py`: 마지막 사용 경로 저장
- `app/utils/path_validator.py`: 폴더명 검증
- `directory_management_system.spec`: PyInstaller 설정

## 11. CLI 예시

### 초기화

```bash
python -m app.main --init
python -m app.main --init --file ./sample_master.xlsx
```

### dry-run

```bash
python -m app.main --file ./directory_master.xlsx --root /data/project_root --dry-run
```

### apply

```bash
python -m app.main --file ./directory_master.xlsx --root /data/project_root --apply
```
