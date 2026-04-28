# Directory Management System

엑셀 원장을 기준으로 디렉토리 구조를 분석하고, dry-run과 apply를 통해 실제 폴더 구조를 안전하게 관리하는 Windows용 GUI 도구입니다.

## 1. 프로그램 개요

이 프로그램은 다음 작업을 지원합니다.

- 기본 원장 엑셀 생성
- 엑셀 row를 읽어 목표 디렉토리 구조 분석
- 실제 디렉토리와 비교한 dry-run 결과 표시
- 안전 검증 후 실제 반영(apply)
- 실제 마지막 depth 셀에 상대경로 하이퍼링크 갱신
- 실행 로그 및 엑셀 백업 생성
- 마지막 사용 엑셀 경로 / 루트 디렉토리 저장

기본 GUI는 `PySide6`, 엑셀 처리는 `openpyxl`, exe 패키징은 `PyInstaller`를 사용합니다.

## 2. 설치 및 실행

### 권장 환경

- OS: Windows 10/11
- Python: 3.10 ~ 3.12 권장
- 인코딩/경로 문제를 줄이기 위해 프로젝트 경로는 가능한 한 영문 경로 권장

### 가상환경(venv) 생성 및 의존성 설치 (권장)

PowerShell 기준:

```powershell
cd <프로젝트_루트>
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

CMD 기준:

```bat
cd <프로젝트_루트>
py -3 -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

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

```powershell
python -m PyInstaller --clean directory_management_system.spec
```

빌드가 끝나면 실행 파일은 `dist/DirectoryManagementSystem/` 아래에 생성됩니다.

> `pyinstaller` 명령이 인식되지 않으면, PATH 문제일 가능성이 큽니다.  
> 이 프로젝트에서는 `python -m PyInstaller ...` 형태를 기본 명령으로 사용하세요.

### 빌드 빠른 점검 명령

```powershell
python --version
python -m pip --version
python -m pip show pyinstaller
python -m PyInstaller --version
```

정상이라면 마지막 명령에서 버전이 출력됩니다.

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

- Depth1
- Depth2
- Depth3
- Depth4
- 비고

구조 규칙:

- `Depth1` 필수
- `Depth2~Depth4`는 선택
- 값이 존재하는 마지막 depth까지만 폴더 생성
- 중간 누락 구조는 오류  
  (예: `Depth1=값`, `Depth2=빈값`, `Depth3=값`)
- 중복 구조는 오류

예시:

| Depth1 | Depth2 | Depth3 | Depth4 | 비고 |
| --- | --- | --- | --- | --- |
| 사업A |  |  |  | 단일 depth |
| 사업A | 운영 |  |  | 2-depth |
| 사업A | 운영 | 정산 |  | 3-depth |
| 사업A | 운영 | 정산 | 비용정리 | 4-depth |

생성 경로 예시:

- `Depth1=A, Depth2=B, Depth3=빈값, Depth4=빈값` → `A/B`
- `Depth1=A, Depth2=B, Depth3=C, Depth4=빈값` → `A/B/C`
- `Depth1=A, Depth2=빈값, Depth3=빈값, Depth4=빈값` → `A`

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
6. 각 row의 **마지막 depth 셀** 하이퍼링크 갱신
7. 빈 폴더만 삭제
8. 로그 파일 기록

하이퍼링크 규칙:

- 각 row에서 실제 생성된 최종 폴더(depth 마지막 값)의 셀에 설정
- 텍스트는 유지
- 엑셀 파일 위치를 기준으로 한 상대경로 사용

## 8. exe 빌드 트러블슈팅 (Windows)

### 증상 1) `pyinstaller` 명령을 찾을 수 없음

오류 예:

```text
'pyinstaller' 용어가 ... 인식되지 않습니다.
```

원인:

- PyInstaller 미설치
- 설치는 되었지만 `Scripts` 경로가 PATH에 없음
- 다른 Python 인터프리터에 설치됨

해결:

```powershell
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller
python -m PyInstaller --clean directory_management_system.spec
```

핵심은 `pyinstaller` 직접 호출 대신 `python -m PyInstaller`를 사용하는 것입니다.

### 증상 2) 빌드 결과 폴더가 안 생김

점검:

```powershell
python -m PyInstaller --version
python -m PyInstaller --clean --noconfirm directory_management_system.spec
```

- 오류가 발생하면 해당 스택트레이스 기준으로 누락 모듈/권한/경로 이슈를 먼저 해결하세요.
- 일반적으로 결과물은 아래 경로에 생성됩니다.
  - `dist/DirectoryManagementSystem/DirectoryManagementSystem.exe`

### 증상 3) 실행 시 DLL/플러그인 관련 오류

- 보안 프로그램이 dist 내부 파일을 격리했는지 확인
- 한글/특수문자 경로 대신 짧은 영문 경로에서 재빌드
- 가상환경 새로 생성 후 재설치:

```powershell
deactivate
rmdir /s /q .venv
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m PyInstaller --clean directory_management_system.spec
```

## 9. 로그 설명

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

## 10. 주의사항

- 삭제는 **항상 마지막**에만 수행됩니다.
- 삭제 대상은 **완전히 빈 폴더만** 허용됩니다.
- 삭제 후보 중 하나라도 하위 파일/폴더가 있으면 apply는 중단됩니다.
- apply 전에 원본 엑셀 파일이 Excel, 미리보기, 동기화 충돌 상태로 열려 있으면 저장 가능 여부 검사에서 중단됩니다.
- apply 중 오류가 나면 가능한 범위에서 롤백을 시도하지만, 파일 시스템 작업은 운영체제 잠금 상태, Dropbox 동기화 지연, 외부 프로그램 점유, 강제 종료 상황 때문에 완전한 원자성을 보장할 수 없습니다.
- apply 전에 반드시 dry-run으로 결과를 확인하는 것을 권장합니다.

## 11. 파일 구성

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

## 12. CLI 예시

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
