# 개발 인수인계 메모

작성일: 2026-05-28

이 문서는 다음 AI 또는 엔지니어가 현재 상태를 빠르게 이해하고 이어서 개발할 수 있도록 남긴 작업 기록입니다.

## 저장소 정보

로컬 경로:

```text
C:\Users\User\Desktop\0528
```

원격 저장소:

```text
https://github.com/Jeong900/260528test.git
```

중요: GitHub 로그인 정보, 비밀번호, Personal Access Token은 문서나 코드에 남기지 않는다. 현재 환경에서는 `git push` 시 이미 인증되어 있거나 Git Credential Manager가 브라우저 로그인을 처리한다.

현재 주요 브랜치:

- `dev`: 현재 기준 기능 브랜치
- `feature/exe-distribution`: EXE 배포 테스트 브랜치, 현재 추천
- `feature/html-only-distribution`: HTML 단독 배포 실험 브랜치
- `main`: 원격에 존재하지만 현재 작업 기준은 아님

최근 핵심 커밋:

- `4473ba2 Update dashboard period export`
- `3e84932 Add EXE distribution packaging`
- `fa71b73 Harden EXE build outputs`
- `6884f3a Hide EXE console and expire idle sessions`

최신 상태는 항상 아래 명령으로 확인한다.

```powershell
git status --short --branch
git log --oneline --decorate --max-count=10
git remote -v
```

## 현재 작업 상태 주의

2026-05-28 작업 종료 시점 기준, `feature/exe-distribution` 브랜치에서 작업했다.

`scratch_test.py`는 임시 테스트 파일이며 Git에 포함하지 않았다. 필요 없으면 삭제 가능하다.

`dashboard.html`에 IDE 자동 포맷처럼 보이는 대규모 미커밋 변경이 생길 수 있다. 기능 변경인지 단순 포맷인지 반드시 `git diff -- dashboard.html`로 확인한 뒤 커밋 여부를 결정한다.

데이터 폴더, 캐시, 빌드 산출물은 `.gitignore` 대상이다.

```text
2602\
2603\
2604\
2605\
.dashboard_cache\
build\
dist\
*.xlsx
*.spec
```

## 아키텍처 요약

현재 추천 방식은 로컬 EXE 배포다.

```text
사용자 브라우저
  -> http://127.0.0.1:8765
  -> DS_Dashboard.exe 내부 Python 서버
  -> EXE 옆 YYMM 데이터 폴더의 xlsx 파일 분석
  -> .dashboard_cache 에 분석 결과 캐시
```

외부 서버, 클라우드 업로드, 사내망 통신은 사용하지 않는다. `127.0.0.1`은 자기 PC 내부 주소다.

## 주요 코드 흐름

### `dashboard_server.py`

역할:

- `ThreadingHTTPServer`로 로컬 서버 실행
- `/` 또는 `/index.html`: `dashboard.html` 제공
- `/api/start-analysis?folder=2605`: 분석 작업 시작
- `/api/job?id=...`: 분석 진행률과 결과 조회
- `/api/activity`: 프론트엔드 heartbeat 수신
- 15분 유휴 시 `.dashboard_cache` 삭제 후 서버 종료

중요 함수:

- `analyze_folder(folder_name)`: 월별 폴더 분석의 핵심
- `read_workbook(path)`: xlsx 내부 XML을 직접 읽어 집계
- `start_analysis_job(folder_name)`: 백그라운드 분석 job 생성
- `start_idle_monitor(server)`: 15분 유휴 종료 감시

주의:

- Excel 파싱은 `openpyxl`이 아니라 `zipfile`과 XML streaming으로 직접 처리한다.
- 파일명 정규식 `FILE_RE`와 사업부명 정렬 `get_unit_sort_key`가 업무 규칙에 강하게 묶여 있다.
- EXE 환경에서는 `sys.frozen`, `sys._MEIPASS`를 사용해 번들 파일과 EXE 옆 데이터 폴더를 분리한다.

### `dashboard.html`

역할:

- 폴더 선택 UI
- 달력 현황
- 사업부별 비교
- 기간별 비교
- CSV/XLS 다운로드
- SVG 차트 PNG 저장
- `/api/activity` heartbeat 전송

주의:

- 서버 API에 의존한다. EXE 브랜치의 HTML은 단독 파일로만 열면 분석되지 않는다.
- HTML 단독 실험은 `feature/html-only-distribution` 브랜치에서 별도로 진행했다.

## 개발 실행

Python 서버로 빠르게 테스트:

```powershell
python -B .\dashboard_server.py
```

브라우저:

```text
http://127.0.0.1:8765
```

서버가 이미 떠 있으면 포트 충돌이 날 수 있다. 확인:

```powershell
netstat -ano | Select-String ":8765"
Get-Process DS_Dashboard -ErrorAction SilentlyContinue
```

테스트 프로세스 종료:

```powershell
Stop-Process -Name DS_Dashboard -Force
```

Python으로 실행한 서버는 해당 `python.exe` 프로세스를 확인한 뒤 종료한다.

## EXE 빌드 및 테스트

빌드:

```powershell
.\build_exe.ps1
```

스크립트는 PyInstaller를 사용한다. 빌드 옵션:

- `--onefile`: 단일 EXE 생성
- `--noconsole`: 검은 콘솔창 숨김
- `--add-data "dashboard.html;."`: HTML을 EXE에 포함

결과:

```text
dist\DS_Dashboard.exe
```

배포 테스트 구조:

```text
dist\
  DS_Dashboard.exe
  2605\
```

API 점검:

```powershell
Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8765/
Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8765/api/activity
```

분석 시작:

```powershell
$start = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:8765/api/start-analysis?folder=2605" | ConvertFrom-Json
$start.jobId
```

진행률 확인:

```powershell
Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:8765/api/job?id=$($start.jobId)" | ConvertFrom-Json
```

## 유휴 종료 정책

요구사항:

- 대시보드 사용 후 15분 동안 작동이 없으면 프로세스 종료
- 종료 시 캐시 삭제

구현:

- `dashboard.html`이 클릭, 키보드, 마우스 이동, 스크롤, 터치, 변경 이벤트를 감지한다.
- 활동이 있으면 `/api/activity`를 호출한다.
- `dashboard_server.py`는 마지막 활동 시간을 저장한다.
- 15분 동안 활동이 없으면 `.dashboard_cache`를 삭제하고 서버를 종료한다.

주의:

- 사용자가 브라우저를 열어두고 아무것도 하지 않으면 15분 뒤 종료된다.
- 종료 후에는 EXE를 다시 실행해야 한다.

## Git 작업 흐름

작업 시작:

```powershell
git status --short --branch
git fetch origin
git checkout feature/exe-distribution
```

작업 중 확인:

```powershell
git diff --stat
git diff -- dashboard_server.py
git diff -- dashboard.html
```

커밋:

```powershell
git add <수정한 파일>
git commit -m "간단한 변경 설명"
git push
```

중요:

- 사용자가 만든 변경을 임의로 되돌리지 않는다.
- 데이터 파일, 캐시, 빌드 산출물은 커밋하지 않는다.
- GitHub 인증 정보는 절대 커밋하지 않는다.

## 알려진 이슈와 다음 개선 후보

1. `dashboard.html`이 크고 단일 파일이다. 장기적으로 CSS/JS 분리를 고려할 수 있다.
2. EXE가 이미 실행 중일 때 새 EXE를 또 실행하면 포트 충돌 또는 중복 프로세스가 생길 수 있다. 단일 인스턴스 처리 개선 여지가 있다.
3. 첫 분석은 캐시가 없어 느릴 수 있다. 현재는 XML streaming과 6개 worker로 처리한다.
4. 15분 유휴 종료는 heartbeat 기반이다. 브라우저가 백그라운드에서 timer throttling을 강하게 적용하면 heartbeat 주기가 늘어날 수 있다.
5. HTML 단독 브랜치는 서버 없이 동작하는 실험판이다. 최신 Chrome/Edge API 의존성이 있어 EXE 방식보다 안정성이 낮다.
6. 배포 ZIP 자동 생성, 아이콘 적용, 버전 표시, 종료 안내 UI는 아직 미구현이다.

## 다음 AI에게 권장하는 첫 행동

1. `git status --short --branch`로 현재 브랜치와 미커밋 변경 확인
2. `dashboard.html` 변경이 단순 포맷인지 기능 변경인지 판단
3. EXE 방식 유지라면 `feature/exe-distribution`에서 작업
4. 수정 후 `python -m py_compile .\dashboard_server.py`와 HTML script syntax check 실행
5. `.\build_exe.ps1`로 새 EXE 빌드
6. `dist\DS_Dashboard.exe` 옆에 테스트 데이터 폴더를 두고 실제 분석 확인
