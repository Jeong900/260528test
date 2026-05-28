# DS 정기기부 대시보드

로컬 PC에 있는 월별 Excel 데이터를 읽어 DS부문 정기기부 현황을 달력, 사업부별 비교, 기간별 비교 화면으로 보여주는 대시보드입니다.

이 저장소는 외부 서버 없이 사용하는 배포 방식을 테스트 중입니다.

## 현재 추천 배포 방식

현재는 `feature/exe-distribution` 브랜치의 EXE 방식이 가장 안정적입니다.

```text
DS_Dashboard.exe
2605\
2604\
2603\
```

EXE를 실행하면 브라우저가 자동으로 열리고, `127.0.0.1:8765` 로컬 주소에서 동작합니다. 이 주소는 외부 인터넷이 아니라 사용자 PC 내부 통신입니다.

## 주요 파일

- `dashboard.html`: 사용자 화면, 차트, 다운로드 기능
- `dashboard_server.py`: 로컬 HTTP 서버, Excel 분석, 캐시, 유휴 종료
- `build_exe.ps1`: PyInstaller EXE 빌드 스크립트
- `EXE_DISTRIBUTION.md`: EXE 배포 및 테스트 안내
- `HTML_DISTRIBUTION.md`: HTML 단독 배포 브랜치 안내
- `DEVELOPMENT_HANDOFF.md`: 다음 개발자를 위한 상세 인수인계 문서

## 데이터 규칙

- 폴더명: `YYMM`, 예: `2605`
- 파일명: `YYMMDD_데이터(전체).xlsx` 또는 `YYMMDD_데이터(전체)_dummy.xlsx`
- 데이터 폴더와 Excel 파일은 Git에 올리지 않습니다.

## 빠른 실행

개발 중에는 Python 서버로 실행합니다.

```powershell
python -B .\dashboard_server.py
```

브라우저에서 엽니다.

```text
http://127.0.0.1:8765
```

## EXE 빌드

```powershell
.\build_exe.ps1
```

결과물:

```text
dist\DS_Dashboard.exe
```

## 보안 메모

이 저장소에는 GitHub 비밀번호, 토큰, 회사 계정 정보, 실제 업무 데이터 파일을 저장하지 않습니다. GitHub 로그인은 로컬 PC의 Git Credential Manager 또는 브라우저 인증 흐름을 사용합니다.
