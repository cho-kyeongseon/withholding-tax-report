# HANDOFF

## 지금까지 한 작업 (2026-11-12)
- 폴더를 Git 저장소로 만들고 공개 GitHub 저장소에 연결: https://github.com/cho-kyeongseon/withholding-tax-report (브랜치 main)
- .gitignore 추가: 급여대장·직원정보 엑셀, 생성된 신고서(hwp/pdf)·검토 엑셀은 제외하고, `program/`의 간이세액표 CSV와 별지21호 HWP 서식만 포함
- CLAUDE.md(작업 규칙), HANDOFF.md 작성
- Vercel은 웹사이트가 없는 PC용 프로그램이라 연결하지 않음

## 현재 상태
- 프로그램(`program/wht_report.py`)은 기존 그대로이며 이번 세션에서 코드 수정·실행 테스트는 하지 않음.
- 알려진 오류 없음.

## 남은 일
- (없음) 필요하면 bat 파일의 Python 경로 정리(아래 주의할 점 참고).

## 주의할 점
- USB(파일시스템이 소유자 기록 안 함)라서 새 PC에서는 Git이 "dubious ownership" 오류를 낼 수 있음 →
  `git config --global --add safe.directory "<USB드라이브>:/주말수업/원천세신고서_자동작성"` 실행.
- 새 PC에서 푸시하려면 GitHub CLI 설치 후 `gh auth login` 필요.
- `원천세신고서_만들기.bat`에 `C:\Users\seoil\anaconda3\python.exe`가 들어 있음(없으면 `python` 사용). 공개 저장소에 노출되는 사용자명.
- bat 파일은 CP949(ANSI) 인코딩이므로 UTF-8로 저장하지 말 것.
- 간이세액표·서식·세법 기준값은 해마다 개정 여부 확인(사용설명서.txt "해마다 바꿔야 하는 것").

## 실행 방법
- 필요: 한컴오피스 한글 2018 이상, Python(Anaconda 권장) + `openpyxl`, `pywin32` 패키지.
- 급여대장 엑셀을 `원천세신고서_만들기.bat` 위에 끌어다 놓기, 또는
  `python program\wht_report.py 급여대장.xlsx [--month 2025-06] [--paid 2025-06] [--yes]`
- 결과(HWP, 검토 엑셀)는 급여대장 엑셀과 같은 폴더에 생성됨.

## 원본 자료 위치
- 급여대장 엑셀·직원정보 엑셀은 GitHub에 없음. 사용자가 고른 엑셀 파일 위치(드래그한 파일의 폴더)에서 읽고 쓴다.
- 프로그램이 읽는 고정 자료: `program/간이세액표_2025.csv`, `program/별지21호_원천징수이행상황신고서_2025.hwp`
  (스크립트 위치 기준 상대 경로라 드라이브 문자와 무관).
