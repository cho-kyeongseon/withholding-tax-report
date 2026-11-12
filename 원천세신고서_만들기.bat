@echo off
title 원천징수이행상황신고서 자동 작성
setlocal

if "%~1"=="" (
  echo.
  echo  급여대장 엑셀 파일을 이 아이콘 위에 끌어다 놓으세요.
  echo.
  pause
  exit /b
)

set "PY=C:\Users\seoil\anaconda3\python.exe"
if not exist "%PY%" set "PY=python"
set PYTHONIOENCODING=utf-8

"%PY%" "%~dp0program\wht_report.py" "%~1"
if errorlevel 1 goto fail
echo.
echo  완료되었습니다. 결과 파일이 있는 폴더를 엽니다.
explorer "%~dp1"
goto end
:fail
echo.
echo  작성하지 못했습니다. 위 메시지를 확인하세요.
:end
echo.
pause