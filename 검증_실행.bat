@echo off
chcp 65001 > nul
python -m py_compile app.py
if errorlevel 1 (
  echo [실패] app.py 문법 오류가 있습니다.
  pause
  exit /b 1
)
echo [통과] app.py Python 문법 검사
python 엔진연결_검증.py
pause
