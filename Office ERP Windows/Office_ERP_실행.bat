@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if not exist ".venv\Scripts\python.exe" (
    echo [Office ERP] 가상환경을 준비합니다...
    py -3 -m venv .venv
    if errorlevel 1 (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo.
        echo Python을 찾지 못했습니다.
        echo Python 3.11 이상을 설치한 뒤 다시 실행해 주세요.
        echo https://www.python.org/downloads/windows/
        pause
        exit /b 1
    )
)

echo [Office ERP] 필요한 패키지를 확인합니다...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
)

echo [Office ERP] 런처를 실행합니다...
".venv\Scripts\python.exe" "ERP 업무 런처.py"
if errorlevel 1 (
    echo.
    echo 실행 중 오류가 발생했습니다.
    pause
)

endlocal
