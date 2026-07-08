@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"

set "REPO_URL=https://github.com/ryangroe0618-maker/office-erp.git"
set "BRANCH=main"
set "PS_UPDATE=%~dp0업데이트_확인.ps1"

where git > nul 2> nul
if errorlevel 1 (
    echo Git이 설치되어 있지 않아 EXE 릴리즈 업데이트로 진행합니다.
    echo.
    goto release_update
)

if not exist ".git" (
    echo ZIP으로 받은 설치본입니다. EXE 릴리즈 업데이트로 진행합니다.
    echo.
    goto release_update
)

echo [Office ERP] GitHub에서 최신 코드를 가져옵니다...
git remote set-url origin %REPO_URL%
git fetch origin %BRANCH%
if errorlevel 1 (
    echo.
    echo 최신 코드 확인에 실패했습니다. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
)

git pull --ff-only origin %BRANCH%
if errorlevel 1 (
    echo.
    echo 자동 업데이트에 실패했습니다.
    echo 변경 충돌이 있거나 네트워크 문제가 있을 수 있습니다.
    pause
    exit /b 1
)

echo.
echo 업데이트 완료.
goto launch_app

:release_update
if exist "%PS_UPDATE%" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_UPDATE%"
) else (
    echo 업데이트 확인 스크립트를 찾지 못했습니다.
)

:launch_app
echo.
if exist "Office_ERP.exe" (
    echo Office_ERP.exe 를 실행합니다...
    start "" "%~dp0Office_ERP.exe"
    exit /b 0
)

if exist "START_WINDOWS.cmd" (
    echo START_WINDOWS.cmd 를 실행합니다...
    call "START_WINDOWS.cmd"
    exit /b %errorlevel%
)

echo 실행 파일을 찾지 못했습니다.
pause
endlocal
