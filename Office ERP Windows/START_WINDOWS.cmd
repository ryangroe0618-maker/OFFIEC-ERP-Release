@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

cd /d "%~dp0"

set "REPO_URL=https://github.com/ryangroe0618-maker/office-erp.git"
set "BRANCH=main"
set "PS_UPDATE=%~dp0업데이트_확인.ps1"
set "EXE_RELEASE_CHECKED=0"

echo.
echo ========================================
echo   Office ERP Windows 실행
echo ========================================
echo.

if exist ".git" (
    where git > nul 2> nul
    if errorlevel 1 (
        echo Git이 설치되어 있지 않아 자동 업데이트를 건너뜁니다.
        echo Git 설치: https://git-scm.com/download/win
        echo.
    ) else (
        echo [Office ERP] GitHub 최신본을 확인합니다...
        git remote set-url origin %REPO_URL% > nul 2> nul
        git fetch origin %BRANCH% > nul 2> nul
        if errorlevel 1 (
            echo 최신본 확인에 실패했습니다. 현재 설치된 버전으로 실행합니다.
            echo.
        ) else (
            for /f "delims=" %%A in ('git rev-parse HEAD') do set "LOCAL_COMMIT=%%A"
            for /f "delims=" %%A in ('git rev-parse origin/%BRANCH%') do set "REMOTE_COMMIT=%%A"

            if /I "!LOCAL_COMMIT!"=="!REMOTE_COMMIT!" (
                echo 이미 최신 버전입니다.
                echo.
            ) else (
                echo 새 버전이 있습니다. 자동 업데이트를 진행합니다...
                git pull --ff-only origin %BRANCH%
                if errorlevel 1 (
                    echo.
                    echo 자동 업데이트에 실패했습니다.
                    echo 변경 충돌이 있거나 네트워크 문제가 있을 수 있습니다.
                    echo 문제 확인 후 다시 실행해 주세요.
                    pause
                    exit /b 1
                )
                echo 업데이트 완료. 최신 버전으로 실행합니다.
                echo.
            )
        )
    )
) else (
    echo Git 저장소가 아닌 설치본입니다. EXE 릴리즈 업데이트를 확인합니다...
    if exist "%PS_UPDATE%" (
        powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_UPDATE%"
        if errorlevel 1 (
            echo.
            echo 업데이트 적용에 실패했습니다. 위 오류와 update_debug.log 를 확인해 주세요.
            pause
            exit /b 1
        )
        set "EXE_RELEASE_CHECKED=1"
        echo.
    ) else (
        echo 업데이트 확인 스크립트를 찾지 못했습니다. 현재 버전으로 실행합니다.
        echo.
    )
)

if exist "%PS_UPDATE%" (
    if "!EXE_RELEASE_CHECKED!"=="0" (
        echo [Office ERP] EXE 릴리즈 업데이트를 확인합니다...
        powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_UPDATE%"
        if errorlevel 1 (
            echo.
            echo 업데이트 적용에 실패했습니다. 위 오류와 update_debug.log 를 확인해 주세요.
            pause
            exit /b 1
        )
        set "EXE_RELEASE_CHECKED=1"
        echo.
    )
)

if exist "%PS_UPDATE%" (
    if not exist "Office_ERP.exe" (
        echo Office_ERP.exe 가 없어 최신 EXE 릴리즈를 확인합니다...
        powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_UPDATE%"
        if errorlevel 1 (
            echo.
            echo 업데이트 적용에 실패했습니다. 위 오류와 update_debug.log 를 확인해 주세요.
            pause
            exit /b 1
        )
        echo.
    )
)

if exist "Office_ERP.exe" (
    echo Office_ERP.exe 를 실행합니다...
    start "" "%~dp0Office_ERP.exe"
    exit /b 0
)

if exist "Office_ERP_실행.bat" (
    echo Office_ERP.exe 가 아직 없습니다.
    echo Python 설치형 실행 파일을 실행합니다...
    echo.
    call "Office_ERP_실행.bat"
    exit /b %errorlevel%
)

echo 실행 파일을 찾지 못했습니다.
echo 압축을 완전히 푼 폴더 안에서 START_WINDOWS.cmd 를 실행해 주세요.
echo.
pause

endlocal
