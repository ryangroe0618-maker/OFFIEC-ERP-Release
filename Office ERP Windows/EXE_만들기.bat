@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [Office ERP] 빌드용 가상환경을 준비합니다...
    py -3 -m venv .venv
    if errorlevel 1 (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo.
        echo Python을 찾지 못했습니다.
        echo exe를 만드는 PC에는 Python 3.11 이상이 한 번 필요합니다.
        echo https://www.python.org/downloads/windows/
        pause
        exit /b 1
    )
)

echo [Office ERP] 빌드 패키지를 설치합니다...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 (
    echo.
    echo 빌드 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
)

echo [Office ERP] 단독 실행 파일을 생성합니다...
".venv\Scripts\python.exe" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --name Office_ERP ^
    --hidden-import office_erp_bundle_imports ^
    --collect-all PySide6 ^
    --collect-all pandas ^
    --collect-all numpy ^
    --collect-all openpyxl ^
    --collect-all gspread ^
    --collect-all google ^
    --collect-all google_auth_oauthlib ^
    --collect-all pypdf ^
    --collect-all pymupdf ^
    --collect-all fitz ^
    "ERP 업무 런처.py"

if errorlevel 1 (
    echo.
    echo exe 생성에 실패했습니다.
    pause
    exit /b 1
)

copy /Y "dist\Office_ERP.exe" "Office_ERP.exe" > nul

echo [Office ERP] 배포용 폴더를 정리합니다...
if exist "배포" rmdir /S /Q "배포"
mkdir "배포\Office ERP Windows"
mkdir "배포\Office ERP Windows EXE"

robocopy "%~dp0" "%~dp0배포\Office ERP Windows" /E ^
    /XD ".venv" "build" "dist" "__pycache__" ".git" "배포" ^
    /XF "Office_ERP.exe" "Office_ERP_Windows.zip" "Office_ERP_Windows_EXE.zip" "*.spec" > nul

if errorlevel 8 (
    echo.
    echo 소스 배포용 폴더 복사에 실패했습니다.
    pause
    exit /b 1
)

robocopy "%~dp0" "%~dp0배포\Office ERP Windows EXE" /E ^
    /XD ".venv" "build" "dist" "__pycache__" ".git" "배포" ^
    /XF "Office_ERP_Windows.zip" "Office_ERP_Windows_EXE.zip" "*.spec" > nul

if errorlevel 8 (
    echo.
    echo EXE 배포용 폴더 복사에 실패했습니다.
    pause
    exit /b 1
)

if exist "Office_ERP_Windows.zip" del /Q "Office_ERP_Windows.zip"
if exist "Office_ERP_Windows_EXE.zip" del /Q "Office_ERP_Windows_EXE.zip"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Compress-Archive -Path '배포\Office ERP Windows' -DestinationPath 'Office_ERP_Windows.zip' -Force"

if errorlevel 1 (
    echo.
    echo 소스 ZIP 생성에 실패했습니다. 배포 폴더는 생성되어 있습니다: 배포\Office ERP Windows
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Compress-Archive -Path '배포\Office ERP Windows EXE' -DestinationPath 'Office_ERP_Windows_EXE.zip' -Force"

if errorlevel 1 (
    echo.
    echo EXE ZIP 생성에 실패했습니다. 배포 폴더는 생성되어 있습니다: 배포\Office ERP Windows EXE
    pause
    exit /b 1
)

echo.
echo 완료: Office_ERP.exe 생성됨
echo 완료: Office_ERP_Windows.zip 생성됨 ^(작은 소스/Python 실행용, EXE 제외^)
echo 완료: Office_ERP_Windows_EXE.zip 생성됨 ^(큰 단독 실행 EXE 포함^)
echo.
echo 자동 업데이트는 ZIP이 아니라 git clone 설치 후 START_WINDOWS.cmd 실행을 권장합니다.
echo Python 없는 PC에는 Office_ERP_Windows_EXE.zip 압축을 풀고 Office_ERP.exe 를 실행하면 됩니다.
pause

endlocal
