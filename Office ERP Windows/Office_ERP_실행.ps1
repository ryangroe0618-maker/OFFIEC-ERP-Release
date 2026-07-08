$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Set-Location -LiteralPath $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "[Office ERP] 가상환경을 준비합니다..."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -3 -m venv .venv
    } else {
        $global:LASTEXITCODE = 1
    }
    if ($LASTEXITCODE -ne 0 -and (Get-Command python -ErrorAction SilentlyContinue)) {
        python -m venv .venv
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Python 3.11 이상을 찾지 못했습니다. https://www.python.org/downloads/windows/ 에서 Python을 설치해 주세요."
    }
}

Write-Host "[Office ERP] 필요한 패키지를 확인합니다..."
& ".venv\Scripts\python.exe" -m pip install --upgrade pip
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt

Write-Host "[Office ERP] 런처를 실행합니다..."
& ".venv\Scripts\python.exe" "ERP 업무 런처.py"
