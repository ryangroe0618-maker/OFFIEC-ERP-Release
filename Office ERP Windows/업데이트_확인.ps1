param(
    [string]$Repo = "ryangroe0618-maker/OFFIEC-ERP-Release",
    [string]$AssetName = ""
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$VersionFile = Join-Path $Root ".office_erp_version"
$ManifestFile = Join-Path $Root ".office_erp_manifest.json"
$LogFile = Join-Path $Root "update_debug.log"
$ExePath = Join-Path $Root "Office_ERP.exe"
$TempRoot = Join-Path $env:TEMP ("office_erp_update_" + [guid]::NewGuid().ToString("N"))
$ZipPath = Join-Path $TempRoot "Office_ERP_Update.zip"
$ManifestPath = Join-Path $TempRoot "update_manifest.json"
$ExtractPath = Join-Path $TempRoot "extract"

function Read-LocalVersion {
    if (Test-Path $VersionFile) {
        return (Get-Content $VersionFile -Raw).Trim()
    }
    return ""
}

function Write-Log([string]$Message) {
    $Line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $Message
    Add-Content -Path $LogFile -Value $Line -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Find-AssetLike($Release, [string]$Pattern) {
    return $Release.assets |
        Where-Object { $_.name -like $Pattern } |
        Sort-Object -Property updated_at -Descending |
        Select-Object -First 1
}

function Get-ManifestChangedFiles($Manifest) {
    if (-not $Manifest -or -not $Manifest.changed_files) {
        return @()
    }
    return @($Manifest.changed_files | ForEach-Object {
        ([string]$_).Trim().Trim('"').Replace('\"', '').Replace('\', '/')
    })
}

function Test-SmallUpdateManifest($Manifest) {
    $ChangedFiles = Get-ManifestChangedFiles $Manifest
    if ($ChangedFiles.Count -eq 0) {
        return $false
    }
    foreach ($Path in $ChangedFiles) {
        if ($Path -notmatch '^(scripts|apps-script)/' -and $Path -notin @("업데이트_확인.ps1", "START_WINDOWS.cmd", "ERP_강제업데이트.cmd", ".github/workflows/build-windows-exe.yml")) {
            return $false
        }
    }
    return $true
}

function Find-AssetByNameOrPattern($Release, [string]$Name, [string]$Pattern) {
    if ($Name) {
        $Asset = $Release.assets | Where-Object { $_.name -eq $Name } | Select-Object -First 1
        if ($Asset) {
            return $Asset
        }
    }
    return Find-AssetLike $Release $Pattern
}

function Find-ScriptByContent([string]$DestinationPath, [string]$Pattern, [string]$Needle) {
    $ScriptsPath = Join-Path $DestinationPath "scripts"
    if (-not (Test-Path $ScriptsPath)) {
        return $null
    }

    $Candidates = @(Get-ChildItem -Path $ScriptsPath -File -Filter $Pattern -ErrorAction SilentlyContinue)
    foreach ($Candidate in $Candidates) {
        try {
            $Text = Get-Content $Candidate.FullName -Raw -Encoding UTF8
            if ($Text -like "*$Needle*") {
                return $Candidate.FullName
            }
        } catch {
        }
    }
    return $null
}

function Resolve-PackageRoot([string]$Path) {
    $Candidates = @()
    $Candidates += Get-Item -Path $Path
    $Candidates += Get-ChildItem -Path $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue

    $Best = $null
    $BestScore = -1
    foreach ($Candidate in $Candidates) {
        $Score = 0
        if (Test-Path (Join-Path $Candidate.FullName "Office_ERP.exe")) { $Score += 2 }
        if (Test-Path (Join-Path $Candidate.FullName "scripts")) { $Score += 2 }
        if (Test-Path (Join-Path $Candidate.FullName "업데이트_확인.ps1")) { $Score += 1 }
        if ($Score -gt $BestScore) {
            $Best = $Candidate.FullName
            $BestScore = $Score
        }
    }

    if ($Best -and $BestScore -gt 0) {
        return $Best
    }

    return $Path
}

function Stop-RunningOfficeErp {
    $Processes = @(Get-Process -Name "Office_ERP" -ErrorAction SilentlyContinue)
    if ($Processes.Count -eq 0) {
        return
    }

    Write-Log "[Office ERP] 실행 중인 Office_ERP.exe를 종료하고 업데이트합니다."
    $Processes | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}

function Copy-PackageContents([string]$SourcePath, [string]$DestinationPath) {
    Stop-RunningOfficeErp

    $Robocopy = Get-Command robocopy.exe -ErrorAction SilentlyContinue
    if ($Robocopy) {
        & robocopy.exe $SourcePath $DestinationPath /E /R:3 /W:2 /XD ".git" ".github" ".venv" "build" "dist" "__pycache__" "배포" /XF ".gitignore" ".office_erp_version" ".office_erp_manifest.json" | Out-Null
        if ($LASTEXITCODE -gt 7) {
            throw "robocopy 복사 실패: exit code $LASTEXITCODE"
        }
        return
    }

    Get-ChildItem -Path $SourcePath -Force | ForEach-Object {
        if ($_.Name -notin @(".git", ".github", ".gitignore", ".venv", "build", "dist", "__pycache__", "배포", ".office_erp_version", ".office_erp_manifest.json") -and (Test-Path $_.FullName)) {
            Copy-Item -Path $_.FullName -Destination $DestinationPath -Recurse -Force
        }
    }
}

function Assert-UpdateApplied([string]$DestinationPath) {
    $LauncherPath = Join-Path $DestinationPath "ERP 업무 런처.py"
    $FullmakePath = Find-ScriptByContent $DestinationPath "fullmake *.py" "2PACX-1vRzG4kGawePQdGxZB9CII2zfEJKV4Vgdp4Ux3MmiXgr9KFHSX00xdOPFQZ_YyxO46lc0Jq-lcA8AuS5"
    $PoizonPath = Find-ScriptByContent $DestinationPath "*poizon*.py" 'split("-", 1)[0]'

    if (-not $FullmakePath) {
        throw "업데이트 검증 실패: fullmake 리스트 업로드.py 파일을 찾을 수 없습니다."
    }

    if (Test-Path $LauncherPath) {
        $LauncherText = Get-Content $LauncherPath -Raw -Encoding UTF8
        if ($LauncherText -notlike "*brander_list_upload*" -or $LauncherText -notlike "*brand 리스트 업로드.py*") {
            throw "업데이트 검증 실패: 런처가 최신 브랜더/풀메이커 구성으로 바뀌지 않았습니다."
        }
    } else {
        Write-Log "[Office ERP] ERP 업무 런처.py가 없는 EXE 설치본입니다. 스크립트 파일 기준으로 검증합니다."
    }

    if ($PoizonPath -and (Test-Path $PoizonPath)) {
        $PoizonText = Get-Content $PoizonPath -Raw -Encoding UTF8
        if ($PoizonText -notlike '*split("-", 1)[0]*') {
            throw "업데이트 검증 실패: 플랫폼 POIZON 품번 정리 수정이 반영되지 않았습니다."
        }
    }

    $FullmakeListText = Get-Content $FullmakePath -Raw -Encoding UTF8
    if ($FullmakeListText -notlike "*2PACX-1vRzG4kGawePQdGxZB9CII2zfEJKV4Vgdp4Ux3MmiXgr9KFHSX00xdOPFQZ_YyxO46lc0Jq-lcA8AuS5*") {
        throw "업데이트 검증 실패: FULLMAKE 리스트 업로드 원본 시트 수정이 반영되지 않았습니다."
    }

    $FullmakeCollectPath = Find-ScriptByContent $DestinationPath "fullmake *.py" '.eq("P")'
    if (-not $FullmakeCollectPath) {
        throw "업데이트 검증 실패: FULLMAKE 출고처 P 기준 수정이 반영되지 않았습니다."
    }
}

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Headers = @{
        "User-Agent" = "Office-ERP-Updater"
        "Accept" = "application/vnd.github+json"
    }
    $Release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" -Headers $Headers -TimeoutSec 20

    $RemoteVersion = [string]$Release.tag_name
    $LocalVersion = Read-LocalVersion
    Write-Log "[Office ERP] 설치 폴더: $Root"
    Write-Log "[Office ERP] 현재 버전: $LocalVersion / 최신 버전: $RemoteVersion"
    $ManifestAsset = Find-AssetLike $Release "Office_ERP_UpdateManifest*.json"
    $Manifest = $null

    New-Item -ItemType Directory -Force -Path $TempRoot, $ExtractPath | Out-Null

    if ($ManifestAsset) {
        Invoke-WebRequest -Uri $ManifestAsset.browser_download_url -OutFile $ManifestPath -Headers @{ "User-Agent" = "Office-ERP-Updater" } -TimeoutSec 60
        $Manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json
        Write-Log "[Office ERP] 업데이트 방식: $($Manifest.update_type)"
    }

    if ((Test-Path $ExePath) -and $LocalVersion -eq $RemoteVersion) {
        Write-Log "[Office ERP] 이미 최신 버전입니다: $RemoteVersion"
        exit 0
    }

    $Asset = $null
    $UpdateType = "exe"
    if ($Manifest -and $Manifest.update_type) {
        $UpdateType = [string]$Manifest.update_type
    }
    if (Test-SmallUpdateManifest $Manifest) {
        $UpdateType = "script"
        Write-Log "[Office ERP] 변경 파일 기준으로 작은 업데이트를 적용합니다."
    }

    if ($AssetName) {
        $Asset = $Release.assets | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1
    }
    if (-not $Asset -and $UpdateType -eq "script") {
        $Asset = Find-AssetByNameOrPattern $Release ([string]$Manifest.script_zip) "Office_ERP_Update*.zip"
    }
    if (-not $Asset) {
        $Asset = Find-AssetByNameOrPattern $Release ([string]$Manifest.exe_zip) "Office_ERP_Windows_EXE*.zip"
        $UpdateType = "exe"
    }

    if (-not $Asset) {
        Write-Log "[Office ERP] 최신 릴리즈에 업데이트 ZIP 파일이 없습니다. 현재 버전으로 실행합니다."
        exit 0
    }

    Write-Log "[Office ERP] 최신 패키지를 다운로드합니다: $RemoteVersion / $($Asset.name) / $UpdateType"
    Invoke-WebRequest -Uri $Asset.browser_download_url -OutFile $ZipPath -Headers @{ "User-Agent" = "Office-ERP-Updater" } -TimeoutSec 600

    Write-Log "[Office ERP] 최신 패키지를 적용합니다..."
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force

    $SourcePath = Resolve-PackageRoot $ExtractPath
    Write-Log "[Office ERP] 패키지 적용 원본: $SourcePath"
    Copy-PackageContents $SourcePath $Root
    Assert-UpdateApplied $Root

    Set-Content -Path $VersionFile -Value $RemoteVersion -Encoding UTF8
    if ($Manifest) {
        Copy-Item -Path $ManifestPath -Destination $ManifestFile -Force
    }
    Write-Log "[Office ERP] 업데이트 완료: $RemoteVersion"
    exit 0
} catch {
    Write-Log "[Office ERP] 자동 EXE 업데이트 확인 실패: $($_.Exception.Message)"
    Write-Log "[Office ERP] 현재 설치된 버전으로 계속 실행합니다."
    exit 1
} finally {
    if (Test-Path $TempRoot) {
        Remove-Item -Path $TempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
