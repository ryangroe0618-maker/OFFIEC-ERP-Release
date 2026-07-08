@echo off
chcp 65001 > nul
setlocal

cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$repo='ryangroe0618-maker/OFFIEC-ERP-Release';" ^
  "$root=(Get-Location).Path;" ^
  "$tmp=Join-Path $env:TEMP ('office_erp_force_update_' + [guid]::NewGuid().ToString('N'));" ^
  "$zip=Join-Path $tmp 'Office_ERP_Windows_EXE.zip';" ^
  "$extract=Join-Path $tmp 'extract';" ^
  "function Log($m){ Write-Host ('[Office ERP] ' + $m) };" ^
  "try {" ^
  "  [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;" ^
  "  $headers=@{'User-Agent'='Office-ERP-Force-Updater';'Accept'='application/vnd.github+json'};" ^
  "  Log '최신 릴리즈를 확인합니다...';" ^
  "  $release=Invoke-RestMethod -Uri ('https://api.github.com/repos/' + $repo + '/releases/latest') -Headers $headers -TimeoutSec 30;" ^
  "  $asset=$release.assets | Where-Object { $_.name -like 'Office_ERP_Windows_EXE*.zip' } | Sort-Object updated_at -Descending | Select-Object -First 1;" ^
  "  if(-not $asset){ throw '최신 EXE ZIP 파일을 찾을 수 없습니다.' };" ^
  "  Log ('다운로드: ' + $release.tag_name + ' / ' + $asset.name);" ^
  "  New-Item -ItemType Directory -Force -Path $tmp,$extract | Out-Null;" ^
  "  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -Headers @{'User-Agent'='Office-ERP-Force-Updater'} -TimeoutSec 900;" ^
  "  Log '압축을 해제합니다...';" ^
  "  Expand-Archive -Path $zip -DestinationPath $extract -Force;" ^
  "  $candidates=@(Get-Item $extract) + @(Get-ChildItem -Path $extract -Directory -Recurse -Force);" ^
  "  $source=$null; $best=-1;" ^
  "  foreach($c in $candidates){" ^
  "    $score=0;" ^
  "    if(Test-Path (Join-Path $c.FullName 'Office_ERP.exe')){ $score+=2 };" ^
  "    if(Test-Path (Join-Path $c.FullName 'scripts')){ $score+=2 };" ^
  "    if(Test-Path (Join-Path $c.FullName '업데이트_확인.ps1')){ $score+=1 };" ^
  "    if($score -gt $best){ $source=$c.FullName; $best=$score }" ^
  "  };" ^
  "  if(-not $source -or $best -le 0){ throw '압축 안에서 ERP 실행 폴더를 찾지 못했습니다.' };" ^
  "  Log ('적용 원본: ' + $source);" ^
  "  Get-Process -Name 'Office_ERP' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue;" ^
  "  Start-Sleep -Seconds 2;" ^
  "  $skipDirs=@('.git','.github','.venv','build','dist','__pycache__','배포');" ^
  "  $skipFiles=@('.gitignore','.office_erp_version','.office_erp_manifest.json');" ^
  "  Get-ChildItem -LiteralPath $source -Force | ForEach-Object {" ^
  "    if($_.PSIsContainer -and ($skipDirs -contains $_.Name)){ return };" ^
  "    if((-not $_.PSIsContainer) -and ($skipFiles -contains $_.Name)){ return };" ^
  "    if(-not (Test-Path -LiteralPath $_.FullName)){ return };" ^
  "    Copy-Item -LiteralPath $_.FullName -Destination $root -Recurse -Force -ErrorAction Stop;" ^
  "  };" ^
  "  $launcher=Join-Path $root 'ERP 업무 런처.py';" ^
  "  if(-not (Test-Path $launcher)){ throw 'ERP 업무 런처.py 복사 검증 실패' };" ^
  "  $text=Get-Content $launcher -Raw -Encoding UTF8;" ^
  "  if($text -notlike '*brander_list_upload*' -or $text -notlike '*brand 리스트 업로드.py*'){ throw '최신 런처 내용 검증 실패' };" ^
  "  Set-Content -Path (Join-Path $root '.office_erp_version') -Value ([string]$release.tag_name) -Encoding UTF8;" ^
  "  Log ('강제 업데이트 완료: ' + $release.tag_name);" ^
  "  Log 'START_WINDOWS.cmd로 다시 실행해 주세요.';" ^
  "} catch {" ^
  "  Write-Host ('[Office ERP] 강제 업데이트 실패: ' + $_.Exception.Message);" ^
  "  exit 1;" ^
  "} finally {" ^
  "  if(Test-Path $tmp){ Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue }" ^
  "}"

echo.
pause

endlocal
