$ErrorActionPreference = "Continue"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Proto = "https" + "://"

$Repo = "MaksimTMB/putevka-rps"

$LatestExeUrl = $Proto + "github.com/" + $Repo + "/releases/latest/download/Putevka.exe"

$AppDir = Join-Path $env:LOCALAPPDATA "PutevkaApp"

$DataDir = Join-Path $AppDir "data"

New-Item -ItemType Directory -Path $DataDir -Force | Out-Null

Write-Host "=== АСУ ПУТЁВКА v7.1.1 - Прямое обновление из GitHub ===" -ForegroundColor Cyan

$CurDb = Join-Path $DataDir "putevka_db.json"

if (Test-Path $CurDb) { Copy-Item -LiteralPath $CurDb -Destination (Join-Path $DataDir "putevka_db.backup.json") -Force }

Get-Process -Name "Putevka*", "pythonw*", "python*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Get-NetTCPConnection -LocalPort 18080 -ErrorAction SilentlyContinue | ForEach-Object { if ($_.OwningProcess -gt 0) { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } }

Start-Sleep -Milliseconds 700

Remove-Item -Path (Join-Path $AppDir "static") -Recurse -Force -ErrorAction SilentlyContinue

$TargetExe = Join-Path $AppDir "Putevka.exe"

$DesktopExe = Join-Path ([Environment]::GetFolderPath("Desktop")) "Putevka.exe"

Invoke-WebRequest -Uri $LatestExeUrl -OutFile $TargetExe -UseBasicParsing

Copy-Item -LiteralPath $TargetExe -Destination $DesktopExe -Force -ErrorAction SilentlyContinue

Write-Host "[OK] УСПЕШНО! Приложение обновлено до версии v7.1.1: $DesktopExe" -ForegroundColor Green

Start-Process -FilePath $TargetExe

