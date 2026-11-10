$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Proto = "https" + "://"
$Repo = "MaksimTMB/putevka-rps"
$LatestExeUrl = $Proto + "github.com/" + $Repo + "/releases/latest/download/Putevka.exe"
$AppDir = Join-Path $env:LOCALAPPDATA "PutevkaApp"
$DataDir = Join-Path $AppDir "data"
New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
Write-Host "=== АСУ ПУТЁВКА v6.8.0 - Прямое обновление из GitHub ===" -ForegroundColor Cyan
$CurDb = Join-Path $DataDir "putevka_db.json"
if (Test-Path $CurDb) {
    Copy-Item -LiteralPath $CurDb -Destination (Join-Path $DataDir "putevka_db.backup.json") -Force
    Write-Host "[OK] Сохранена резервная копия БД: putevka_db.backup.json" -ForegroundColor Green
}
Get-Process -Name "Putevka*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$TargetExe = Join-Path $AppDir "Putevka.exe"
$DesktopExe = Join-Path ([Environment]::GetFolderPath("Desktop")) "Putevka.exe"
Write-Host "[INFO] Загрузка последнего релиза Putevka.exe из GitHub Releases..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $LatestExeUrl -OutFile $TargetExe -UseBasicParsing
Copy-Item -LiteralPath $TargetExe -Destination $DesktopExe -Force -ErrorAction SilentlyContinue
Write-Host "[OK] Готовый файл Putevka.exe обновлён на Рабочем столе: $DesktopExe" -ForegroundColor Green
Start-Process -FilePath $TargetExe
