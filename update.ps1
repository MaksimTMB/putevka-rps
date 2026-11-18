$ErrorActionPreference = "Continue"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Proto = "https" + "://"

$Repo = "MaksimTMB/putevka-rps"

$LatestExeUrl = $Proto + "github.com/" + $Repo + "/releases/latest/download/Putevka.exe"

$AppDir = Join-Path $env:LOCALAPPDATA "PutevkaApp"

$DataDir = Join-Path $AppDir "data"

New-Item -ItemType Directory -Path $DataDir -Force | Out-Null

Write-Host "=== АСУ ПУТЁВКА v8.5.6 - Прямое обновление из GitHub ===" -ForegroundColor Cyan

$CurDb = Join-Path $DataDir "putevka_db.json"

if (Test-Path $CurDb) { Copy-Item -LiteralPath $CurDb -Destination (Join-Path $DataDir "putevka_db.backup.json") -Force }

Get-Process -Name "Putevka*", "pythonw*", "python*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Get-NetTCPConnection -LocalPort 18080 -ErrorAction SilentlyContinue | ForEach-Object { if ($_.OwningProcess -gt 0) { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } }

Start-Sleep -Milliseconds 700

Remove-Item -Path (Join-Path $AppDir "static") -Recurse -Force -ErrorAction SilentlyContinue

$TargetExe = Join-Path $AppDir "Putevka.exe"

Invoke-WebRequest -Uri $LatestExeUrl -OutFile $TargetExe -UseBasicParsing

# Настоящие ярлыки вместо копии exe: «ПУТЁВКА.lnk» на Рабочем столе и в меню «Пуск».
# Повторный запуск скрипта обновляет ярлык, а не создаёт второй.
$OldDesktopExe = Join-Path ([Environment]::GetFolderPath("Desktop")) "Putevka.exe"
if (Test-Path $OldDesktopExe) { Remove-Item -LiteralPath $OldDesktopExe -Force -ErrorAction SilentlyContinue }

$DesktopLnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "ПУТЁВКА.lnk"
$StartMenuDir = Join-Path ([Environment]::GetFolderPath("Programs")) "АСУ ПУТЁВКА"
New-Item -ItemType Directory -Path $StartMenuDir -Force | Out-Null
$StartMenuLnk = Join-Path $StartMenuDir "ПУТЁВКА.lnk"

$Shell = New-Object -ComObject WScript.Shell
foreach ($LnkPath in @($DesktopLnk, $StartMenuLnk)) {
    $Lnk = $Shell.CreateShortcut($LnkPath)
    $Lnk.TargetPath = $TargetExe
    $Lnk.WorkingDirectory = $AppDir
    $Lnk.IconLocation = "$TargetExe,0"
    $Lnk.Description = "АСУ ПУТЁВКА — путевые листы (форма № 3)"
    $Lnk.Save()
}

Write-Host "[OK] УСПЕШНО! Приложение обновлено до версии v8.5.6" -ForegroundColor Green
Write-Host "Ярлык на Рабочем столе: $DesktopLnk" -ForegroundColor Green
Write-Host "Ярлык в меню «Пуск»: $StartMenuLnk" -ForegroundColor Green

Start-Process -FilePath $TargetExe

