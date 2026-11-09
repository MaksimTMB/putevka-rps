$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Установка АСУ ПУТЁВКА v6.5.0 (Windows PowerShell)" -ForegroundColor Cyan
Write-Host "  Директория: $ProjectRoot" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python 3 не найден в PATH. Установите Python 3.10+ с https://www.python.org/downloads/windows/"
    exit 1
}

if (-not (Test-Path "$ProjectRoot\venv")) {
    python -m venv "$ProjectRoot\venv"
}

& "$ProjectRoot\venv\Scripts\python.exe" -m pip install --upgrade pip
& "$ProjectRoot\venv\Scripts\pip.exe" install -r "$ProjectRoot\requirements.txt"

if (-not (Test-Path "$ProjectRoot\data\putevka_db.json")) {
    Copy-Item "$ProjectRoot\data\app_data.json" "$ProjectRoot\data\putevka_db.json" -Force
}

try {
    New-NetFirewallRule -DisplayName "ASU Putevka 8080" -Direction Inbound -LocalPort 8080 -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null
} catch {}

Write-Host ">>> Запуск сервера ПУТЁВКА v6.5.0 на http://localhost:8080 ..." -ForegroundColor Green
Start-Process "http://localhost:8080"
& "$ProjectRoot\venv\Scripts\python.exe" "$ProjectRoot\server.py"
