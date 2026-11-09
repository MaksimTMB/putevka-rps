$ErrorActionPreference = "Continue"
$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host ">>> Остановка старого процесса server.py (порт 8080)..." -ForegroundColor Yellow
Get-NetTCPConnection -LocalPort 8080 -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path "$ProjectRoot\venv")) {
    python -m venv "$ProjectRoot\venv"
}
& "$ProjectRoot\venv\Scripts\pip.exe" install -r "$ProjectRoot\requirements.txt"
Copy-Item "$ProjectRoot\data\app_data.json" "$ProjectRoot\data\putevka_db.json" -Force

Start-Process -FilePath "$ProjectRoot\venv\Scripts\python.exe" -ArgumentList "$ProjectRoot\server.py" -WorkingDirectory $ProjectRoot -WindowStyle Hidden
Write-Host ">>> Сервер ПУТЁВКА v6.5.0 успешно обновлен и перезапущен на порту 8080!" -ForegroundColor Green
