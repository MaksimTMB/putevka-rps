param(
    [switch]$KeepDb = $false,
    [int]$Port = 8080
)

$ErrorActionPreference = "Stop"
$AppVersion = "v6.6.0"
$ProjectRoot = if (Test-Path "$PSScriptRoot\server.py") { $PSScriptRoot } else { Split-Path -Parent $PSScriptRoot }
Set-Location $ProjectRoot

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  АСУ «ПУТЁВКА» $AppVersion — Полная проверка окружения и запуск (Windows)" -ForegroundColor Cyan
Write-Host "  Рабочая директория: $ProjectRoot" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

# 1. Проверка наличия Python 3 (и автоустановка через winget при отсутствии)
$PyCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $PyCmd) {
    Write-Host "[WARN] Python не найден в PATH. Попытка автоматической установки Python 3.11 через winget..." -ForegroundColor Yellow
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget install -e --id Python.Python.3.11 --scope machine --accept-package-agreements --accept-source-agreements
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    } else {
        Write-Error "Python 3 не установлен. Скачайте установщик с https://www.python.org/downloads/windows/ (отметьте галочку 'Add python.exe to PATH')."
        exit 1
    }
}

$PyVer = & python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"
Write-Host "[ OK ] Python проверен: версия $PyVer" -ForegroundColor Green

# 2. Проверка обязательных файлов проекта
$RequiredFiles = @("server.py", "requirements.txt", "static\index.html", "templates_xls\template_0345001.xlsx", "data\app_data.json")
foreach ($rf in $RequiredFiles) {
    $fullPath = Join-Path $ProjectRoot $rf
    if (-not (Test-Path $fullPath)) {
        Write-Error "Отсутствует обязательный файл проекта: $fullPath"
        exit 1
    }
}
Write-Host "[ OK ] Все обязательные файлы проекта проверены." -ForegroundColor Green

# 3. Инициализация или проверка базы данных
$DbFile = Join-Path $ProjectRoot "data\putevka_db.json"
$InitFile = Join-Path $ProjectRoot "data\app_data.json"
if (-not $KeepDb -or -not (Test-Path $DbFile)) {
    Copy-Item $InitFile $DbFile -Force
    Write-Host "[ OK ] База данных инициализирована эталоном релиза $AppVersion." -ForegroundColor Green
} else {
    Write-Host "[ OK ] Сохранена текущая база данных data\putevka_db.json (-KeepDb)." -ForegroundColor Green
}

# 4. Создание/проверка виртуального окружения venv и установка пакетов
$VenvPy = Join-Path $ProjectRoot "venv\Scripts\python.exe"
$VenvPip = Join-Path $ProjectRoot "venv\Scripts\pip.exe"
if (-not (Test-Path $VenvPy)) {
    Write-Host "[INFO] Создание виртуального окружения venv..." -ForegroundColor Cyan
    & python -m venv "$ProjectRoot\venv"
}

Write-Host "[INFO] Проверка и установка зависимостей (fastapi, uvicorn, openpyxl, pydantic)..." -ForegroundColor Cyan
& $VenvPy -m pip install --upgrade pip setuptools wheel | Out-Null
& $VenvPip install -r "$ProjectRoot\requirements.txt"

& $VenvPy -c "import fastapi, uvicorn, openpyxl, pydantic, server; assert server.APP_VERSION == '$AppVersion'"
Write-Host "[ OK ] Все библиотеки и модуль server.py прошли предстартовую проверку." -ForegroundColor Green

# 5. Освобождение порта 8080 и открытие правила брандмауэра Windows
Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue
}
try {
    New-NetFirewallRule -DisplayName "ASU Putevka $Port" -Direction Inbound -LocalPort $Port -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null
} catch {}

# 6. Запуск сервера в фоновом процессе и проверка HTTP-ответа
Write-Host "[INFO] Запуск сервера ПУТЁВКА $AppVersion на порту $Port..." -ForegroundColor Cyan
$Proc = Start-Process -FilePath $VenvPy -ArgumentList "server.py" -WorkingDirectory $ProjectRoot -PassThru -WindowStyle Hidden

$HealthOk = $false
for ($i = 0; $i -lt 6; $i++) {
    Start-Sleep -Seconds 1
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/state" -TimeoutSec 3
        if ($null -ne $resp.profiles) {
            $HealthOk = $true
            break
        }
    } catch {}
}

if (-not $HealthOk) {
    Write-Error "Сервер не ответил на контрольный запрос http://127.0.0.1:$Port/api/state"
    exit 1
}

Write-Host "==================================================================" -ForegroundColor Green
Write-Host "  >>> УСПЕШНО! АСУ «ПУТЁВКА» $AppVersion запущена (PID: $($Proc.Id))" -ForegroundColor Green
Write-Host "  Адрес в браузере: http://localhost:$Port" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Green
Start-Process "http://localhost:$Port"
