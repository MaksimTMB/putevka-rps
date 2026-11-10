$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RepoDir = $PSScriptRoot
Set-Location -LiteralPath $RepoDir

Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "  Выгрузка репозитория АСУ ПУТЁВКА (v5.0.0 - v6.7.0) в GitHub" -ForegroundColor Cyan
Write-Host "  Целевой репозиторий: https://github.com/MaksimTMB/putevka-rps" -ForegroundColor Cyan
Write-Host "  Локальная папка:     $RepoDir" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan

function Find-GitExe {
    $cmd = Get-Command git -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { return $cmd.Source }
    $candidates = @(
        "$env:ProgramFiles\Git\cmd\git.exe",
        "$env:ProgramFiles\Git\bin\git.exe",
        "${env:ProgramFiles(x86)}\Git\cmd\git.exe",
        "$env:LocalAppData\Programs\Git\cmd\git.exe"
    )
    $ghDesktop = Get-ChildItem -Path "$env:LocalAppData\GitHubDesktop" -Filter "git.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ghDesktop) { $candidates += $ghDesktop.FullName }
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) { return $c }
    }
    return $null
}

$GitExe = Find-GitExe
if (-not $GitExe) {
    Write-Host "[WARN] Git не найден в PATH. Пробуем установить Git через winget..." -ForegroundColor Yellow
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        & winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        $GitExe = Find-GitExe
    }
}

if (-not $GitExe) {
    Write-Host "[ОШИБКА] На компьютере не установлен Git (git.exe)." -ForegroundColor Red
    Write-Host "Скачайте и установите Git для Windows: https://git-scm.com/download/win" -ForegroundColor Yellow
    Write-Host "После установки снова запустите push_to_github.bat" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] Используется Git: $GitExe" -ForegroundColor Green

$safePath = $RepoDir -replace '\\', '/'
& $GitExe config --global --add safe.directory $safePath 2>$null
& $GitExe config --global --add safe.directory "*" 2>$null
& $GitExe config --global http.postBuffer 524288000 2>$null

if (-not (Test-Path -LiteralPath (Join-Path $RepoDir ".git"))) {
    Write-Host "[INFO] Инициализация локального репозитория .git..." -ForegroundColor Cyan
    & $GitExe init
}

& $GitExe config user.name "MaksimTMB"
& $GitExe config user.email "maksimtmb@users.noreply.github.com"

& $GitExe remote remove origin 2>$null
& $GitExe remote add origin "https://github.com/MaksimTMB/putevka-rps.git"
& $GitExe branch -M main

Write-Host "[INFO] Отправка ветки main (все исходники и папка releases/*.zip) в GitHub..." -ForegroundColor Cyan
Write-Host "       (Если появится окно входа GitHub - подтвердите вход в браузере)" -ForegroundColor Yellow
& $GitExe push -u origin main --force
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ОШИБКА] Не удалось отправить ветку main. Проверьте авторизацию GitHub." -ForegroundColor Red
    exit 1
}

Write-Host "[INFO] Отправка всех тегов версий (v5.0.0 - v6.7.0)..." -ForegroundColor Cyan
& $GitExe push origin --tags --force
& $GitExe push origin refs/tags/v6.7.0 --force

Write-Host ""
Write-Host "==================================================================" -ForegroundColor Green
Write-Host "  УСПЕШНО! Весь код, папка releases/ (14 архивов) и теги v5.0-v6.7" -ForegroundColor Green
Write-Host "  выгружены в https://github.com/MaksimTMB/putevka-rps" -ForegroundColor Green
Write-Host "  GitHub Actions запущен для сборки Putevka_Windows_v6_7.exe!" -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Green
Start-Process "https://github.com/MaksimTMB/putevka-rps"
