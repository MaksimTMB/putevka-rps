@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

echo ==================================================
echo   Установка АСУ ПУТЁВКА v6.5.0 для Windows
echo   Рабочая папка: %CD%
echo ==================================================

where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ОШИБКА] Python не найден в PATH! Установите Python 3.10+ с python.org с галочкой "Add Python to PATH".
    pause
    exit /b 1
)

if not exist "venv" (
    echo Создание виртуального окружения venv...
    python -m venv venv
)

call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

if not exist "data\putevka_db.json" (
    copy /Y "data\app_data.json" "data\putevka_db.json"
)

echo Запуск сервера ПУТЁВКА v6.5.0 на http://localhost:8080 ...
start "" http://localhost:8080
python server.py
