@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

echo ==================================================
echo   Обновление АСУ ПУТЁВКА до v6.5.0 (Windows)
echo ==================================================

taskkill /F /IM python.exe /FI "WINDOWTITLE eq *server.py*" 2>nul

if not exist "venv" (
    python -m venv venv
)
call venv\Scripts\activate.bat
pip install -r requirements.txt

copy /Y "data\app_data.json" "data\putevka_db.json"

echo Перезапуск сервера на порту 8080...
start "PutevkaServer" cmd /c "call venv\Scripts\activate.bat && python server.py"
echo Готово! Обновите страницу в браузере (Ctrl+F5): http://localhost:8080
