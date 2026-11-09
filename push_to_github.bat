@echo off
chcp 65001 >nul
cd /d "%~dp0"
git remote remove origin 2>nul
git remote add origin https://github.com/MaksimTMB/putevka-rps.git
git branch -M main
git push -u origin main --tags --force
echo.
echo >>> Весь код, коммиты и теги (v5.0.0 - v6.6.0) успешно выгружены в https://github.com/MaksimTMB/putevka-rps.git!
pause
