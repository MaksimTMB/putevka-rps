@echo off
pushd "%~dp0.."
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\update_windows.ps1" %*
popd
pause
