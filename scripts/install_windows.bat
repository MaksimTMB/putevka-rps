@echo off
pushd "%~dp0.."
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\install_windows.ps1" %*
popd
pause
