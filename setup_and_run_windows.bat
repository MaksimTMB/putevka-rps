@echo off
pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "setup_and_run_windows.ps1" %*
popd
pause
