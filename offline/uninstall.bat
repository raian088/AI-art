@echo off
rem AI Drawing - remove shortcuts (Desktop / Startup)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
echo.
pause
