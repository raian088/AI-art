@echo off
rem AI Drawing - remove Windows kiosk shortcuts
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
echo.
pause
