@echo off
rem AI Drawing - Windows kiosk setup (Chrome shortcut to the server URL)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" %*
echo.
pause
