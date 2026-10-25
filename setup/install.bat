@echo off
rem AI Drawing - install to Desktop + create shortcuts (Desktop / Startup)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
echo.
pause
