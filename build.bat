@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo   AI Drawing - build
echo ============================================================
echo   Working dir: %CD%
echo.

set "PYCMD="
where py >nul 2>&1
if %errorlevel%==0 set "PYCMD=py -3"
if "%PYCMD%"=="" (
  where python >nul 2>&1
  if %errorlevel%==0 set "PYCMD=python"
)

if "%PYCMD%"=="" (
  echo [ERROR] Python was not found.
  pause
  exit /b 1
)

if not exist "kiosk.py" (
  echo [ERROR] kiosk.py was not found.
  pause
  exit /b 1
)

if not exist "desktopaap.py" (
  echo [ERROR] desktopaap.py was not found.
  pause
  exit /b 1
)

if not exist "index.html" (
  echo [ERROR] index.html was not found.
  pause
  exit /b 1
)

if not exist "admin.html" (
  echo [ERROR] admin.html was not found.
  pause
  exit /b 1
)

if not exist "BGM\BGM1.mp3" (
  echo [ERROR] BGM\BGM1.mp3 was not found.
  pause
  exit /b 1
)

if not exist "BGM\BGM2.mp3" (
  echo [ERROR] BGM\BGM2.mp3 was not found.
  pause
  exit /b 1
)

if not exist "assets\logo-manoi.png" (
  echo [ERROR] assets\logo-manoi.png was not found.
  pause
  exit /b 1
)

if not exist "assets\logo-lubo.png" (
  echo [ERROR] assets\logo-lubo.png was not found.
  pause
  exit /b 1
)

echo [1/3] Installing required packages...
%PYCMD% -m pip install --no-cache-dir --upgrade pyinstaller flask flask-cors requests python-dotenv pillow pywin32 openpyxl certifi "pythonnet>=3.0.5" pywebview
if errorlevel 1 (
  echo [ERROR] Package install failed.
  pause
  exit /b 1
)

echo.
echo [2/3] Cleaning old normal build...
if exist "build" rmdir /s /q "build"
if exist "dist\AIDrawing" rmdir /s /q "dist\AIDrawing"

echo.
echo [3/3] Building dist\AIDrawing\AIDrawing.exe ...
%PYCMD% -m PyInstaller ^
  --noconfirm ^
  --onedir ^
  --windowed ^
  --name "AIDrawing" ^
  --add-data "index.html;." ^
  --add-data "admin.html;." ^
  --add-data "BGM;BGM" ^
  --add-data "assets;assets" ^
  --hidden-import win32print ^
  --hidden-import win32ui ^
  --hidden-import win32con ^
  --hidden-import win32gui ^
  --hidden-import pywintypes ^
  --hidden-import PIL.ImageWin ^
  --hidden-import PIL.ImageFont ^
  --hidden-import PIL.ImageDraw ^
  --hidden-import flask ^
  --hidden-import flask_cors ^
  --hidden-import werkzeug ^
  --hidden-import jinja2 ^
  --hidden-import click ^
  --hidden-import itsdangerous ^
  --hidden-import blinker ^
  --hidden-import markupsafe ^
  --hidden-import dotenv ^
  --hidden-import requests ^
  --hidden-import urllib3 ^
  --hidden-import certifi ^
  --hidden-import idna ^
  --hidden-import charset_normalizer ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  --hidden-import et_xmlfile ^
  --hidden-import webview ^
  --hidden-import webview.platforms.edgechromium ^
  --hidden-import webview.platforms.winforms ^
  --hidden-import clr_loader ^
  --hidden-import pythonnet ^
  --hidden-import kiosk ^
  --collect-submodules win32 ^
  --collect-submodules flask ^
  --collect-submodules werkzeug ^
  --collect-submodules requests ^
  --collect-submodules urllib3 ^
  --collect-submodules charset_normalizer ^
  --collect-submodules openpyxl ^
  --collect-submodules webview ^
  --collect-data certifi ^
  --collect-data webview ^
  desktopaap.py

if errorlevel 1 (
  echo [ERROR] Build failed.
  pause
  exit /b 1
)

if exist ".env" (
  copy /y ".env" "dist\AIDrawing\.env" >nul
)
if exist ".env.example" copy /y ".env.example" "dist\AIDrawing\.env.example" >nul
if exist "setup" xcopy /E /I /Y /Q "setup" "dist\AIDrawing\setup" >nul

echo.
echo [OK] Build complete: dist\AIDrawing\AIDrawing.exe
echo      Next: run dist\AIDrawing\setup\install.bat to install to Desktop and register auto-start.
pause
endlocal
