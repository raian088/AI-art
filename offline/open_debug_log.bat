@echo off
setlocal
cd /d "%~dp0.."

set "DEBUG_URL=http://127.0.0.1:5000/debug"

echo ============================================================
echo   AI Drawing - debug log
echo ============================================================
echo.

if exist "debug.xlsx" (
  echo [OK] debug.xlsx exists: %CD%\debug.xlsx
  echo      Excel locks this file, so this bat does not open it automatically.
) else if exist "dist\AIDrawing\debug.xlsx" (
  echo [OK] debug.xlsx exists: %CD%\dist\AIDrawing\debug.xlsx
  echo      Excel locks this file, so this bat does not open it automatically.
) else (
  echo [INFO] debug.xlsx was not found yet.
  echo        It will be created after the app writes logs.
)

echo.
echo [1/3] Checking realtime debug server...
call :check_debug
if not errorlevel 1 goto open_debug

echo [2/3] Debug server is not running. Starting debug-only server...
call :start_debug_server

echo [3/3] Waiting for debug server...
for /l %%i in (1,1,20) do (
  call :check_debug
  if not errorlevel 1 goto open_debug
  timeout /t 1 >nul
)

echo.
echo [ERROR] Could not open the realtime debug page.
echo         The app server did not start on %DEBUG_URL%.
echo.
echo [INFO] Opening the logs folder instead.
if exist "logs" start "" "logs"
pause
exit /b 1

:open_debug
echo [OK] Opening realtime debug page
echo      %DEBUG_URL%
start "" "%DEBUG_URL%"
timeout /t 4 >nul
exit /b 0

:check_debug
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:5000/api/debug/data' -TimeoutSec 2; if ($r.StatusCode -ge 200 -and $r.StatusCode -lt 500 -and $r.Content.Contains('xlsx_path') -and $r.Content.Contains('events')) { exit 0 } } catch { }; exit 1"
exit /b %errorlevel%

:start_debug_server
set "PYCMD="
where py >nul 2>&1
if not errorlevel 1 set "PYCMD=py -3"
if "%PYCMD%"=="" where python >nul 2>&1
if "%PYCMD%"=="" if not errorlevel 1 set "PYCMD=python"

if not "%PYCMD%"=="" if exist "desktopaap.py" (
  start "AI Drawing Debug Server" /min %PYCMD% "%CD%\desktopaap.py" --debug-only
  exit /b 0
)

if exist "dist\AIDrawing\AIDrawing.exe" (
  start "AI Drawing Debug Server" /min "dist\AIDrawing\AIDrawing.exe" --debug-only
  exit /b 0
)

echo [WARN] Could not find Python or dist\AIDrawing\AIDrawing.exe.
exit /b 1
