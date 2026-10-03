@echo off
cd /d %~dp0
where node >nul 2>nul
if errorlevel 1 (
  echo [FAIL] Node.js not found in PATH. Please install Node.js first.
  pause
  exit /b 1
)
node collect.js
set EXITCODE=%errorlevel%
echo.
if %EXITCODE%==0 (
  echo [OK] Done. Excel in download dir, PNG in sales folder, brand image auto-opened.
) else (
  echo [FAIL] exit code %EXITCODE%. See logs\ for details and screenshots.
)
pause
exit /b %EXITCODE%
