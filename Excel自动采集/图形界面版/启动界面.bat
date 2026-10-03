@echo off
cd /d %~dp0
where node >nul 2>nul
if errorlevel 1 (
  echo [FAIL] Node.js not found in PATH. Please install Node.js first.
  pause
  exit /b 1
)
node server.js
set EXITCODE=%errorlevel%
if not %EXITCODE%==0 (
  echo [FAIL] server exited with code %EXITCODE%
  pause
)
exit /b %EXITCODE%
