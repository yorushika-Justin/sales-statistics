@echo off
setlocal enabledelayedexpansion

:: Sales Summary Web - Auto Start Manager
:: Uses Windows Task Scheduler for auto start on login

set "TASK_NAME=SalesSummaryWeb"
set "SCRIPT_DIR=%~dp0"
set "VBS_PATH=%SCRIPT_DIR%autostart_silent.vbs"
set "PYTHONW_PATH=D:\Python\pythonw.exe"

:: Check if pythonw.exe exists
if not exist "%PYTHONW_PATH%" (
    echo [ERROR] pythonw.exe not found: %PYTHONW_PATH%
    echo Please check Python installation path.
    pause
    exit /b 1
)

:menu
cls
echo ========================================
echo   Sales Summary Web - Auto Start Manager
echo ========================================
echo.
echo Task name: %TASK_NAME%
echo.
echo [1] Register auto start
echo [2] Unregister auto start
echo [3] Check status
echo [0] Exit
echo.
echo ========================================
set /p choice=Select (0-3):

if "%choice%"=="1" goto register
if "%choice%"=="2" goto unregister
if "%choice%"=="3" goto status
if "%choice%"=="0" exit /b 0
echo Invalid choice.
pause
goto menu

:register
echo.
echo Registering auto start task...

:: Create silent vbs file (no browser popup on startup)
(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run "pythonw ""%SCRIPT_DIR%web_server.py"" --silent", 0, False
) > "%VBS_PATH%"

:: Create task (start 30s after login, limited user privilege)
schtasks /create /tn "%TASK_NAME%" /tr "wscript.exe \"%VBS_PATH%\"" /sc onlogon /delay 0000:30 /f /rl limited

:: Check if task was created successfully
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo [SUCCESS] Auto start registered!
    echo - Task name: %TASK_NAME%
    echo - Delay: 30 seconds after login
    echo - Mode: silent (no browser popup)
    echo.
    echo Edge browser saved tab http://127.0.0.1:5000 can be used directly.
) else (
    echo.
    echo [FAILED] Registration failed. Try running as Administrator.
)
echo.
pause
goto menu

:unregister
echo.
echo Unregistering auto start task...
schtasks /delete /tn "%TASK_NAME%" /f

if %errorlevel% equ 0 (
    echo [SUCCESS] Auto start unregistered!
    if exist "%VBS_PATH%" del "%VBS_PATH%"
) else (
    echo [INFO] Task not found or already unregistered.
)
echo.
pause
goto menu

:status
echo.
echo Checking task status...
schtasks /query /tn "%TASK_NAME%" /fo list 2>nul
if %errorlevel% neq 0 (
    echo [STATUS] Auto start not registered.
)
echo.
pause
goto menu
