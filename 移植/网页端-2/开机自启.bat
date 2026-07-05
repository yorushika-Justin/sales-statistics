@echo off
setlocal enabledelayedexpansion

:: 销售汇总网页端 - 开机自启管理
:: 使用 Windows 任务计划程序实现开机自启

set "TASK_NAME=销售汇总网页端"
set "SCRIPT_DIR=%~dp0"
set "VBS_PATH=%SCRIPT_DIR%启动网页端_静默.vbs"
set "PYTHONW_PATH=D:\Python\pythonw.exe"

:: 检查 pythonw.exe 是否存在
if not exist "%PYTHONW_PATH%" (
    echo [错误] 未找到 pythonw.exe: %PYTHONW_PATH%
    echo 请检查 Python 安装路径。
    pause
    exit /b 1
)

:menu
cls
echo ========================================
echo   销售汇总网页端 - 开机自启管理
echo ========================================
echo.
echo 任务名称: %TASK_NAME%
echo.
echo [1] 注册开机自启
echo [2] 注销开机自启
echo [3] 查看当前状态
echo [0] 退出
echo.
echo ========================================
set /p choice=请选择 (0-3):

if "%choice%"=="1" goto register
if "%choice%"=="2" goto unregister
if "%choice%"=="3" goto status
if "%choice%"=="0" exit /b 0
echo 无效选择。
pause
goto menu

:register
echo.
echo 正在注册开机自启任务...

:: 创建静默启动 vbs 文件（开机时不弹浏览器）
(
echo Set WshShell = CreateObject^("WScript.Shell"^)
echo WshShell.Run "pythonw ""%SCRIPT_DIR%web_server.py"" --silent", 0, False
) > "%VBS_PATH%"

:: 创建任务计划（登录后延迟 30 秒启动，受限用户权限）
schtasks /create /tn "%TASK_NAME%" /tr "wscript.exe \"%VBS_PATH%\"" /sc onlogon /delay 0000:30 /f /rl limited

:: 检查任务是否创建成功
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %errorlevel% equ 0 (
    echo.
    echo [成功] 开机自启已注册！
    echo - 任务名称: %TASK_NAME%
    echo - 延迟时间: 登录后 30 秒
    echo - 启动模式: 静默（不弹浏览器）
    echo.
    echo Edge 浏览器中保存的 http://127.0.0.1:5000 书签可直接使用。
) else (
    echo.
    echo [失败] 注册失败，请尝试以管理员身份运行。
)
echo.
pause
goto menu

:unregister
echo.
echo 正在注销开机自启任务...
schtasks /delete /tn "%TASK_NAME%" /f

if %errorlevel% equ 0 (
    echo [成功] 开机自启已注销！
    if exist "%VBS_PATH%" del "%VBS_PATH%"
) else (
    echo [提示] 任务未找到或已注销。
)
echo.
pause
goto menu

:status
echo.
echo 正在查询任务状态...
schtasks /query /tn "%TASK_NAME%" /fo list 2>nul
if %errorlevel% neq 0 (
    echo [状态] 开机自启未注册。
)
echo.
pause
goto menu
