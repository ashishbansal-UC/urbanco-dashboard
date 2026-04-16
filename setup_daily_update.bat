@echo off
REM ============================================================
REM UrbanCo Dashboard - Daily Auto-Update Setup
REM Creates a Windows Task Scheduler task to run fetch_data.py
REM daily at 6:30 PM IST (after market close)
REM ============================================================

set SCRIPT_DIR=%~dp0
set PYTHON_PATH=python
set TASK_NAME=UrbanCo_Dashboard_Update

echo.
echo UrbanCo Dashboard - Setting up daily auto-update
echo ================================================
echo.
echo This will create a Windows scheduled task to refresh
echo dashboard data daily at 6:30 PM (after market close).
echo.

REM Create the scheduled task
schtasks /create /tn "%TASK_NAME%" /tr "\"%PYTHON_PATH%\" \"%SCRIPT_DIR%fetch_data.py\"" /sc daily /st 18:30 /f

if %errorlevel% equ 0 (
    echo.
    echo SUCCESS! Daily update scheduled at 6:30 PM.
    echo Task name: %TASK_NAME%
    echo.
    echo To modify: schtasks /change /tn "%TASK_NAME%" /st HH:MM
    echo To delete: schtasks /delete /tn "%TASK_NAME%" /f
    echo To run now: schtasks /run /tn "%TASK_NAME%"
) else (
    echo.
    echo FAILED. Try running this script as Administrator.
)

echo.
pause
