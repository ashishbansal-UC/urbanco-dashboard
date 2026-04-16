@echo off
cd /d "%~dp0"
python fetch_data.py
echo.
echo Dashboard data refreshed! Open dashboard.html to view.
pause
