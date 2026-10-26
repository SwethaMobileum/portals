@echo off
cd /d "%~dp0"
echo Starting the live Excel view...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0live-excel.ps1"
echo.
pause
