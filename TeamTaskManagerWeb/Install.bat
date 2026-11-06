@echo off
cd /d "%~dp0"
echo ========================================================
echo   TeamFlow - ניהול משימות צוות
echo   התקנה מהירה
echo ========================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-TeamTaskManager.ps1"
echo.
pause
