@echo off
title GeometryHome - Workforce & Meal Management System v2
color 1F
echo.
echo =========================================================
echo   GeometryHome - Workforce ^& Meal Management v2.0
echo =========================================================
echo.
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install from https://python.org
    echo Make sure to check "Add Python to PATH" during install.
    pause & exit /b 1
)
echo [1/3] Python found: & python --version
echo [2/3] Installing dependencies (first run may take a minute)...
pip install fastapi uvicorn python-multipart openpyxl pandas --quiet --disable-pip-version-check
echo [3/3] Launching application...
echo.
echo =========================================================
echo   App URL   : http://localhost:8000
echo   Meal Lookup: http://localhost:8000/lookup
echo   Press Ctrl+C to stop
echo =========================================================
echo.
start "" /B cmd /C "timeout /t 2 >nul && start http://localhost:8000"
python app.py
echo. & echo Server stopped. & pause
