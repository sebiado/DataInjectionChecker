@echo off
echo ==========================================
echo Data Injection Checker (ORG 200 ANA detect)
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python is not installed or not in PATH.
    pause
    exit /b
)

echo Running comparison...
python check_injection.py

echo.
echo Done.
pause
``