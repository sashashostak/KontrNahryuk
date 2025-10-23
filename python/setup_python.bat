@echo off
echo ========================================
echo Python Installation Test
echo ========================================
echo.

echo [1/3] Checking Python...
python --version
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python not found!
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo.
echo [2/3] Installing openpyxl...
cd /d "%~dp0"
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install openpyxl
    pause
    exit /b 1
)

echo.
echo [3/3] Testing openpyxl...
python -c "import openpyxl; print('openpyxl version:', openpyxl.__version__)"
if %errorlevel% neq 0 (
    echo.
    echo ERROR: openpyxl import failed
    pause
    exit /b 1
)

echo.
echo ========================================
echo SUCCESS! Python is ready to use
echo ========================================
echo.
pause
