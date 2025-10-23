@echo off
chcp 65001 >nul
echo ========================================
echo 🐍 Python Excel Processor Test
echo ========================================
echo.

set PYTHON_PATH=C:\Users\sasha\AppData\Local\Programs\Python\Python311\python.exe

echo [1/2] Checking Python...
%PYTHON_PATH% --version
if %errorlevel% neq 0 (
    echo ❌ ERROR: Python not found!
    pause
    exit /b 1
)

echo.
echo [2/2] Checking openpyxl...
%PYTHON_PATH% -c "import openpyxl; print('✅ openpyxl version:', openpyxl.__version__)"
if %errorlevel% neq 0 (
    echo ❌ ERROR: openpyxl not found!
    pause
    exit /b 1
)

echo.
echo ========================================
echo ✅ SUCCESS! Python is ready
echo ========================================
echo.
echo Python path: %PYTHON_PATH%
echo Python script: python\excel_processor.py
echo.
pause
