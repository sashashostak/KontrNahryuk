@echo off
title KontrNahryuk
color 0A
cls

echo.
echo     ╔══════════════════════════════════════════════════════════╗
echo     ║                      KontrNahryuk v1.0.0                ║
echo     ║                      Starting...                        ║  
echo     ╚══════════════════════════════════════════════════════════╝
echo.
echo  🚀 Starting KontrNahryuk...
echo     Please wait while the application loads...
echo.

REM Перевіряємо чи є npm
where npm >nul 2>&1
if errorlevel 1 (
    echo  ❌ ERROR: npm not found!
    echo     Please install Node.js from: https://nodejs.org/
    echo.
    pause
    exit /b 1
)

REM Запускаємо через npm start у фоні
echo  📦 Running npm start...
start /min "" npm start

REM Даємо час на запуск
timeout /t 4 /nobreak >nul

REM Перевіряємо чи electron запустився
tasklist /FI "IMAGENAME eq electron.exe" 2>nul | find /I /N "electron.exe">nul
if errorlevel 1 (
    echo  ⚠️  Electron process not detected, but application may still be starting...
) else (
    echo  ✅ KontrNahryuk started successfully!
    echo     The main application window should be visible now.
)

echo.
echo  ℹ️  You can close this launcher window.
echo     The application will continue running independently.
echo.
echo  🔧 To fully close the application, use the application's Exit button
echo     or close all KontrNahryuk windows.
echo.
timeout /t 2 /nobreak >nul