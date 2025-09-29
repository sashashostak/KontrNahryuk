@echo off
REM Скрипт для створення релізу Windows

echo 🚀 Створення релізу KontrNahryuk...

REM 1. Збираємо проект
echo 📦 Збираємо проект...
call npm run build
call npm run package

REM 2. Отримуємо версію з package.json
for /f "tokens=*" %%i in ('node -p "require('./package.json').version"') do set VERSION=%%i
set ZIP_NAME=KontrNahryuk-Portable-v%VERSION%.zip

REM 3. Видаляємо старий архів якщо є
if exist "%ZIP_NAME%" del "%ZIP_NAME%"

REM 4. Створюємо новий архів
echo 📁 Створюємо архів...
powershell "Compress-Archive -Path 'KontrNahryuk-Portable' -DestinationPath '%ZIP_NAME%'"

echo ✅ Реліз готовий: %ZIP_NAME%
echo 📝 Тепер завантажте його на GitHub Releases:
echo    https://github.com/sashashostak/KontrNahryuk/releases/new

REM 5. Показуємо розмір файлу
for /f %%A in ('powershell "(Get-Item '%ZIP_NAME%').Length / 1MB"') do echo 📊 Розмір архіву: %%AMB

pause