#!/bin/bash
# Скрипт для створення релізу

echo "🚀 Створення релізу KontrNahryuk..."

# 1. Збираємо проект
echo "📦 Збираємо проект..."
npm run build
npm run package

# 2. Створюємо архів
echo "📁 Створюємо архів..."
VERSION=$(node -p "require('./package.json').version")
ZIP_NAME="KontrNahryuk-Portable-v${VERSION}.zip"

# Видаляємо старий архів якщо є
if [ -f "$ZIP_NAME" ]; then
    rm "$ZIP_NAME"
fi

# Створюємо новий архів (Windows)
powershell "Compress-Archive -Path 'KontrNahryuk-Portable' -DestinationPath '$ZIP_NAME'"

echo "✅ Релів готовий: $ZIP_NAME"
echo "📝 Тепер завантажте його на GitHub Releases:"
echo "   https://github.com/sashashostak/KontrNahryuk/releases/new"

# 3. Показуємо розмір файлу
FILE_SIZE=$(powershell "(Get-Item '$ZIP_NAME').Length / 1MB")
echo "📊 Розмір архіву: ${FILE_SIZE}MB"