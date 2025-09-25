# Ukrainian Document Processor 🇺🇦

**Професійне рішення для обробки українських військових документів**

Застосунок для пошуку ключових слів у документах Word та генерації структурованих звітів Excel з підтримкою пакетної обробки.

## ✨ Основні функції

- 🔍 **Розумний пошук** - Пошук ключових слів у документах Word
- 📄 **Контекстна витяжка** - Вилучення релевантних абзаців з контекстом  
- 📊 **Excel звіти** - Генерація структурованих звітів Excel
- 📚 **Пакетна обробка** - Обробка множини документів одночасно
- ⚙️ **Налаштування** - Гнучкі параметри пошуку та фільтрування
- 🌐 **Багатоформатність** - Підтримка різних форматів документів

## 🛠️ Технічні вимоги

- **OS**: Windows 10/11 (x64)
- **Node.js**: 18+ (для розробки)
- **Пам'ять**: 4GB RAM (рекомендовано)
- **Диск**: 100MB вільного місця

## 🚀 Швидкий старт

### Готова версія
1. Завантажте `Ukrainian-Document-Processor-Setup-FINAL.exe`
2. Запустіть інсталятор як адміністратор
3. Слідуйте інструкціям інсталятора
4. Запустіть програму через ярлик

### Розробка
```bash
# Клонувати репозиторій
git clone <repository-url>
cd ukrainian-document-processor

# Встановити залежності
npm install

# Запустити в режимі розробки
npm run dev

# Збілдити проект
npm run build

# Створити Windows інсталятор
npm run dist:win
```

## 📋 Використання

1. **Запустіть програму** - Подвійний клік на ярлик
2. **Оберіть документи** - Виберіть Word файли для обробки
3. **Введіть ключові слова** - Вкажіть терміни для пошуку
4. **Налаштуйте вихід** - Оберіть формат звіту (Excel/Word)
5. **Запустіть обробку** - Натисніть "Обробити"
6. **Збережіть результат** - Оберіть місце збереження звіту

## 🏗️ Архітектура проекту

```
📁 ukrainian-document-processor/
├── 📁 src/                     # Інтерфейс користувача
│   ├── 🎨 index.html          # Основний HTML
│   ├── ⚡ main.ts             # TypeScript логіка UI
│   └── 📁 styles/             # CSS стилі
├── 📁 electron/               # Electron головний процес
│   ├── 🔧 main.ts             # Головний процес Electron
│   ├── 🔗 preload.ts          # Preload скрипти
│   └── 📁 services/           # Основні сервіси
│       ├── 📁 batch/          # Пакетна обробка
│       │   ├── BatchProcessor.ts   # Логіка пакетної обробки
│       │   ├── ExcelReader.ts     # Читання Excel файлів
│       │   ├── ExcelWriter.ts     # Генерація Excel звітів
│       │   └── SheetDateParser.ts # Парсинг дат
│       ├── autoUpdate.ts      # Автоматичні оновлення
│       ├── osIntegration.ts   # Інтеграція з ОС
│       ├── storage.ts         # Локальне зберігання
│       └── updateService.ts   # Сервіс оновлень
├── 📁 exe-launcher/           # Standalone launcher
├── 📁 build/                  # Ресурси для збірки
├── 📁 dist/                   # Скомпільовані файли
└── 📁 test-batch-data/        # Тестові дані
```

## 🔧 Технологічний стек

- **Frontend**: HTML5, CSS3, TypeScript
- **Desktop**: Electron 25+, Node.js 18+
- **Документи**: mammoth.js (Word), docx
- **Excel**: exceljs 
- **Збірка**: Vite, electron-builder
- **Типізація**: TypeScript 5.0+

## 📦 Пакети

### Основний проект
- `electron` - Desktop framework
- `mammoth` - Word document processing  
- `docx` - Word document generation
- `exceljs` - Excel files manipulation
- `vite` - Build tool

### Launcher (exe-launcher)
- `pkg` - Node.js executable compiler
- Standalone launcher для незалежного запуску

## 🤝 Розробка

```bash
# Development сервер
npm run dev

# Збірка renderer процесу  
npm run build:renderer

# Збірка main процесу
npm run build:main

# Повна збірка
npm run build

# Windows інсталятор
npm run dist:win

# Portable версія
npm run dist:portable
```

## 📄 Ліцензія

Copyright © 2024-2025 Ukrainian Document Processor  
Розроблено для Збройних Сил України 🇺🇦

## 🛡️ Підтримка

Для технічної підтримки та звітів про помилки створіть issue в репозиторії.

---

**Слава Україні! Героям слава!** 🇺🇦