# 🐷 КонтрНагрюк# KontrNahryuk 🇺🇦



[![Version](https://img.shields.io/badge/version-1.3.0-blue.svg)](https://github.com/sashashostak/KontrNahryuk)**Професійне рішення для обробки документів**

[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

[![Electron](https://img.shields.io/badge/Electron-38.1.2-47848F.svg)](https://www.electronjs.org/)Застосунок для пошуку ключових слів у документах Word та генерації структурованих звітів Excel з підтримкою пакетної обробки.

[![TypeScript](https://img.shields.io/badge/TypeScript-5.0-3178C6.svg)](https://www.typescriptlang.org/)

## ✨ Основні функції

**КонтрНагрюк** — десктопний додаток для автоматизованої обробки військових документів та Excel файлів.

- 🔍 **Розумний пошук** - Пошук ключових слів у документах Word

## 📋 Основні функції- 📄 **Контекстна витяжка** - Вилучення релевантних абзаців з контекстом  

- 📊 **Excel звіти** - Генерація структурованих звітів Excel

- 📄 **Обробка Word документів** - автоматизоване заповнення військових документів- 📚 **Пакетна обробка** - Обробка множини документів одночасно

- 📊 **Зведення Excel файлів** - об'єднання даних з множини Excel файлів- ⚙️ **Налаштування** - Гнучкі параметри пошуку та фільтрування

- 📦 **Пакетна обробка** - масова обробка файлів з прогрес-баром- 🌐 **Багатоформатність** - Підтримка різних форматів документів

- 🔄 **Автооновлення** - автоматична перевірка та встановлення оновлень

- 🎨 **Темна/світла тема** - підтримка теми системи## 🛠️ Технічні вимоги

- 🔐 **Система ліцензування** - контроль доступу до функціоналу

- **OS**: Windows 10/11 (x64)

## 🚀 Швидкий старт- **Node.js**: 18+ (для розробки)

- **Пам'ять**: 4GB RAM (рекомендовано)

### Для користувачів- **Диск**: 100MB вільного місця



1. Завантажте останню версію з [Releases](https://github.com/sashashostak/KontrNahryuk/releases)## 🚀 Швидкий старт

2. Розпакуйте архів

3. Запустіть `KontrNahryuk.exe`### 📥 Для користувачів (Готова програма)

4. Введіть ліцензійний ключ при першому запуску1. **Завантажити**: Перейдіть на https://github.com/sashashostak/KontrNahryuk/releases/latest

2. **Скачати**: Фінальну версію `KontrNahryuk-v1.1.2-Final.zip` (134.61 МБ)

### Для розробників3. **Розпакувати**: Архів у будь-яку папку

4. **Запустити**: `KontrNahryuk.exe`

```bash5. **Ліцензія**: Введіть ключ `KONTR-NAHRYUK-2024`

# Клонування репозиторію

git clone https://github.com/sashashostak/KontrNahryuk.git> 💡 **Програма повністю портативна** - не потребує інсталяції!  

cd KontrNahryuk> 🔧 **Версія 1.1.2** - з виправленими оновленнями та покращеною ліцензійною системою



# Встановлення залежностей### 👨‍💻 Для розробників

npm install```bash

# Клонувати репозиторій

# Запуск в режимі розробкиgit clone <repository-url>

npm run devcd KontrNahryuk



# Збірка для продакшену# Встановити залежності

npm run buildnpm install



# Створення інсталятора# Запустити в режимі розробки

npm run packagenpm run dev

```

# Збілдити проект

## 🛠️ Технологіїnpm run build



- **Electron 38.1.2** - фреймворк для десктопних додатків# Створити Windows інсталятор

- **TypeScript 5.0** - типізований JavaScriptnpm run dist:win

- **Vite 5.4** - швидкий bundler```

- **Node.js** - серверна частина

## 🆕 Що нового у v1.1.2

## 📁 Структура проекту

### 🔑 Система ліцензування

```- **Обов'язкова активація** - програма заблокована до введення ключа

KontrNahryuk/- **Універсальний ключ**: `KONTR-NAHRYUK-2024` (для всіх користувачів)

├── electron/              # Electron main процес- **Покращений UX** - поле ключа зникає після активації

│   ├── main.ts           # Точка входу Electron

│   ├── preload.ts        # Preload скрипт### 🌐 Виправлена перевірка оновлень

│   └── services/         # Backend сервіси- ✅ Усунено помилку "Перевірте інтернет-з'єднання"

├── src/                  # Frontend код- ✅ Працююча інтеграція з GitHub API

│   ├── services/         # Frontend сервіси- ✅ Автоматичне виявлення нових версій

│   ├── utils/            # Утиліти

│   ├── types/            # TypeScript типи### 🧹 Очищення коду

│   ├── constants/        # Константи- Видалено GitHub налаштування (токени, private repos)

│   └── main.ts           # Точка входу UI- Спрощено архітектуру оновлень

├── build/                # Іконки та ресурси- Покращено .gitignore для великих файлів

└── scripts/              # Допоміжні скрипти

```## 📋 Використання



## 📦 Команди1. **Запустіть програму** - Подвійний клік на ярлик

2. **Оберіть документи** - Виберіть Word файли для обробки

```bash3. **Введіть ключові слова** - Вкажіть терміни для пошуку

# Розробка4. **Налаштуйте вихід** - Оберіть формат звіту (Excel/Word)

npm run dev              # Запуск dev сервера з hot reload5. **Запустіть обробку** - Натисніть "Обробити"

6. **Збережіть результат** - Оберіть місце збереження звіту

# Збірка

npm run build            # Компіляція TypeScript + Vite build## 🏗️ Архітектура проекту

npm run package          # Створення portable версії

npm run dist             # Створення інсталятора```

📁 KontrNahryuk/

# Утиліти├── 📁 src/                     # Інтерфейс користувача

npm run typecheck        # Перевірка TypeScript типів│   ├── 🎨 index.html          # Основний HTML

npm run preview          # Попередній перегляд production build│   ├── ⚡ main.ts             # TypeScript логіка UI

```│   └── 📁 styles/             # CSS стилі

├── 📁 electron/               # Electron головний процес

## 🔧 Налаштування│   ├── 🔧 main.ts             # Головний процес Electron

│   ├── 🔗 preload.ts          # Preload скрипти

Налаштування зберігаються в:│   └── 📁 services/           # Основні сервіси

- **Windows**: `%APPDATA%/KontrNahryuk/settings.json`│       ├── 📁 batch/          # Пакетна обробка

- **Linux**: `~/.config/KontrNahryuk/settings.json`│       │   ├── BatchProcessor.ts   # Логіка пакетної обробки

- **macOS**: `~/Library/Application Support/KontrNahryuk/settings.json`│       │   ├── ExcelReader.ts     # Читання Excel файлів

│       │   ├── ExcelWriter.ts     # Генерація Excel звітів

## 🐛 Відомі проблеми│       │   └── SheetDateParser.ts # Парсинг дат

│       ├── autoUpdate.ts      # Автоматичні оновлення

Див. [Issues](https://github.com/sashashostak/KontrNahryuk/issues)│       ├── osIntegration.ts   # Інтеграція з ОС

│       ├── storage.ts         # Локальне зберігання

## 📄 Ліцензія│       └── updateService.ts   # Сервіс оновлень

├── 📁 exe-launcher/           # Standalone launcher

MIT License - див. [LICENSE](LICENSE) файл├── 📁 build/                  # Ресурси для збірки

├── 📁 dist/                   # Скомпільовані файли

## 👥 Контакти└── 📁 test-batch-data/        # Тестові дані

```

- **GitHub**: [@sashashostak](https://github.com/sashashostak)

- **Issues**: [КонтрНагрюк Issues](https://github.com/sashashostak/KontrNahryuk/issues)## 🔧 Технологічний стек



## 🎯 Roadmap- **Frontend**: HTML5, CSS3, TypeScript

- **Desktop**: Electron 25+, Node.js 18+

- [ ] Підтримка macOS та Linux- **Документи**: mammoth.js (Word), docx

- [ ] Інтеграція з cloud storage- **Excel**: exceljs 

- [ ] Розширений API для плагінів- **Збірка**: Vite, electron-builder

- [ ] Мультимовність інтерфейсу- **Типізація**: TypeScript 5.0+



---## 📦 Пакети



**Версія**: 1.3.0  ### Основний проект

**Останнє оновлення**: Жовтень 2025- `electron` - Desktop framework

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

Copyright © 2024-2025 KontrNahryuk

## 🛡️ Підтримка

Для технічної підтримки та звітів про помилки створіть issue в репозиторії.

---

**Слава Україні! Героям слава!** 🇺🇦
