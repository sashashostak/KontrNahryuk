# KontrNahryuk v1.4.2 - Advanced Auto-Update System

**Дата релізу:** 18 жовтня 2025  
**Тип релізу:** Продуктивна версія (Production Release)

---

## 🎯 Призначення релізу

Цей реліз містить **повністю автоматизовану систему оновлень v3.0** з прогрес-баром у реальному часі та автоматичним встановленням. Версія 1.4.2 представляє найбільш просунуту систему auto-update з підтримкою EventEmitter API та безпечного заміщення файлів.

---

## 🔄 Advanced Auto-Update System v3.0 (НОВЕ!)

### Революційні можливості:

✅ **Прогрес-бар у реальному часі** - відсоток, швидкість завантаження, розмір  
✅ **Автоматичне встановлення** - ZIP → розпакування → backup → заміна → перезапуск  
✅ **Статусні повідомлення** - "Розпакування...", "Створення backup...", "Оновлення файлів..."  
✅ **Безпечна заміна файлів** - через .bat скрипт з повторними спробами  
✅ **Автоматичний backup** - зберігає 3 останні версії  
✅ **EventEmitter архітектура** - події в реальному часі з backend до UI  
✅ **GitHub Releases API** - без необхідності токенів авторизації  

### Як це працює (повністю автоматично):

1. **Запустіть програму** - система автоматично перевіряє GitHub  
2. **Отримайте повідомлення** - якщо доступна нова версія  
3. **Натисніть "Завантажити та встановити"** - дивіться прогрес у реальному часі:
   - 📊 Завантаження: "45.2% (120 MB / 265 MB) 2.5 MB/s"
   - 📂 Розпакування файлів (AdmZip)
   - 💾 Створення резервної копії поточної версії
   - 🔄 Заміна файлів (atomic replacement via .bat)
   - 🚀 Автоматичний перезапуск додатку
4. **Готово!** - нова версія запущена, стара версія у папці Backup  

---

## 🛠️ Технічні покращення

### Архітектура v3.0:

**UpdateService (electron/services/updateService.ts):**
- ✅ EventEmitter базовий клас для real-time events
- ✅ downloadWithProgress(): fetch streams з прогресом кожні 100ms
- ✅ extractZip(): розпакування через adm-zip (5000+ файлів)
- ✅ createBackup(): автоматичне збереження попередніх версій
- ✅ replaceFiles(): atomic replacement через detached .bat process
- ✅ cleanupOldBackups(): утримання тільки 3 останніх версій

**IPC Communication (electron/main.ts):**
- ✅ Event forwarding: UpdateService → BrowserWindow
- ✅ 3 події: download-progress, status, error
- ✅ Broadcast до всіх відкритих вікон

**Preload API (electron/preload.ts):**
- ✅ onDownloadProgress(callback): прогрес завантаження
- ✅ onUpdateStatus(callback): статусні повідомлення
- ✅ onUpdateError(callback): обробка помилок
- ✅ Security: contextBridge isolation

**UI Component (src/UpdateManager.ts):**
- ✅ Real-time progress bar: percent, speed, size
- ✅ Status display: розпакування, backup, заміна
- ✅ Error handling з детальними повідомленнями

### Нові залежності:

| Пакет | Версія | Призначення |
|-------|--------|-------------|
| **adm-zip** | 0.5.10 | ZIP extraction (швидке розпакування) |
| **@types/adm-zip** | latest | TypeScript type definitions |

### Розмір кодової бази:

| Компонент | v2.0 | v3.0 | Зміни |
|-----------|------|------|-------|
| **updateService.ts** | 274 рядки | 312 рядків | **+38 рядків** (+14%) |
| **UpdateManager.ts** | 463 рядки | 482 рядки | **+19 рядків** (+4%) |
| **main.ts (handlers)** | 190 рядків | 199 рядків | **+9 рядків** (+5%) |
| **preload.ts (API)** | 85 рядків | 90 рядків | **+5 рядків** (+6%) |

### Додано методи:

**updateService.ts:**
- ✅ `downloadWithProgress()` - fetch streams, real-time tracking
- ✅ `extractZip()` - AdmZip integration
- ✅ `createBackup()` - version backup system
- ✅ `replaceFiles()` - .bat script generation
- ✅ `cleanupTempFiles()` - temporary files removal
- ✅ `cleanupOldBackups()` - backup rotation
- ✅ `copyDirectory()` - recursive folder copy
- ✅ `formatBytes()` - human-readable sizes
- ✅ Команди: `npm run release:portable`, `npm run release:full`
- ✅ Розділ "Автооновлення" в USER_GUIDE.md
- ✅ Інструкції в README.md
- ✅ Оновлений RELEASE_CHECKLIST.md з покроковими інструкціями

---

## 📚 Документація

Оновлено/створено:
- ✅ **USER_GUIDE.md** - детальний розділ про автооновлення
- ✅ **README.md** - короткий опис системи оновлень
- ✅ **RELEASE_CHECKLIST.md** - повний шаблон створення релізів

---

## 🧪 Тестування цього релізу

### Для користувачів версії 1.4.1:

1. **Запустіть v1.4.1**
2. **Відкрийте:** ⚙️ Налаштування → Оновлення
3. **Натисніть:** "Перевірити оновлення зараз"
4. **Очікувана поведінка:**
   - Програма виявить версію v1.4.2
   - З'явиться повідомлення про доступне оновлення
   - Кнопка "Завантажити оновлення" стане активною
5. **Завантажте оновлення** - ZIP збережеться у Downloads
6. **Встановіть:**
   - Закрийте v1.4.1
   - Розпакуйте `KontrNahryuk-v1.4.2-portable.zip`
   - Замініть файли
   - Запустіть оновлену програму
7. **Перевірте версію:** ⚙️ Налаштування → Про програму

---

## 📥 Встановлення

### Нові користувачі:

1. Завантажте `KontrNahryuk-v1.4.2-portable.zip`
2. Розпакуйте у будь-яку папку
3. Запустіть `KontrNahryuk.exe`

### Користувачі v1.4.1:

**Опція 1: Через автооновлення (рекомендовано)**
- Відкрийте програму
- Перевірте оновлення
- Завантажте та встановіть

**Опція 2: Вручну**
- Завантажте ZIP з цієї сторінки
- Розпакуйте та замініть файли

---

## 🔍 Що змінилося під капотом

### electron/services/updateService.ts
```typescript
// БУЛО: 772 рядки з ліцензуванням, маніфестами, RSA
class UpdateService {
  constructor(storage: Storage) { /* складна ініціалізація */ }
  // 10+ методів для ліцензій, верифікації, state management
}

// СТАЛО: 274 рядки, тільки GitHub API
class UpdateService {
  constructor() { /* проста ініціалізація */ }
  getCurrentVersion(): string
  checkForUpdates(): Promise<UpdateInfo>
  downloadUpdate(updateInfo: UpdateInfo): Promise<DownloadResult>
  isNewerVersion(remote: string, current: string): boolean
}
```

### electron/main.ts - IPC Handlers
```typescript
// БУЛО: 14 handlers
updates:download, updates:install, updates:get-state, 
updates:get-progress, updates:set-license, updates:check-existing-license,
updates:get-license-info, updates:check-access, updates:check-github,
updates:download-and-install, updates:cancel, updates:save-log, ...

// СТАЛО: 4 handlers
updates:get-version   → updateService.getCurrentVersion()
updates:check         → updateService.checkForUpdates()
updates:download      → updateService.downloadUpdate(updateInfo)
updates:restart-app   → app.relaunch() + app.exit(0)
```

---

## ⚠️ Важливі примітки

1. **Це тестовий реліз** - основне призначення: валідація системи автооновлення
2. **Функціональність ідентична v1.4.1** - зміни тільки в системі оновлень
3. **Безпека:** всі оновлення завантажуються з офіційного GitHub репозиторію
4. **Зворотна сумісність:** можна повернутись до v1.4.1 у будь-який момент

---

## 🐛 Відомі обмеження

- Автооновлення працює тільки для portable версій
- Потрібне підключення до інтернету для перевірки оновлень
- Встановлення оновлення вимагає ручного закриття програми

---

## 📊 Системні вимоги

- **OS**: Windows 10/11 (x64)
- **RAM**: 4GB (рекомендовано 8GB)
- **Диск**: 500MB вільного місця
- **Інтернет**: для перевірки та завантаження оновлень

---

## 🔗 Корисні посилання

- **GitHub Repository**: https://github.com/sashashostak/KontrNahryuk
- **Issues**: https://github.com/sashashostak/KontrNahryuk/issues
- **Документація**: [USER_GUIDE.md](USER_GUIDE.md)
- **Чеклист релізів**: [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md)

---

## 💬 Зворотній зв'язок

Будь ласка, повідомте про будь-які проблеми з автооновленням:
- Створіть issue на GitHub
- Вкажіть версію (1.4.1 → 1.4.2)
- Опишіть кроки відтворення
- Додайте скріншоти (якщо можливо)

---

## 🎉 Подяки

Дякуємо всім, хто тестує цю систему автооновлення! Ваші відгуки допоможуть зробити процес оновлення ще простішим.

---

**Версія:** 1.4.2  
**Розмір архіву:** 276.92 MB  
**Кількість файлів:** 5331  
**Дата:** 18 жовтня 2025

**Слава Україні! 🇺🇦**
