# UI Cache Issues After Update - Fix Documentation

## 🐛 Проблема

Після оновлення на версію 1.5.7 деякі користувачі скаржилися на "злітлий інтерфейс":
- Білий екран замість UI
- Елементи не відображаються
- Стара версія CSS/JS завантажується
- Кнопки не працюють

## 🔍 Причини

### 1. **Electron Cache**
Electron кешує HTML/CSS/JS файли, і після оновлення завантажує старі версії з кешу замість нових файлів.

### 2. **Неповне оновлення**
При portable оновленні файли замінюються через .bat скрипт, який може не спрацювати якщо:
- Користувач закрив програму раніше часу
- Антивірус заблокував bat файл
- Недостатньо прав

### 3. **Конфлікт версій assets**
Vite генерує файли з хешами (`index-B5rsGJAO.js`), але старий HTML може посилатися на інше ім'я.

## ✅ Рішення

### 1. **Автоматичне очищення кешу після оновлення**

**Файл:** `electron/services/updateService.ts`

```typescript
private async clearElectronCache(): Promise<void> {
  const { session } = require('electron')
  
  // Очистити всі типи кешу
  await session.defaultSession.clearCache()
  await session.defaultSession.clearStorageData({
    storages: ['cookies', 'filesystem', 'indexdb', 'shadercache', 
               'websql', 'serviceworkers', 'cachestorage', 'localstorage']
  })
  await session.defaultSession.clearAuthCache()
}
```

**Виклик:**
- Після завантаження portable
- Після застосування patch
- Перед перезавантаженням додатку

### 2. **Перевірка версії при запуску**

**Файл:** `electron/main.ts`

```typescript
async function checkAndClearCacheIfNeeded() {
  const currentVersion = app.getVersion()
  const lastVersion = storage.getSetting('app.lastVersion', null)
  
  // Якщо версія змінилась - очищаємо кеш
  if (lastVersion && lastVersion !== currentVersion) {
    console.log(`🔄 Оновлення: ${lastVersion} → ${currentVersion}`)
    await session.defaultSession.clearCache()
    await session.defaultSession.clearStorageData({...})
  }
  
  storage.setSetting('app.lastVersion', currentVersion)
}
```

**Логіка:**
- При кожному запуску перевіряється версія
- Якщо версія змінилась → очищення кешу
- Зберігається поточна версія для наступної перевірки

### 3. **Версійований hash в URL**

**Файл:** `electron/main.ts`

```typescript
// Додаємо версію як hash для запобігання кешуванню
mainWindow.loadFile(htmlPath, {
  hash: appVersion.replace(/\./g, '-') // 1.5.7 → 1-5-7
})
```

**Результат:**
- URL: `file:///.../index.html#1-5-7`
- Браузер сприймає це як нову сторінку
- Кеш старої версії ігнорується

## 📋 Workflow оновлення (оновлений)

### Portable Update:
```
1. Завантаження portable.zip
2. Розпакування
3. Створення backup
4. Заміна файлів
5. Очищення тимчасових файлів
6. ✨ НОВОЕ: Очищення кешу Electron
7. Перезавантаження додатку
8. ✨ НОВОЕ: Перевірка версії при старті
9. ✨ НОВОЕ: Повторне очищення кешу якщо потрібно
```

### Patch Update:
```
1. Завантаження patch.zip
2. Розпакування
3. Створення backup
4. Застосування patch
5. Очищення тимчасових файлів
6. ✨ НОВОЕ: Очищення кешу Electron
7. Перезавантаження
8. ✨ НОВОЕ: Перевірка версії при старті
```

## 🧪 Тестування

### Сценарій 1: Оновлення portable
```bash
1. Запустити v1.5.6
2. Оновити до v1.5.7 (portable)
3. Перевірити логи: "🧹 Очищення кешу Electron..."
4. Після перезапуску: "🔄 Оновлення: 1.5.6 → 1.5.7"
5. UI повинен завантажитись коректно
```

### Сценарій 2: Оновлення patch
```bash
1. Запустити v1.5.6
2. Оновити до v1.5.7 (patch)
3. Перевірити очищення кешу
4. UI коректний після перезапуску
```

### Сценарій 3: Повторний запуск
```bash
1. Запустити v1.5.7
2. Закрити і запустити знову
3. НЕ повинно бути очищення кешу (версія не змінилась)
4. UI завантажується швидко
```

## 📊 Логи

### При оновленні:
```
[UpdateService] 🧹 Очищення кешу Electron...
[UpdateService]   ✓ Очищено cache
[UpdateService]   ✓ Очищено storage data
[UpdateService]   ✓ Очищено auth cache
[UpdateService] ✅ Кеш Electron повністю очищено
```

### При першому запуску після оновлення:
```
🔄 Виявлено оновлення версії: 1.5.6 → 1.5.7
🧹 Очищення кешу для запобігання проблем з UI...
✅ Кеш успішно очищено
```

### При звичайному запуску:
```
(немає повідомлень про кеш - версія не змінилась)
```

## ✅ Переваги рішення

1. **Автоматичне виправлення**: Користувач не робить нічого вручну
2. **Подвійний захист**: 
   - Очищення під час оновлення
   - Перевірка при запуску (якщо relaunch не спрацював)
3. **Версійований URL**: Додатковий захист від кешу браузера
4. **Логування**: Можна діагностувати проблеми
5. **Безпечно**: Помилки очищення кешу не критичні

## 🔮 Майбутні покращення

- [ ] Діалог підтвердження перезапуску
- [ ] Збереження стану UI перед оновленням
- [ ] Автоматичний rollback при проблемах
- [ ] Перевірка цілісності файлів після оновлення

## 📝 Версія

**Додано в:** 1.5.7 (hotfix)

**Файли змінено:**
- `electron/services/updateService.ts` - метод `clearElectronCache()`
- `electron/main.ts` - функція `checkAndClearCacheIfNeeded()`
- `electron/main.ts` - версійований hash в `loadFile()`

**Тести:** Очікується зворотній зв'язок від користувачів після deploy
