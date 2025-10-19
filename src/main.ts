/**
 * main.ts - Головний файл додатку (refactored)
 * FIXED: Інтеграція всіх модулів після розділення monolithic main.ts
 * 
 * Відповідальність:
 * - Імпорт та ініціалізація всіх менеджерів
 * - Налаштування глобальних event listeners
 * - Координація між модулями
 * - Завантаження початкових налаштувань
 * 
 * Структура:
 * - Imports: Всі класи та утиліти
 * - Settings: Завантаження та збереження основних налаштувань
 * - Initialization: Створення екземплярів класів
 * - Event Listeners: Глобальні обробники подій
 * 
 * @module main
 * @version 2.0.0 - Refactored version
 */

// === AUTOFILL ERRORS SUPPRESSION ===
// FIXED: Приховуємо DevTools Autofill помилки, які не критичні для Electron
(function suppressAutofillErrors() {
  const originalError = console.error;
  console.error = function(...args: any[]) {
    const msg = args[0]?.toString() || '';
    // Ігноруємо тільки Autofill помилки
    if (msg.includes('Autofill.enable') || msg.includes('Autofill.setAddresses')) {
      return;
    }
    originalError.apply(console, args);
  };
})();

// === IMPORTS ===

// Types and constants
import type { Route } from './types';
import { byId, log } from './helpers';

// Managers
import { SectionManager } from './SectionManager';
import { SourceSelectionManager } from './SourceSelectionManager';
import { BatchManager } from './BatchManager';
import { UpdateManager } from './UpdateManager';
import { ExcelProcessor } from './ExcelProcessor';

// Services
import { ThemeService } from './services/ThemeService';
import { SettingsManager } from './services/SettingsManager';
import { NavigationService } from './services/NavigationService';
import { UILoggerService } from './services/UILoggerService';

// Utils
import { initializeFilePickers } from './filePicker';

// === GLOBAL STATE ===

let sectionManager: SectionManager;
let sourceSelectionManager: SourceSelectionManager;
let batchManager: BatchManager;
let updateManager: UpdateManager;
let excelProcessor: ExcelProcessor;
let themeService: ThemeService;
let settingsManager: SettingsManager;
let navigationService: NavigationService;
let uiLoggerService: UILoggerService;

// === SETTINGS MANAGEMENT ===

/**
 * Завантаження збережених налаштувань при старті
 * FIXED: Завантажує налаштування для головної сторінки
 */
async function loadSettings(): Promise<void> {
  try {
    // Load checkbox states
    const is2BSP = await window.api?.getSetting?.('is2BSP', true);
    const isOrder = await window.api?.getSetting?.('isOrder', false);
    const autoOpen = await window.api?.getSetting?.('autoOpen', true);
    
    // Apply saved states
    const is2BSPCheckbox = byId<HTMLInputElement>('t-2bsp');
    const isOrderCheckbox = byId<HTMLInputElement>('t-order');
    const autoOpenCheckbox = byId<HTMLInputElement>('t-autopen');
    
    if (is2BSPCheckbox) is2BSPCheckbox.checked = is2BSP;
    if (isOrderCheckbox) isOrderCheckbox.checked = isOrder;
    if (autoOpenCheckbox) autoOpenCheckbox.checked = autoOpen;
    
    log('⚙️ Основні налаштування завантажено');
  } catch (err) {
    console.warn('Failed to load settings:', err);
  }
}

/**
 * Налаштування автозбереження для чекбоксів
 * FIXED: Підписка на зміни чекбоксів основних налаштувань
 */
function setupSettingsAutoSave(): void {
  const checkboxes = ['t-2bsp', 't-order', 't-autopen'];
  const settingsMap: Record<string, string> = {
    't-2bsp': 'is2BSP', 
    't-order': 'isOrder',
    't-autopen': 'autoOpen'
  };
  
  checkboxes.forEach(id => {
    const checkbox = byId<HTMLInputElement>(id);
    if (checkbox) {
      checkbox.addEventListener('change', async () => {
        const setting = settingsMap[id];
        await window.api?.setSetting?.(setting, checkbox.checked);
        log(`💾 Збережено ${setting} = ${checkbox.checked}`);
      });
    }
  });
}

// === NAVIGATION HELPERS ===

/**
 * Старий обробник навігації (для сумісності)
 * FIXED: Використовує navigationService всередині
 * @deprecated Використовуйте navigationService.navigateTo() замість цього
 */
function navigate(hash: string): void {
  const route = (hash.replace('#', '') || '/functions') as Route;
  document.querySelectorAll<HTMLElement>('.route').forEach(s => {
    s.hidden = s.dataset.route !== route;
  });
  document.querySelectorAll('.nav a').forEach(a => {
    const href = a.getAttribute('href') || '';
    a.classList.toggle('active', href === `#${route}`);
  });
}

// === INITIALIZATION ===

/**
 * Ініціалізація всіх менеджерів та сервісів
 * FIXED: Створює екземпляри класів у правильному порядку
 */
async function initializeManagers(): Promise<void> {
  try {
    log('🚀 Ініціалізація менеджерів...');

    // 1. Спочатку ініціалізуємо базові сервіси
    themeService = new ThemeService();
    log('✅ ThemeService ініціалізовано');

    settingsManager = new SettingsManager(themeService);
    log('✅ SettingsManager ініціалізовано');

    navigationService = new NavigationService();
    log('✅ NavigationService ініціалізовано');

    // Ініціалізуємо UILoggerService для відображення логів в UI
    uiLoggerService = new UILoggerService();
    log('✅ UILoggerService ініціалізовано');

    // 2. Ініціалізуємо менеджери основного функціоналу
    sectionManager = new SectionManager();
    log('✅ SectionManager ініціалізовано');

    sourceSelectionManager = new SourceSelectionManager();
    log('✅ SourceSelectionManager ініціалізовано');

    batchManager = new BatchManager();
    log('✅ BatchManager ініціалізовано');

    // 3. Ініціалізуємо менеджер оновлень
    updateManager = new UpdateManager();
    log('✅ UpdateManager ініціалізовано');

    // 4. Ініціалізуємо Excel процесор
    excelProcessor = new ExcelProcessor();
    log('✅ ExcelProcessor ініціалізовано');

    log('🎉 Всі менеджери ініціалізовано успішно!');
  } catch (error) {
    console.error('❌ Помилка ініціалізації менеджерів:', error);
  }
}

/**
 * Налаштування глобальних обробників подій
 * FIXED: Підписка на глобальні події
 */
function setupGlobalEventListeners(): void {
  // Старий обробник навігації (для сумісності з існуючою розміткою)
  window.addEventListener('hashchange', () => navigate(location.hash));
  navigate(location.hash);

  // Підписка на зміни маршруту для налаштувань
  navigationService.onRouteChange((route) => {
    if (route === '/settings') {
      // Затримка для завантаження DOM
      setTimeout(() => {
        settingsManager.loadAllSettings();
        settingsManager.setupAutoSave();
        settingsManager.addSettingsAnimations();
      }, 100);
    }
  });

  // Кнопки topbar
  byId('btn-notify')?.addEventListener('click', async () => {
    await window.api?.notify?.('Test', 'This is a test notification');
  });

  byId('btn-docs')?.addEventListener('click', async () => {
    await window.api?.openExternal?.('https://github.com/sashashostak/KontrNahryuk');
  });

  // Кнопка обробки наказу (Functions)
  byId('btn-process-order')?.addEventListener('click', async () => {
    try {
      log('🚀 Початок обробки наказу...');
      
      // 1. Отримання значень з форми
      const sourceType = document.querySelector<HTMLInputElement>('input[name="source-type"]:checked')?.value || 'single-file';
      const resultPath = byId<HTMLInputElement>('result-path')?.value;
      const is2BSP = byId<HTMLInputElement>('t-2bsp')?.checked || false;
      const isOrder = byId<HTMLInputElement>('t-order')?.checked || false;
      const autoOpen = byId<HTMLInputElement>('t-autopen')?.checked || false;
      const excelPath = byId<HTMLInputElement>('excel-path')?.value;
      
      // 2. Перевірка обов'язкових полів
      if (!resultPath) {
        log('❌ Помилка: Оберіть місце збереження результату');
        await window.api?.notify?.('Помилка', 'Оберіть місце збереження результату');
        return;
      }
      
      // 3. Отримання вибраного Word файлу
      let wordFile: File | undefined;
      
      if (sourceType === 'single-file') {
        const fileInput = byId<HTMLInputElement>('word-file');
        wordFile = fileInput?.files?.[0];
        if (!wordFile) {
          log('❌ Помилка: Оберіть Word файл');
          await window.api?.notify?.('Помилка', 'Оберіть Word файл');
          return;
        }
      } else if (sourceType === 'multiple-files') {
        const fileInput = byId<HTMLInputElement>('word-files');
        wordFile = fileInput?.files?.[0];
        if (!wordFile) {
          log('❌ Помилка: Оберіть хоча б один Word файл');
          await window.api?.notify?.('Помилка', 'Оберіть Word файли');
          return;
        }
      } else if (sourceType === 'folder') {
        log('❌ Режим папки поки не підтримується');
        await window.api?.notify?.('Помилка', 'Режим папки поки не реалізований');
        return;
      }
      
      if (!wordFile) {
        log('❌ Помилка: Не вдалося отримати Word файл');
        return;
      }
      
      log(`📂 Обробка файлу: ${wordFile.name}`);
      
      // 4. Читання файлу як ArrayBuffer через FileReader
      log('📖 Читання Word файлу...');
      const fileBuffer = await new Promise<ArrayBuffer>((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as ArrayBuffer);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(wordFile);
      });
      
      if (!fileBuffer) {
        log('❌ Помилка: Не вдалося прочитати файл');
        await window.api?.notify?.('Помилка', 'Не вдалося прочитати Word файл');
        return;
      }
      
      // 5. Підготовка payload для processOrder
      const payload = {
        wordBuf: fileBuffer,
        outputPath: resultPath,
        excelPath: isOrder && excelPath ? excelPath : undefined,
        flags: {
          saveDBPath: false,
          is2BSP: is2BSP,
          isOrder: isOrder,
          tokens: false,
          autoOpen: autoOpen
        },
        mode: is2BSP ? '2BSP' : (isOrder ? 'order' : 'default')
      };
      
      // Очищуємо логи перед початком обробки
      uiLoggerService.clear();
      
      log('⚙️ Обробка наказу...');
      
      // 6. Виклик API
      const result = await window.api?.processOrder?.(payload);
      
      // 7. Обробка результату
      if (result?.ok) {
        const stats = result.stats as { tokens?: number; paragraphs?: number; matched?: number; totalDocuments?: number } | undefined;
        log(`✅ Наказ успішно оброблено!`);
        
        // Показати статистику
        if (stats?.totalDocuments) {
          log(`📊 Створено документів: ${stats.totalDocuments}`);
          log(`📊 Знайдено збігів: ${stats?.matched || 0}`);
        } else {
          log(`📊 Статистика: параграфів - ${stats?.paragraphs || 0}, знайдено - ${stats?.matched || 0}`);
        }
        
        if (result.out) {
          log(`💾 Результат збережено: ${result.out}`);
        }
        
        // Повідомлення про успіх
        const matchCount = stats?.matched || 0;
        await window.api?.notify?.('Успіх', `Наказ оброблено! Знайдено збігів: ${matchCount}`);
        
        // Автовідкриття вже реалізоване в electron/main.ts
        // Не потрібно викликати openExternal тут
        
      } else {
        const errorMsg = result?.error || 'Невідома помилка';
        log(`❌ Помилка обробки: ${errorMsg}`);
        await window.api?.notify?.('Помилка', errorMsg);
      }
      
    } catch (error) {
      log(`❌ Критична помилка: ${error}`);
      console.error('Помилка обробки наказу:', error);
      await window.api?.notify?.('Помилка', `Помилка: ${error}`);
    }
  });

  // Кнопка додавання нотатки (Notes)
  byId('btn-add-note')?.addEventListener('click', async () => {
    const noteInput = byId<HTMLTextAreaElement>('note-input');
    const notesList = byId('notes-list');
    
    if (!noteInput?.value?.trim()) {
      log('⚠️ Нотатка порожня');
      return;
    }
    
    // Створення елемента нотатки
    const noteItem = document.createElement('li');
    noteItem.textContent = noteInput.value;
    noteItem.style.cssText = 'padding: 8px; margin: 4px 0; background: var(--panel); border-radius: 4px;';
    
    // Додавання кнопки видалення
    const deleteBtn = document.createElement('button');
    deleteBtn.textContent = '✕';
    deleteBtn.className = 'btn ghost small';
    deleteBtn.style.cssText = 'margin-left: 8px; float: right;';
    deleteBtn.onclick = () => noteItem.remove();
    
    noteItem.appendChild(deleteBtn);
    notesList?.appendChild(noteItem);
    
    // Очищення поля вводу
    noteInput.value = '';
    log('📝 Нотатку додано');
  });

  // Кнопка перевірки оновлень
  byId('btn-check-updates')?.addEventListener('click', async () => {
    await updateManager.checkForUpdates();
  });

  log('🔗 Глобальні event listeners налаштовано');
}

/**
 * Ініціалізація file pickers
 * FIXED: Викликає utility для file inputs
 */
function setupFilePickers(): void {
  initializeFilePickers();
  log('📁 File pickers ініціалізовано');
}

// === MAIN INITIALIZATION ===

/**
 * Головна функція ініціалізації додатку
 * FIXED: Викликається при завантаженні DOM
 */
async function initializeApp(): Promise<void> {
  try {
    // Отримуємо версію з Electron
    const version = await (window as any).api?.invoke?.('updates:get-version') || '1.4.1';
    log(`🐷 KontrNahryuk v${version} - Завантаження...`);

    // Відображаємо версію в UI
    const versionEl = byId('current-version');
    if (versionEl) versionEl.textContent = version;

    // 1. Ініціалізація менеджерів
    await initializeManagers();

    // 2. Завантаження основних налаштувань
    await loadSettings();
    setupSettingsAutoSave();

    // 3. Налаштування глобальних обробників
    setupGlobalEventListeners();

    // 4. Ініціалізація file pickers
    setupFilePickers();

    log('✅ KontrNahryuk готовий до роботи!');
  } catch (error) {
    console.error('❌ Критична помилка ініціалізації:', error);
  }
}

// === EVENT LISTENERS ===

// Ініціалізація при завантаженні DOM
document.addEventListener('DOMContentLoaded', () => {
  initializeApp();
});

// Ініціалізація теми при завантаженні window
window.addEventListener('load', async () => {
  // Застосування теми
  const theme = await window.api?.getSetting?.('theme', 'system');
  themeService?.applyTheme(theme);
  
  log('🎨 Тему застосовано при window.load');
});

// === EXPORTS (для можливого використання в інших модулях) ===

// Експортуємо для доступу з dev tools консолі після ініціалізації
window.addEventListener('load', () => {
  (window as any).__managers = {
    section: sectionManager,
    source: sourceSelectionManager,
    batch: batchManager,
    update: updateManager,
    excel: excelProcessor,
    theme: themeService,
    settings: settingsManager,
    navigation: navigationService
  };

  console.log('💡 Доступ до менеджерів через window.__managers');
});
