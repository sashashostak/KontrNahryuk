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

// === IMPORTS ===

// Types and constants
import type { Route } from './types';
import { byId, log } from './helpers';

// Managers
import { SectionManager } from './SectionManager';
import { SourceSelectionManager } from './SourceSelectionManager';
import { BatchManager } from './BatchManager';
import { UpdateManager } from './UpdateManager';
import { LicenseManager } from './LicenseManager';
import { ExcelProcessor } from './ExcelProcessor';

// Services
import { ThemeService } from './services/ThemeService';
import { SettingsManager } from './services/SettingsManager';
import { NavigationService } from './services/NavigationService';

// Utils
import { initializeFilePickers } from './utils/filePicker';

// === GLOBAL STATE ===

let sectionManager: SectionManager;
let sourceSelectionManager: SourceSelectionManager;
let batchManager: BatchManager;
let updateManager: UpdateManager;
let licenseManager: LicenseManager;
let excelProcessor: ExcelProcessor;
let themeService: ThemeService;
let settingsManager: SettingsManager;
let navigationService: NavigationService;

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

    // 2. Ініціалізуємо менеджери основного функціоналу
    sectionManager = new SectionManager();
    log('✅ SectionManager ініціалізовано');

    sourceSelectionManager = new SourceSelectionManager();
    log('✅ SourceSelectionManager ініціалізовано');

    batchManager = new BatchManager();
    log('✅ BatchManager ініціалізовано');

    // 3. Ініціалізуємо менеджери оновлень та ліцензій
    updateManager = new UpdateManager();
    log('✅ UpdateManager ініціалізовано');

    licenseManager = new LicenseManager();
    log('✅ LicenseManager ініціалізовано');

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
    log('🐷 KontrNahryuk v1.3.0 - Завантаження...');

    // 1. Перевірка ліцензії (блокуючий крок)
    await licenseManager?.checkLicenseOnStartup?.();

    // 2. Завантаження основних налаштувань
    await loadSettings();
    setupSettingsAutoSave();

    // 3. Ініціалізація менеджерів
    await initializeManagers();

    // 4. Налаштування глобальних обробників
    setupGlobalEventListeners();

    // 5. Ініціалізація file pickers
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
    license: licenseManager,
    excel: excelProcessor,
    theme: themeService,
    settings: settingsManager,
    navigation: navigationService
  };

  console.log('💡 Доступ до менеджерів через window.__managers');
});
