/**
 * SettingsManager - Управління налаштуваннями додатку
 * FIXED: Винесено з main.ts (рядки 1265-1520)
 * 
 * Відповідальність:
 * - Завантаження та збереження налаштувань
 * - Експорт/імпорт налаштувань
 * - Скидання до заводських
 * - Автозбереження при зміні
 * - Завантаження інформації про додаток
 * 
 * @class SettingsManager
 */

import { APP_VERSION } from '../constants';
import { byId, log } from '../helpers';
import { ThemeService } from './ThemeService';

interface AppSettings {
  is2BSP: boolean;
  isOrder: boolean;
  autoOpen: boolean;
  showTokens: boolean;
  maxFiles: number;
  batchNotifications: boolean;
  batchLogs: boolean;
  powerSave: boolean;
  startupCheck: boolean;
  minimizeTray: boolean;
  theme: string;
}

export class SettingsManager {
  private themeService: ThemeService;

  /**
   * Конструктор - ініціалізує SettingsManager
   * FIXED: Приймає ThemeService для роботи з темами
   */
  constructor(themeService: ThemeService) {
    this.themeService = themeService;
    this.bindSettingsButtons();
  }

  /**
   * Прив'язка обробників до кнопок налаштувань
   * FIXED: Налаштовує кнопки експорту/імпорту/скидання
   * @private
   */
  private bindSettingsButtons(): void {
    byId('export-settings')?.addEventListener('click', () => {
      this.exportSettings();
    });

    byId('import-settings')?.addEventListener('click', () => {
      this.importSettings();
    });

    byId('reset-settings')?.addEventListener('click', () => {
      this.resetSettings();
    });
  }

  /**
   * Завантаження всіх налаштувань для settings сторінки
   * FIXED: Завантажує налаштування з API та оновлює UI
   * @public
   */
  public async loadAllSettings(): Promise<void> {
    try {
      // Основні налаштування обробки
      const is2BSP = await window.api?.getSetting?.('is2BSP', true);
      const isOrder = await window.api?.getSetting?.('isOrder', false);
      const autoOpen = await window.api?.getSetting?.('autoOpen', true);
      const showTokens = await window.api?.getSetting?.('showTokens', true);
      
      // Пакетна обробка
      const maxFiles = await window.api?.getSetting?.('maxFiles', 100);
      const batchNotifications = await window.api?.getSetting?.('batchNotifications', true);
      const batchLogs = await window.api?.getSetting?.('batchLogs', true);
      
      // Системні налаштування
      const powerSave = await window.api?.getSetting?.('powerSave', true);
      const startupCheck = await window.api?.getSetting?.('startupCheck', true);
      const minimizeTray = await window.api?.getSetting?.('minimizeTray', false);
      
      // Тема
      const theme = await window.api?.getSetting?.('theme', 'system');
      
      // FIXED: Застосування налаштувань через helper
      const set = (id: string, value: any) => {
        const el = byId<HTMLInputElement>(id);
        if (el) {
          if (el.type === 'checkbox') {
            el.checked = Boolean(value);
          } else if (el.type === 'number') {
            el.value = String(value);
          } else {
            el.value = String(value);
          }
        }
      };
      
      set('settings-2bsp', is2BSP);
      set('settings-order', isOrder);
      set('settings-autoopen', autoOpen);
      set('settings-tokens', showTokens);
      set('settings-max-files', maxFiles);
      set('settings-batch-notifications', batchNotifications);
      set('settings-batch-logs', batchLogs);
      set('settings-power-save', powerSave);
      set('settings-startup-check', startupCheck);
      set('settings-minimize-tray', minimizeTray);
      
      // Налаштування теми
      const themeSelect = byId<HTMLSelectElement>('theme-select');
      if (themeSelect) {
        themeSelect.value = theme;
      }
      
      // Застосовуємо тему до документа через сервіс
      this.themeService.applyTheme(theme);
      
      // Завантаження інформації про додаток
      await this.loadAppInfo();
      
      log('⚙️ Налаштування завантажено');
    } catch (err) {
      console.warn('Failed to load extended settings:', err);
    }
  }

  /**
   * Автоматичне збереження налаштувань при зміні
   * FIXED: Підписка на зміни всіх полів налаштувань
   * @public
   */
  public setupAutoSave(): void {
    const settingsMap: { [key: string]: string } = {
      'settings-2bsp': 'is2BSP',
      'settings-order': 'isOrder',
      'settings-autoopen': 'autoOpen',
      'settings-tokens': 'showTokens',
      'settings-batch-notifications': 'batchNotifications',
      'settings-batch-logs': 'batchLogs',
      'settings-power-save': 'powerSave',
      'settings-startup-check': 'startupCheck',
      'settings-minimize-tray': 'minimizeTray',
      'settings-max-files': 'maxFiles'
    };
    
    Object.keys(settingsMap).forEach(id => {
      const el = byId<HTMLInputElement>(id);
      if (el) {
        el.addEventListener('change', async () => {
          const settingKey = settingsMap[id];
          let value: any = el.value;
          
          if (el.type === 'checkbox') {
            value = el.checked;
          } else if (el.type === 'number') {
            value = parseInt(el.value) || 0;
          }
          
          await window.api?.setSetting?.(settingKey, value);
          console.log(`Setting ${settingKey} = ${value}`);
        });
      }
    });
    
    // FIXED: Додаємо обробник для теми через ThemeService
    const themeSelect = byId<HTMLSelectElement>('theme-select');
    if (themeSelect) {
      themeSelect.addEventListener('change', async () => {
        const selectedTheme = themeSelect.value as 'light' | 'dark' | 'system';
        console.log('🎨 Зміна теми на:', selectedTheme);
        await this.themeService.setTheme(selectedTheme);
        console.log('✅ Тема збережена');
      });
      console.log('✅ Theme selector event listener додано');
    } else {
      console.warn('⚠️ theme-select не знайдено в DOM');
    }
  }

  /**
   * Завантаження інформації про додаток
   * FIXED: Показує версію, шлях до даних, останню перевірку
   * @private
   */
  private async loadAppInfo(): Promise<void> {
    try {
      const version = await window.api?.getVersion?.();
      const lastCheck = await window.api?.getSetting?.('lastUpdateCheck', null);
      
      const versionEl = byId('app-version');
      const dataPathEl = byId('data-path');
      const lastCheckEl = byId('last-update-check');
      
      if (versionEl) versionEl.textContent = version || APP_VERSION;
      if (dataPathEl) dataPathEl.textContent = 'Локальні дані додатку';
      if (lastCheckEl && lastCheck) {
        lastCheckEl.textContent = new Date(lastCheck).toLocaleString();
      }
    } catch (err) {
      console.warn('Failed to load app info:', err);
    }
  }

  /**
   * Експорт налаштувань у JSON файл
   * FIXED: Завантажує файл з поточними налаштуваннями
   * @private
   */
  private async exportSettings(): Promise<void> {
    try {
      const settings: Partial<AppSettings> & { version: string; timestamp: string } = {
        version: APP_VERSION,
        timestamp: new Date().toISOString(),
        is2BSP: await window.api?.getSetting?.('is2BSP', true),
        isOrder: await window.api?.getSetting?.('isOrder', false),
        autoOpen: await window.api?.getSetting?.('autoOpen', true),
        showTokens: await window.api?.getSetting?.('showTokens', true),
        maxFiles: await window.api?.getSetting?.('maxFiles', 100),
        batchNotifications: await window.api?.getSetting?.('batchNotifications', true),
        batchLogs: await window.api?.getSetting?.('batchLogs', true),
        powerSave: await window.api?.getSetting?.('powerSave', true),
        startupCheck: await window.api?.getSetting?.('startupCheck', true),
        minimizeTray: await window.api?.getSetting?.('minimizeTray', false),
        theme: await window.api?.getSetting?.('theme', 'system')
      };
      
      const blob = new Blob([JSON.stringify(settings, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = url;
      a.download = `KontrNahryuk_Settings_${new Date().toISOString().split('T')[0]}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      
      URL.revokeObjectURL(url);
      
      log('✅ Налаштування експортовано');
    } catch (err) {
      log(`❌ Помилка експорту: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  /**
   * Імпорт налаштувань з JSON файлу
   * FIXED: Завантажує налаштування з файлу та застосовує
   * @private
   */
  private importSettings(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    
    input.addEventListener('change', async (e) => {
      const file = (e.target as HTMLInputElement).files?.[0];
      if (!file) return;
      
      try {
        const text = await file.text();
        const data = JSON.parse(text);
        
        if (!data || typeof data !== 'object') {
          throw new Error('Неправильний формат файлу налаштувань');
        }
        
        // FIXED: Застосування налаштувань (пропускаємо version та timestamp)
        for (const [key, value] of Object.entries(data)) {
          if (key !== 'version' && key !== 'timestamp') {
            await window.api?.setSetting?.(key, value);
          }
        }
        
        // Перезавантаження налаштувань на сторінці
        await this.loadAllSettings();
        
        log('✅ Налаштування імпортовано та застосовано');
      } catch (err) {
        log(`❌ Помилка імпорту: ${err instanceof Error ? err.message : String(err)}`);
      }
    });
    
    input.click();
  }

  /**
   * Скидання налаштувань до заводських
   * FIXED: Повертає всі налаштування до значень за замовчуванням
   * @private
   */
  private async resetSettings(): Promise<void> {
    if (!confirm('Ви впевнені, що хочете скинути всі налаштування до заводських? Цю дію неможливо відмінити.')) {
      return;
    }
    
    try {
      const defaultSettings: AppSettings = {
        is2BSP: true,
        isOrder: false,
        autoOpen: true,
        showTokens: true,
        maxFiles: 100,
        batchNotifications: true,
        batchLogs: true,
        powerSave: true,
        startupCheck: true,
        minimizeTray: false,
        theme: 'system'
      };
      
      for (const [key, value] of Object.entries(defaultSettings)) {
        await window.api?.setSetting?.(key, value);
      }
      
      // Перезавантаження налаштувань на сторінці
      await this.loadAllSettings();
      
      log('✅ Налаштування скинуто до заводських');
    } catch (err) {
      log(`❌ Помилка скидання: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  /**
   * Додавання анімацій для секцій налаштувань
   * FIXED: Публічний метод для ініціалізації UI ефектів
   * @public
   */
  public addSettingsAnimations(): void {
    const sections = document.querySelectorAll('.settings-section');
    sections.forEach((section, index) => {
      setTimeout(() => {
        section.classList.add('animated');
      }, index * 100);
    });
    
    // Додаємо ripple ефект до кнопок
    const buttons = document.querySelectorAll('.settings-buttons .btn');
    buttons.forEach(btn => {
      btn.classList.add('ripple-effect');
    });
  }
}
