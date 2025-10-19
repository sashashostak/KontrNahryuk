/**
 * ThemeService - Управління темою інтерфейсу
 * FIXED: Винесено з main.ts (рядки 1225-1265)
 * 
 * Відповідальність:
 * - Застосування теми (light/dark/system)
 * - Слухач системної теми
 * - Завантаження та збереження налаштування теми
 * 
 * @class ThemeService
 */

import { log } from '../helpers';

type Theme = 'light' | 'dark' | 'system';

export class ThemeService {
  /**
   * Конструктор - ініціалізує ThemeService
   * FIXED: НЕ завантажує тему автоматично, це робить SettingsManager
   */
  constructor() {
    this.setupSystemThemeListener();
    // Застосовуємо дефолтну тему, поки не завантажаться налаштування
    this.applyTheme('system');
  }

  /**
   * Застосування теми до документа
   * FIXED: Додає CSS класи та обробляє system тему
   * @public
   */
  public applyTheme(theme: string): void {
    const root = document.documentElement;
    
    console.log(`🎨 Застосування теми: ${theme}`);
    
    // Видаляємо всі класи тем
    root.classList.remove('light', 'dark');
    
    if (theme === 'light') {
      root.classList.add('light');
      console.log('✅ Додано клас .light');
    } else if (theme === 'dark') {
      // Темна тема за замовчуванням в :root, тому просто видаляємо .light
      console.log('✅ Темна тема активна (без класу)');
    } else if (theme === 'system') {
      // Визначаємо системну тему
      const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      console.log(`🖥️ Системна тема: ${prefersDark ? 'dark' : 'light'}`);
      
      if (prefersDark) {
        // Темна тема - видаляємо .light
        console.log('✅ Системна темна активна');
      } else {
        // Світла тема - додаємо .light
        root.classList.add('light');
        console.log('✅ Системна світла активна');
      }
    }
    
    console.log(`📝 Підсумок: root classes = "${root.className}"`);
  }

  /**
   * Налаштування слухача системної теми
   * FIXED: Автоматично застосовує тему при зміні системної
   * @private
   */
  private setupSystemThemeListener(): void {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    
    mediaQuery.addEventListener('change', async (e) => {
      const currentTheme = await window.api?.getSetting?.('theme', 'system');
      
      // Тільки якщо використовується system тема
      if (currentTheme === 'system') {
        log(`🌓 Системна тема змінена на ${e.matches ? 'темну' : 'світлу'}`);
        this.applyTheme('system');
      }
    });
  }

  /**
   * Зміна теми та збереження налаштування
   * FIXED: Публічний метод для інших модулів
   * @public
   */
  public async setTheme(theme: Theme): Promise<void> {
    try {
      await window.api?.setSetting?.('theme', theme);
      this.applyTheme(theme);
      log(`🎨 Тема змінена на: ${theme}`);
    } catch (error) {
      console.error('Помилка збереження теми:', error);
    }
  }

  /**
   * Отримання поточної теми
   * FIXED: Публічний метод для інших модулів
   * @public
   */
  public async getCurrentTheme(): Promise<Theme> {
    try {
      const theme = await window.api?.getSetting?.('theme', 'system');
      return theme as Theme;
    } catch (error) {
      console.warn('Помилка завантаження теми:', error);
      return 'system';
    }
  }
}
