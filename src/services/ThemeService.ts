/**
 * ThemeService - Управління темою інтерфейсу
 */

type Theme = 'light' | 'dark';

export class ThemeService {
  constructor() {
    this.applyTheme('dark');
  }

  public applyTheme(theme: string): void {
    const root = document.documentElement;
    root.classList.remove('light');
    
    if (theme === 'light') {
      root.classList.add('light');
    }
  }

  public async setTheme(theme: Theme): Promise<void> {
    try {
      this.applyTheme(theme);
      await window.api?.setSetting?.('theme', theme);
    } catch (error) {
      console.error('Помилка збереження теми:', error);
    }
  }

  public async getCurrentTheme(): Promise<Theme> {
    try {
      const theme = await window.api?.getSetting?.('theme', 'dark');
      return theme as Theme;
    } catch (error) {
      return 'dark';
    }
  }
}
