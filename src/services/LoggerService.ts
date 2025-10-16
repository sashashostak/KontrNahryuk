/**
 * Сервіс логування
 * FIXED: Винесено функцію log() з main.ts (рядки 88-94)
 * 
 * Відповідальність:
 * - Централізоване логування в UI
 * - Підтримка рівнів логування (info, warning, error)
 * - Автоматичний scroll до останнього повідомлення
 * - Форматування з timestamp
 */

import { byId } from '../utils/helpers';

export class LoggerService {
  private logElement: HTMLElement | null;

  constructor(logElementId: string = 'log-body') {
    this.logElement = document.getElementById(logElementId);
  }

  /**
   * Основний метод логування
   * FIXED: Додано підтримку рівнів логування
   * @param message - Повідомлення для логування
   * @param level - Рівень: 'info' | 'error' | 'warning'
   */
  public log(message: string, level: 'info' | 'error' | 'warning' = 'info'): void {
    const el = this.logElement || byId<HTMLPreElement>('log-body');
    if (!el) return;

    const time = new Date().toLocaleTimeString();
    
    // FIXED: Додано іконки для різних рівнів
    const icon = this.getLevelIcon(level);
    const formattedMessage = `[${time}] ${icon} ${message}\n`;
    
    el.textContent += formattedMessage;
    el.scrollTop = el.scrollHeight;

    // FIXED: Дублюємо в console для дебагу
    this.logToConsole(message, level);
  }

  /**
   * Логування інформаційних повідомлень
   * @param message - Повідомлення
   */
  public info(message: string): void {
    this.log(message, 'info');
  }

  /**
   * Логування помилок
   * @param message - Повідомлення про помилку
   */
  public error(message: string): void {
    this.log(message, 'error');
  }

  /**
   * Логування попереджень
   * @param message - Попереджувальне повідомлення
   */
  public warning(message: string): void {
    this.log(message, 'warning');
  }

  /**
   * Отримання іконки для рівня логування
   * FIXED: Візуальна індикація рівня
   * @private
   */
  private getLevelIcon(level: 'info' | 'error' | 'warning'): string {
    switch (level) {
      case 'info':
        return 'ℹ️';
      case 'error':
        return '❌';
      case 'warning':
        return '⚠️';
      default:
        return 'ℹ️';
    }
  }

  /**
   * Дублювання логів в browser console
   * FIXED: Для зручності дебагу
   * @private
   */
  private logToConsole(message: string, level: 'info' | 'error' | 'warning'): void {
    switch (level) {
      case 'error':
        console.error(message);
        break;
      case 'warning':
        console.warn(message);
        break;
      default:
        console.log(message);
    }
  }

  /**
   * Очищення логів
   * FIXED: Додатковий метод для очищення
   * @public
   */
  public clear(): void {
    const el = this.logElement || byId<HTMLPreElement>('log-body');
    if (el) {
      el.textContent = '';
    }
  }
}

// Експорт singleton instance
export const logger = new LoggerService();
