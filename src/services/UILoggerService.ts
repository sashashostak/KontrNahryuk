/**
 * UILoggerService - Сервіс для відображення логів в UI додатку
 * 
 * Відповідальність:
 * - Перехоплення всіх console.log/error/warn з main process
 * - Відображення логів в реальному часі в UI
 * - Автоматичний scroll до останнього логу
 * - Форматування логів з кольорами та емодзі
 * 
 * @module UILoggerService
 * @version 1.0.0
 */

export class UILoggerService {
  private logContainer: HTMLElement | null = null;
  private maxLogs: number = 100; // Зменшено до 100 для простішого перегляду
  private autoScroll: boolean = true;

  constructor() {
    this.initialize();
    this.setupClearButton();
    this.setupCopyButton();
    this.clearOldLogs(); // Очищення при запуску
  }

  /**
   * Ініціалізація сервісу логування
   */
  private initialize(): void {
    this.logContainer = document.getElementById('log-body');
    
    if (!this.logContainer) {
      console.warn('❌ UILoggerService: log-body не знайдено в DOM');
      return;
    }

    // Підписка на логи з main process
    this.setupLogListeners();
    
    // Перехоплення console в renderer process
    this.interceptConsole();
    
    this.log('info', '✅ UILoggerService ініціалізовано');
  }

  /**
   * Підписка на логи з main process через IPC
   */
  private setupLogListeners(): void {
    // Слухаємо логи з main process
    if (window.api && 'onLog' in window.api) {
      (window.api as any).onLog((level: string, message: string) => {
        this.log(level as 'info' | 'warn' | 'error', message);
      });
    }
  }

  /**
   * Налаштування кнопки очищення логів
   */
  private setupClearButton(): void {
    const clearButton = document.getElementById('btn-clear-log');
    if (clearButton) {
      clearButton.addEventListener('click', () => {
        this.clear();
        this.log('info', '🧹 Логи очищено');
      });
    }
  }

  /**
   * Налаштування кнопки копіювання логів
   */
  private setupCopyButton(): void {
    const copyButton = document.getElementById('btn-copy-log');
    if (copyButton) {
      copyButton.addEventListener('click', async () => {
        const logBody = document.getElementById('log-body');
        if (logBody && logBody.textContent) {
          try {
            await navigator.clipboard.writeText(logBody.textContent);
            this.log('info', '📋 Логи скопійовано в буфер обміну');
          } catch (err) {
            this.log('error', `❌ Помилка копіювання: ${err}`);
          }
        } else {
          this.log('warn', '⚠️ Немає логів для копіювання');
        }
      });
    }
  }

  /**
   * Перехоплення console.log/warn/error в renderer process
   */
  private interceptConsole(): void {
    const originalLog = console.log;
    const originalWarn = console.warn;
    const originalError = console.error;

    console.log = (...args: any[]) => {
      originalLog.apply(console, args);
      this.log('info', this.formatArgs(args));
    };

    console.warn = (...args: any[]) => {
      originalWarn.apply(console, args);
      this.log('warn', this.formatArgs(args));
    };

    console.error = (...args: any[]) => {
      // Ігноруємо Autofill помилки
      const msg = args[0]?.toString() || '';
      if (msg.includes('Autofill.enable') || msg.includes('Autofill.setAddresses')) {
        return;
      }
      originalError.apply(console, args);
      this.log('error', this.formatArgs(args));
    };
  }

  /**
   * Форматування аргументів console в строку
   */
  private formatArgs(args: any[]): string {
    return args.map(arg => {
      if (typeof arg === 'object') {
        try {
          return JSON.stringify(arg, null, 2);
        } catch {
          return String(arg);
        }
      }
      return String(arg);
    }).join(' ');
  }

  /**
   * Додавання логу в UI
   */
  public log(level: 'info' | 'warn' | 'error', message: string): void {
    if (!this.logContainer) return;

    const timestamp = new Date().toLocaleTimeString('uk-UA');
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry log-${level}`;
    
    // Визначаємо емодзі залежно від рівня
    let emoji = '';
    switch (level) {
      case 'error':
        emoji = '❌';
        break;
      case 'warn':
        emoji = '⚠️';
        break;
      default:
        // Зберігаємо емодзі з повідомлення якщо є
        emoji = this.extractEmoji(message) || '📝';
    }

    logEntry.innerHTML = `
      <span class="log-time">${timestamp}</span>
      <span class="log-icon">${emoji}</span>
      <span class="log-message">${this.escapeHtml(message)}</span>
    `;

    // Додаємо до контейнера
    this.logContainer.appendChild(logEntry);

    // Обмежуємо кількість логів
    this.limitLogs();

    // Автоматичний scroll вниз
    if (this.autoScroll) {
      this.scrollToBottom();
    }
  }

  /**
   * Витягування емодзі з початку повідомлення
   */
  private extractEmoji(message: string): string | null {
    const emojiRegex = /^[\u{1F300}-\u{1F9FF}]|^[\u{2600}-\u{26FF}]|^[\u{2700}-\u{27BF}]/u;
    const match = message.match(emojiRegex);
    return match ? match[0] : null;
  }

  /**
   * Escape HTML символів для безпечного відображення
   */
  private escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  /**
   * Обмеження кількості логів
   */
  private limitLogs(): void {
    if (!this.logContainer) return;

    const entries = this.logContainer.querySelectorAll('.log-entry');
    if (entries.length > this.maxLogs) {
      const toRemove = entries.length - this.maxLogs;
      for (let i = 0; i < toRemove; i++) {
        entries[i].remove();
      }
    }
  }

  /**
   * Прокрутка до останнього логу
   */
  private scrollToBottom(): void {
    if (!this.logContainer) return;
    this.logContainer.scrollTop = this.logContainer.scrollHeight;
  }

  /**
   * Очищення всіх логів
   */
  public clear(): void {
    if (this.logContainer) {
      this.logContainer.innerHTML = '';
    }
  }

  /**
   * Очищення старих логів при запуску додатку
   * Автоматично видаляє всі логи для чистого старту
   */
  private clearOldLogs(): void {
    // Очищаємо логи при кожному запуску
    this.clear();
    this.log('info', '🧹 Логи очищено при запуску додатку');
  }

  /**
   * Увімкнення/вимкнення автоматичного scroll
   */
  public setAutoScroll(enabled: boolean): void {
    this.autoScroll = enabled;
  }
}
