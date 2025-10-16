/**
 * BatchManager - Управління пакетною обробкою Excel файлів
 * FIXED: Винесено з main.ts (рядки 951-1206)
 * 
 * Відповідальність:
 * - Вибір вхідної папки та файлу результату
 * - Запуск та скасування пакетної обробки
 * - Відображення прогресу обробки
 * - Логування подій обробки
 * - Збереження логу
 */

import type { BatchProgress, BatchResult, LogLevel } from './types';
import { DOM_IDS, EVENT_TYPES, LOCALE, LOG_ICONS, MIME_TYPES, FILE_EXTENSIONS } from './constants';

// FIXED: Додано helper функцію byId
const byId = <T extends HTMLElement>(id: string): T | null => 
  document.getElementById(id) as T | null;

/**
 * Клас для управління пакетною обробкою Excel файлів
 */
export class BatchManager {
  private isProcessing = false;
  private logContainer?: HTMLElement;
  
  /**
   * Конструктор - ініціалізує елементи та налаштовує обробники подій
   */
  constructor() {
    this.setupElements();
    this.setupEventListeners();
    this.loadSavedSettings();
  }

  /**
   * FIXED: Ініціалізація DOM елементів
   */
  private setupElements(): void {
    this.logContainer = byId(DOM_IDS.BATCH_LOG_BODY) || undefined;
  }

  /**
   * FIXED: Налаштування всіх обробників подій
   */
  private setupEventListeners(): void {
    // Вибір папки
    byId(DOM_IDS.BTN_CHOOSE_BATCH_FOLDER)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.selectInputFolder();
    });

    // Вибір файлу результату
    byId(DOM_IDS.BTN_CHOOSE_BATCH_OUTPUT)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.selectOutputFile();
    });

    // Запуск обробки
    byId(DOM_IDS.BTN_START_BATCH)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.startProcessing();
    });

    // Скасування обробки
    byId(DOM_IDS.BTN_CANCEL_BATCH)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.cancelProcessing();
    });

    // Очищення логу
    byId(DOM_IDS.BTN_CLEAR_BATCH_LOG)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.clearLog();
    });

    // Збереження логу
    byId(DOM_IDS.BTN_SAVE_BATCH_LOG)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.saveLog();
    });

    // FIXED: Підписка на події з backend через IPC
    window.api?.onBatchProgress?.((progress: BatchProgress) => {
      this.updateProgress(progress);
    });

    window.api?.onBatchLog?.((logEntry: { level: string, message: string }) => {
      this.addLogEntry(logEntry.level as LogLevel, logEntry.message);
    });

    window.api?.onBatchComplete?.((result: BatchResult) => {
      this.onProcessingComplete(result);
    });
  }

  /**
   * Викликає діалог вибору вхідної папки
   * @async
   */
  private async selectInputFolder(): Promise<void> {
    try {
      const folderPath = await window.api?.selectBatchDirectory?.();
      if (folderPath) {
        const input = byId<HTMLInputElement>(DOM_IDS.BATCH_INPUT_FOLDER);
        if (input) {
          input.value = folderPath;
          this.updateButtonStates();
        }
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка вибору папки: ${error}`);
    }
  }

  /**
   * Викликає діалог вибору файлу результату
   * @async
   */
  private async selectOutputFile(): Promise<void> {
    try {
      const filePath = await window.api?.selectBatchOutputFile?.();
      if (filePath) {
        const input = byId<HTMLInputElement>(DOM_IDS.BATCH_OUTPUT_FILE);
        if (input) {
          input.value = filePath;
        }
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка вибору файлу: ${error}`);
    }
  }

  /**
   * Запускає пакетну обробку
   * @async
   */
  private async startProcessing(): Promise<void> {
    if (this.isProcessing) return;

    const inputFolder = byId<HTMLInputElement>(DOM_IDS.BATCH_INPUT_FOLDER)?.value;
    if (!inputFolder) {
      this.addLogEntry('error', 'Оберіть вхідну папку');
      return;
    }

    let outputFile = byId<HTMLInputElement>(DOM_IDS.BATCH_OUTPUT_FILE)?.value;
    if (!outputFile) {
      // FIXED: Генеруємо автоматичне ім'я файлу
      const now = new Date();
      const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD
      const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-'); // HH-MM-SS
      outputFile = `${inputFolder}\\Індекс_бійців_${dateStr}_${timeStr}${FILE_EXTENSIONS.XLSX}`;
    }

    const options = {
      inputDirectory: inputFolder,
      outputFilePath: outputFile,
      includeStats: true,
      includeConflicts: true,
      includeOccurrences: false, // За замовчуванням відключено через розмір
      resolveConflicts: true
    };

    try {
      this.isProcessing = true;
      this.updateButtonStates();
      this.showProgress();
      this.clearLog();
      this.addLogEntry('info', 'Початок пакетної обробки...');

      await window.api?.startBatchProcessing?.(options);
    } catch (error) {
      this.addLogEntry('error', `Помилка запуску: ${error}`);
      this.isProcessing = false;
      this.updateButtonStates();
      this.hideProgress();
    }
  }

  /**
   * Скасовує поточну обробку
   * @async
   */
  private async cancelProcessing(): Promise<void> {
    if (!this.isProcessing) return;

    try {
      const cancelled = await window.api?.cancelBatchProcessing?.();
      if (cancelled) {
        this.addLogEntry('warning', 'Обробку скасовано');
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка скасування: ${error}`);
    }
  }

  /**
   * Оновлює відображення прогресу обробки
   * @param progress - Об'єкт з інформацією про прогрес
   */
  private updateProgress(progress: BatchProgress): void {
    const progressFill = byId(DOM_IDS.BATCH_PROGRESS_FILL);
    const progressPercent = byId(DOM_IDS.BATCH_PROGRESS_PERCENT);
    const progressStatus = byId(DOM_IDS.BATCH_PROGRESS_STATUS);
    const progressDetail = byId(DOM_IDS.BATCH_PROGRESS_DETAIL);

    if (progressFill) {
      progressFill.style.width = `${progress.percentage}%`;
    }

    if (progressPercent) {
      progressPercent.textContent = `${progress.percentage}%`;
    }

    if (progressStatus) {
      progressStatus.textContent = progress.message;
    }

    if (progressDetail) {
      const timeElapsedStr = Math.round(progress.timeElapsed / 1000);
      let detailText = `Файлів оброблено: ${progress.filesProcessed}/${progress.totalFiles} (${timeElapsedStr}с)`;
      
      if (progress.estimatedTimeRemaining) {
        const etaStr = Math.round(progress.estimatedTimeRemaining / 1000);
        detailText += `, залишилось ~${etaStr}с`;
      }
      
      progressDetail.textContent = detailText;
    }
  }

  /**
   * Додає запис в лог
   * @param level - Рівень логування ('info' | 'warning' | 'error')
   * @param message - Повідомлення для логування
   */
  private addLogEntry(level: LogLevel, message: string): void {
    if (!this.logContainer) return;

    const timestamp = new Date().toLocaleTimeString(LOCALE);
    const levelIcon = LOG_ICONS[level] || LOG_ICONS.info;

    const logLine = `[${timestamp}] ${levelIcon} ${message}\n`;
    this.logContainer.textContent += logLine;
    this.logContainer.scrollTop = this.logContainer.scrollHeight;
  }

  /**
   * Обробляє завершення обробки
   * @param result - Результат обробки
   */
  private onProcessingComplete(result: BatchResult): void {
    this.isProcessing = false;
    this.updateButtonStates();
    this.hideProgress();

    if (result.success) {
      this.addLogEntry('info', `✅ Обробка завершена успішно!`);
      this.addLogEntry('info', `📊 Знайдено ${result.stats.fightersFound} бійців`);
      this.addLogEntry('info', `📁 Результат: ${result.outputFilePath}`);
      
      if (result.stats.conflicts > 0) {
        this.addLogEntry('warning', `⚠️ Знайдено ${result.stats.conflicts} конфліктів`);
      }
    } else {
      this.addLogEntry('error', '❌ Обробка завершилася з помилками');
      result.errors.forEach((error: string) => {
        this.addLogEntry('error', error);
      });
    }
  }

  /**
   * Показує контейнер прогресу
   */
  private showProgress(): void {
    const progressContainer = byId(DOM_IDS.BATCH_PROGRESS);
    if (progressContainer) {
      progressContainer.hidden = false;
    }
  }

  /**
   * Ховає контейнер прогресу
   */
  private hideProgress(): void {
    const progressContainer = byId(DOM_IDS.BATCH_PROGRESS);
    if (progressContainer) {
      progressContainer.hidden = true;
    }
  }

  /**
   * Оновлює стан кнопок (enabled/disabled)
   */
  private updateButtonStates(): void {
    const inputFolder = byId<HTMLInputElement>(DOM_IDS.BATCH_INPUT_FOLDER)?.value;
    const startBtn = byId<HTMLButtonElement>(DOM_IDS.BTN_START_BATCH);
    const cancelBtn = byId<HTMLButtonElement>(DOM_IDS.BTN_CANCEL_BATCH);

    if (startBtn) {
      startBtn.disabled = this.isProcessing || !inputFolder;
      startBtn.textContent = this.isProcessing ? 'Обробляється...' : 'Обробити';
    }

    if (cancelBtn) {
      cancelBtn.disabled = !this.isProcessing;
    }
  }

  /**
   * Очищує вміст логу
   */
  private clearLog(): void {
    if (this.logContainer) {
      this.logContainer.textContent = '';
    }
  }

  /**
   * Зберігає лог у файл
   */
  private saveLog(): void {
    if (!this.logContainer) return;
    
    const logText = this.logContainer.textContent;
    if (!logText) return;

    // FIXED: Створюємо Blob з текстом логу
    const blob = new Blob([logText], { type: MIME_TYPES.TEXT });
    const url = URL.createObjectURL(blob);
    
    // FIXED: Створюємо тимчасовий елемент для завантаження
    const a = document.createElement('a');
    a.href = url;
    a.download = `batch_log_${new Date().toISOString().split('T')[0]}${FILE_EXTENSIONS.TXT}`;
    a.click();
    
    URL.revokeObjectURL(url);
  }

  /**
   * Завантажує збережені налаштування (заглушка для майбутньої функціональності)
   */
  private loadSavedSettings(): void {
    // FIXED: Тут можна додати завантаження збережених шляхів
    this.updateButtonStates();
  }
}
