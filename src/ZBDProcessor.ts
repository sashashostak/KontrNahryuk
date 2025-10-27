/**
 * ZBDProcessor - Управління UI вкладки "ЖБД"
 *
 * Функціонал: Обробка CSV файлів з переносом інформації в Word документ
 * - Вибір CSV файлу для обробки
 * - Вибір Word шаблону (опціонально)
 * - Створення Word документа з даними з CSV
 *
 * @class ZBDProcessor
 */

import { byId } from './helpers';

export class ZBDProcessor {
  private csvFile: string = '';
  private outputFile: string = '';
  private isProcessing: boolean = false;

  constructor() {
    this.setupEventListeners();
    this.loadSavedSettings();

    this.logMessage('✅ Модуль обробки ЖБД ініціалізовано');

    // Оновити стан кнопок після ініціалізації
    setTimeout(() => this.updateButtonStates(), 100);
  }

  /**
   * Налаштування слухачів подій
   */
  private setupEventListeners(): void {
    // Вибір CSV файлу
    const selectCsvBtn = byId('zbd-select-csv');
    const csvFileField = byId<HTMLInputElement>('zbd-csv-file');

    selectCsvBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору CSV файлу');
        const result = await window.api?.selectFile?.({
          title: 'Оберіть CSV файл',
          filters: [
            { name: 'CSV Files', extensions: ['csv'] },
            { name: 'All Files', extensions: ['*'] }
          ]
        });
        console.log('📄 Результат вибору файлу:', result);

        if (result) {
          this.csvFile = result;
          csvFileField!.value = result;
          this.logMessage(`📄 Обрано CSV файл: ${result}`);
          await this.saveCsvFileSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору CSV файлу: ${error}`, 'error');
        console.error('❌ Помилка вибору файлу:', error);
      }
    });

    // Вибір місця збереження результату
    const selectOutputBtn = byId('zbd-select-output');
    const outputField = byId<HTMLInputElement>('zbd-output-file');

    selectOutputBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору місця збереження');
        const result = await window.api?.chooseSavePath?.('ЖБД_результат.docx');
        console.log('💾 Результат вибору місця збереження:', result);

        if (result) {
          this.outputFile = result;
          outputField!.value = result;
          this.logMessage(`💾 Обрано місце збереження: ${result}`);
          await this.saveOutputSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору місця збереження: ${error}`, 'error');
        console.error('❌ Помилка вибору місця збереження:', error);
      }
    });

    // Початок обробки
    const startBtn = byId('zbd-process-btn');
    startBtn?.addEventListener('click', () => {
      this.handleStartProcessing();
    });

    // Зупинка обробки
    const stopBtn = byId('zbd-cancel-btn');
    stopBtn?.addEventListener('click', () => {
      this.handleStopProcessing();
    });

    // Автовідкриття
    const autoOpenCheckbox = byId<HTMLInputElement>('zbd-autoopen');
    autoOpenCheckbox?.addEventListener('change', () => {
      this.saveProcessingSettings();
    });

    // Кнопка копіювання логу
    const copyLogBtn = byId('zbd-copy-log');
    copyLogBtn?.addEventListener('click', () => {
      this.copyLog();
    });

    // Кнопка очищення логу
    const clearLogBtn = byId('zbd-clear-log');
    clearLogBtn?.addEventListener('click', () => {
      this.clearLog();
    });
  }

  /**
   * Обробник початку процесу обробки
   */
  private async handleStartProcessing(): Promise<void> {
    if (this.isProcessing) {
      this.logMessage('⚠️ Обробка вже виконується', 'warn');
      return;
    }

    // Валідація вхідних даних
    if (!this.csvFile) {
      this.logMessage('❌ Оберіть CSV файл для обробки', 'error');
      alert('❌ Оберіть CSV файл для обробки');
      return;
    }

    if (!this.outputFile) {
      this.logMessage('❌ Оберіть місце збереження результату', 'error');
      alert('❌ Оберіть місце збереження результату');
      return;
    }

    this.isProcessing = true;
    this.updateButtonStates();

    try {
      this.logMessage('🚀 Початок обробки CSV файлу...');
      this.logMessage(`📄 CSV файл: ${this.csvFile}`);
      this.logMessage(`💾 Результат: ${this.outputFile}`);
      this.logMessage('');

      // Виклик API для обробки CSV файлу
      this.logMessage('⚙️ Обробка CSV файлу через Python...', 'info');

      const result = await window.api?.invoke?.('process:zbd', {
        csvPath: this.csvFile,
        outputPath: this.outputFile
      });

      if (!result) {
        throw new Error('Не вдалося отримати відповідь від сервера');
      }

      if (!result.ok) {
        throw new Error(result.error || 'Невідома помилка обробки');
      }

      this.logMessage('✅ CSV файл успішно оброблено!', 'success');

      if (result.stats) {
        this.logMessage(`📊 Оброблено рядків: ${result.stats.rowsProcessed || 0}`, 'info');
      }

      if (result.message) {
        this.logMessage(result.message, 'info');
      }

      this.logMessage(`💾 Результат збережено: ${this.outputFile}`, 'success');

      // Автовідкриття файлу якщо налаштування увімкнене
      const autoOpenCheckbox = byId<HTMLInputElement>('zbd-autoopen');
      if (autoOpenCheckbox?.checked) {
        this.logMessage('📂 Відкриваю результат...', 'info');
        try {
          await window.api?.openExternal?.(this.outputFile);
        } catch (openError) {
          console.error('Failed to open file:', openError);
          this.logMessage('⚠️ Не вдалося автоматично відкрити файл', 'warn');
        }
      }

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      this.logMessage(`❌ Помилка обробки: ${errorMsg}`, 'error');
      console.error('Processing error:', error);
      alert(`❌ Помилка обробки: ${errorMsg}`);
    } finally {
      this.isProcessing = false;
      this.updateButtonStates();
    }
  }

  /**
   * Обробник зупинки процесу
   */
  private handleStopProcessing(): void {
    if (!this.isProcessing) {
      return;
    }

    this.logMessage('🛑 Зупинка обробки...', 'warn');
    this.isProcessing = false;
    this.updateButtonStates();
  }

  /**
   * Оновлення стану кнопок
   */
  private updateButtonStates(): void {
    const startBtn = byId<HTMLButtonElement>('zbd-process-btn');
    const stopBtn = byId<HTMLButtonElement>('zbd-cancel-btn');

    if (startBtn) {
      // Кнопка активна тільки якщо обрані CSV файл та місце збереження
      const canStart = !this.isProcessing && this.csvFile && this.outputFile;
      startBtn.disabled = !canStart;

      if (!this.csvFile || !this.outputFile) {
        startBtn.title = 'Оберіть CSV файл та місце збереження';
      } else {
        startBtn.title = 'Розпочати обробку CSV файлу';
      }
    }

    if (stopBtn) {
      stopBtn.style.display = this.isProcessing ? 'inline-block' : 'none';
    }
  }

  /**
   * Логування повідомлень
   */
  private logMessage(message: string, level: 'info' | 'warn' | 'error' | 'success' = 'info'): void {
    const logBody = byId<HTMLPreElement>('zbd-log-body');
    if (!logBody) return;

    const time = new Date().toLocaleTimeString();
    const emoji = {
      info: 'ℹ️',
      warn: '⚠️',
      error: '❌',
      success: '✅'
    }[level];

    const logEntry = `[${time}] ${message}\n`;
    logBody.textContent += logEntry;
    logBody.scrollTop = logBody.scrollHeight;

    console.log(`[${level.toUpperCase()}] ${message}`);
  }

  /**
   * Копіювання логу в буфер обміну
   */
  private copyLog(): void {
    const logBody = byId<HTMLPreElement>('zbd-log-body');
    if (!logBody || !logBody.textContent) {
      this.logMessage('⚠️ Лог порожній', 'warn');
      return;
    }

    navigator.clipboard.writeText(logBody.textContent)
      .then(() => {
        this.logMessage('📋 Лог скопійовано в буфер обміну', 'success');
      })
      .catch(err => {
        this.logMessage(`❌ Помилка копіювання: ${err}`, 'error');
      });
  }

  /**
   * Очищення логу
   */
  private clearLog(): void {
    const logBody = byId<HTMLPreElement>('zbd-log-body');
    if (logBody) {
      logBody.textContent = '';
      this.logMessage('🧹 Лог очищено');
    }
  }

  /**
   * Збереження налаштувань CSV файлу
   */
  private async saveCsvFileSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('zbd.csvFile', this.csvFile);
    } catch (error) {
      console.error('Failed to save CSV file settings:', error);
    }
  }

  /**
   * Збереження налаштувань місця збереження
   */
  private async saveOutputSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('zbd.outputFile', this.outputFile);
    } catch (error) {
      console.error('Failed to save output settings:', error);
    }
  }

  /**
   * Збереження налаштувань обробки
   */
  private async saveProcessingSettings(): Promise<void> {
    const autoOpenCheckbox = byId<HTMLInputElement>('zbd-autoopen');

    try {
      await window.api?.setSetting?.('zbd.autoOpen', autoOpenCheckbox?.checked ?? false);
    } catch (error) {
      console.error('Failed to save processing settings:', error);
    }
  }

  /**
   * Завантаження збережених налаштувань
   */
  private async loadSavedSettings(): Promise<void> {
    try {
      // Завантаження CSV файлу
      const savedCsvFile = await window.api?.getSetting?.('zbd.csvFile', '');
      if (savedCsvFile) {
        this.csvFile = savedCsvFile;
        const csvFileField = byId<HTMLInputElement>('zbd-csv-file');
        if (csvFileField) {
          csvFileField.value = savedCsvFile;
        }
      }

      // Завантаження місця збереження
      const savedOutput = await window.api?.getSetting?.('zbd.outputFile', '');
      if (savedOutput) {
        this.outputFile = savedOutput;
        const outputField = byId<HTMLInputElement>('zbd-output-file');
        if (outputField) {
          outputField.value = savedOutput;
        }
      }

      // Завантаження налаштувань обробки
      const autoOpenCheckbox = byId<HTMLInputElement>('zbd-autoopen');
      if (autoOpenCheckbox) {
        autoOpenCheckbox.checked = await window.api?.getSetting?.('zbd.autoOpen', true);
      }

      // Оновити стан кнопок після завантаження налаштувань
      this.updateButtonStates();

    } catch (error) {
      console.error('Failed to load saved settings:', error);
    }
  }
}
