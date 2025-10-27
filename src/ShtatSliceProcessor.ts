/**
 * ShtatSliceProcessor.ts - Управління UI вкладки "Штат_Slice"
 * 
 * Функціонал: Обробка Excel файлу з розділенням на окремі файли
 * - Вибір вхідного Excel файлу
 * - Вибір папки для збереження результату
 * - Обробка та нарізка файлу
 * 
 * @class ShtatSliceProcessor
 */

import { byId } from './helpers';

export class ShtatSliceProcessor {
  private inputFile: string = '';
  private outputFolder: string = '';
  private isProcessing: boolean = false;

  constructor() {
    this.setupEventListeners();
    this.loadSavedSettings();

    this.logMessage('✅ Модуль Штат_Slice ініціалізовано');

    // Оновити стан кнопок після ініціалізації
    setTimeout(() => this.updateButtonStates(), 100);
  }

  /**
   * Налаштування слухачів подій
   */
  private setupEventListeners(): void {
    // Вибір вхідного файлу
    const selectFileBtn = byId('shtat-slice-select-file');
    const inputFileField = byId<HTMLInputElement>('shtat-slice-input-file');
    
    selectFileBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору файлу (Штат_Slice)');
        const result = await window.api?.selectExcelFile?.();
        console.log('📄 Результат вибору файлу:', result);
        
        if (result) {
          this.inputFile = result;
          inputFileField!.value = result;
          this.logMessage(`📄 Обрано файл: ${result}`);
          await this.saveInputFileSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору файлу: ${error}`, 'error');
        console.error('❌ Помилка вибору файлу:', error);
      }
    });

    // Вибір папки призначення
    const selectFolderBtn = byId('shtat-slice-select-folder');
    const outputFolderField = byId<HTMLInputElement>('shtat-slice-output-folder');
    
    selectFolderBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору папки (Штат_Slice)');
        const result = await window.api?.selectFolder?.();
        console.log('📁 Результат вибору папки:', result);
        
        if (result?.filePath) {
          this.outputFolder = result.filePath;
          outputFolderField!.value = result.filePath;
          this.logMessage(`📁 Обрано папку: ${result.filePath}`);
          await this.saveOutputFolderSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору папки: ${error}`, 'error');
        console.error('❌ Помилка вибору папки:', error);
      }
    });

    // Початок обробки
    const startBtn = byId('shtat-slice-process-btn');
    startBtn?.addEventListener('click', () => {
      this.handleStartProcessing();
    });

    // Зупинка обробки
    const stopBtn = byId('shtat-slice-cancel-btn');
    stopBtn?.addEventListener('click', () => {
      this.handleStopProcessing();
    });

    // Очищення логів
    const clearLogBtn = byId('shtat-slice-clear-log');
    clearLogBtn?.addEventListener('click', () => {
      const logBody = byId<HTMLPreElement>('shtat-slice-log-body');
      if (logBody) {
        logBody.textContent = '';
        this.logMessage('🧹 Логи очищено');
      }
    });

    // Копіювання логів
    const copyLogBtn = byId('shtat-slice-copy-log');
    copyLogBtn?.addEventListener('click', () => {
      const logBody = byId<HTMLPreElement>('shtat-slice-log-body');
      if (logBody && logBody.textContent) {
        navigator.clipboard.writeText(logBody.textContent).then(() => {
          this.logMessage('📋 Логи скопійовано');
        }).catch((err) => {
          this.logMessage(`❌ Помилка копіювання: ${err}`, 'error');
        });
      } else {
        this.logMessage('⚠️ Немає логів для копіювання', 'warn');
      }
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
    if (!this.inputFile) {
      this.logMessage('❌ Оберіть вхідний Excel файл', 'error');
      return;
    }

    if (!this.outputFolder) {
      this.logMessage('❌ Оберіть папку для збереження результату', 'error');
      return;
    }
    
    this.isProcessing = true;
    this.updateButtonStates();
    this.clearLog();

    try {
      this.logMessage('🚀 Початок обробки файлу для Штат_Slice...');
      this.logMessage(`📄 Файл: ${this.inputFile}`);
      this.logMessage(`📁 Папка призначення: ${this.outputFolder}`);
      this.logMessage('');
      
      console.log('🔄 Викликаємо обробку Штат_Slice...');

      const options = {
        inputFile: this.inputFile,
        outputFolder: this.outputFolder
      };

      // Викликаємо Python обробку через IPC
      const result = await window.api?.invoke?.('process:shtat-slice', options);
      
      if (result?.ok) {
        const stats = result.stats || {};
        const subunitsCount = stats.subunitsCount || 0;
        const filesCreated = stats.filesCreated || 0;
        const files = result.files || [];
        
        this.logMessage('✅ Обробка завершена успішно!');
        this.logMessage(`📊 Підрозділів знайдено: ${subunitsCount}`);
        this.logMessage(`📁 Файлів створено: ${filesCreated}`);
        
        // Виводимо список створених файлів
        if (files.length > 0) {
          this.logMessage('\n📄 Створені файли:');
          files.forEach((filePath: string, index: number) => {
            const fileName = filePath.split('\\').pop() || filePath;
            this.logMessage(`   ${index + 1}. ${fileName}`);
          });
        }
        
        // Виводимо всі логи з Python (stdout)
        const rawLogs = typeof result.logs === 'string' ? result.logs : '';
        const fallbackLogs = !rawLogs && typeof result.out === 'string' && result.out.includes('\n') ? result.out : '';
        const combinedLogs = (rawLogs || fallbackLogs).trim();
        if (combinedLogs) {
          this.logMessage('\n📋 Детальний лог:');
          for (const line of combinedLogs.split(/\r?\n/)) {
            if (line.trim() && !line.includes('__RESULT__') && !line.includes('__END__')) {
              this.logMessage(line.trim(), 'info');
            }
          }
        }
        
        // Повідомлення про успіх
        await window.api?.notify?.('Успіх', 
          `Штатка нарізана!\n${subunitsCount} підрозділів → ${filesCreated} файлів`
        );
      } else {
        const errorMsg = result?.error || 'Невідома помилка';
        this.logMessage(`❌ Помилка обробки: ${errorMsg}`, 'error');
        const errorLogs = typeof result?.logs === 'string' ? result.logs.trim() : '';
        if (errorLogs) {
          for (const line of errorLogs.split(/\r?\n/)) {
            if (line.trim()) {
              this.logMessage(line.trim(), 'warn');
            }
          }
        }
        await window.api?.notify?.('Помилка', errorMsg);
      }
      
    } catch (error) {
      this.logMessage(`❌ Помилка обробки: ${error}`, 'error');
      console.error('❌ Помилка обробки:', error);
      await window.api?.notify?.('Помилка', `Помилка: ${error}`);
    } finally {
      this.isProcessing = false;
      this.updateButtonStates();
    }
  }

  /**
   * Обробник зупинки процесу
   */
  private handleStopProcessing(): void {
    this.logMessage('🛑 Скасування обробки...');
    this.isProcessing = false;
    this.updateButtonStates();
  }

  /**
   * Оновлення стану кнопок
   */
  private updateButtonStates(): void {
    const processBtn = byId<HTMLButtonElement>('shtat-slice-process-btn');
    const cancelBtn = byId<HTMLButtonElement>('shtat-slice-cancel-btn');

    if (!processBtn || !cancelBtn) return;

    const canStart = this.inputFile && this.outputFolder && !this.isProcessing;

    processBtn.disabled = !canStart;
    processBtn.style.display = this.isProcessing ? 'none' : 'block';
    cancelBtn.style.display = this.isProcessing ? 'block' : 'none';
  }

  /**
   * Очищення логу
   */
  private clearLog(): void {
    const logBody = byId<HTMLPreElement>('shtat-slice-log-body');
    if (logBody) {
      logBody.textContent = '';
    }
  }

  /**
   * Логування повідомлень
   */
  private logMessage(message: string, type: 'info' | 'warn' | 'error' = 'info'): void {
    const logBody = byId<HTMLPreElement>('shtat-slice-log-body');
    if (!logBody) {
      console.warn('⚠️ Лог-контейнер не знайдено');
      return;
    }

    if (message === '') {
      logBody.append(document.createElement('br'));
      logBody.scrollTop = logBody.scrollHeight;
      return;
    }

    const time = new Date().toLocaleTimeString();
    let prefix = 'ℹ️';
    let cssClass = 'log-info';
    if (type === 'error') {
      prefix = '❌';
      cssClass = 'log-error';
    } else if (type === 'warn') {
      prefix = '⚠️';
      cssClass = 'log-warn';
    }

    const line = document.createElement('span');
    line.className = cssClass;
    line.textContent = `[${time}] ${prefix} ${message}`;

    logBody.append(line, document.createElement('br'));
    logBody.scrollTop = logBody.scrollHeight;
  }

  /**
   * Збереження налаштувань вхідного файлу
   */
  private async saveInputFileSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('shtatSlice.inputFile', this.inputFile);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань файлу:', error);
    }
  }

  /**
   * Збереження налаштувань папки призначення
   */
  private async saveOutputFolderSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('shtatSlice.outputFolder', this.outputFolder);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань папки:', error);
    }
  }

  /**
   * Завантаження збережених налаштувань
   */
  private async loadSavedSettings(): Promise<void> {
    try {
      // Завантаження вхідного файлу
      const savedInputFile = await window.api?.getSetting?.('shtatSlice.inputFile');
      if (savedInputFile) {
        this.inputFile = savedInputFile;
        const inputFileField = byId<HTMLInputElement>('shtat-slice-input-file');
        if (inputFileField) {
          inputFileField.value = savedInputFile;
        }
      }

      // Завантаження папки призначення
      const savedOutputFolder = await window.api?.getSetting?.('shtatSlice.outputFolder');
      if (savedOutputFolder) {
        this.outputFolder = savedOutputFolder;
        const outputFolderField = byId<HTMLInputElement>('shtat-slice-output-folder');
        if (outputFolderField) {
          outputFolderField.value = savedOutputFolder;
        }
      }

      this.updateButtonStates();
      this.logMessage('✅ Налаштування завантажено');
    } catch (error) {
      console.error('❌ Помилка завантаження налаштувань:', error);
    }
  }
}
