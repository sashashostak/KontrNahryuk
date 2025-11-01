/**
 * ZBDCheckProcessor - Управління UI вкладки "ЖБД Перевірка"
 *
 * Функціонал: Перевірка Excel файлу ЖБД на помилки
 * - Вибір Excel файлу для перевірки
 * - Вибір місця збереження результату
 * - Автовідкриття результату після обробки
 * 
 * @class ZBDCheckProcessor
 */

import { byId } from './helpers';

export class ZBDCheckProcessor {
  private wordFiles: string[] = [];
  private inputFile: string = '';
  private configExcelFile: string = '';
  private outputFile: string = '';
  private isProcessing: boolean = false;

  constructor() {
    this.setupEventListeners();
    this.loadSavedSettings();

    this.logMessage('✅ Модуль перевірки ЖБД ініціалізовано');

    // Оновити стан кнопок після ініціалізації
    setTimeout(() => this.updateButtonStates(), 100);
  }

  /**
   * Налаштування слухачів подій
   */
  private setupEventListeners(): void {
    // Вибір Word файлів ЖБД
    const selectWordBtn = byId('zbd-check-select-word');
    const wordFilesField = byId<HTMLInputElement>('zbd-check-word-files');

    selectWordBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору Word файлів');
        const result = await window.api?.invoke?.('dialog:select-file', {
          title: 'Оберіть Word файли ЖБД',
          filters: [
            { name: 'Word Files', extensions: ['docx', 'doc'] },
            { name: 'All Files', extensions: ['*'] }
          ],
          properties: ['openFile', 'multiSelections']
        });

        console.log('📁 Результат вибору файлів:', result);

        if (result && !result.canceled && result.filePaths?.length > 0) {
          this.wordFiles = result.filePaths;
          if (wordFilesField) {
            const fileNames = this.wordFiles.map(f => f.split(/[/\\]/).pop()).join(', ');
            wordFilesField.value = `Вибрано файлів: ${this.wordFiles.length} (${fileNames})`;
          }
          this.saveWordFilesSettings(this.wordFiles);
          this.logMessage(`📝 Вибрано Word файлів: ${this.wordFiles.length}`);
          this.wordFiles.forEach((file, index) => {
            this.logMessage(`  ${index + 1}. ${file}`);
          });
          this.updateButtonStates();
        }
      } catch (error) {
        console.error('❌ Помилка при виборі файлів:', error);
        this.logError(`Помилка при виборі файлів: ${error}`);
      }
    });

    // Вибір вхідного Excel файлу
    const selectInputBtn = byId('zbd-check-select-input');
    const inputFileField = byId<HTMLInputElement>('zbd-check-input-file');

    selectInputBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору Excel файлу');
        const result = await window.api?.invoke?.('dialog:select-file', {
          title: 'Оберіть Excel файл для перевірки',
          filters: [
            { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
            { name: 'All Files', extensions: ['*'] }
          ],
          properties: ['openFile']
        });

        console.log('📁 Результат вибору файлу:', result);

        if (result && !result.canceled && result.filePaths?.length > 0) {
          this.inputFile = result.filePaths[0];
          if (inputFileField) {
            inputFileField.value = this.inputFile;
          }
          this.saveInputFileSettings(this.inputFile);
          this.logMessage(`📄 Вибрано файл: ${this.inputFile}`);
          this.updateButtonStates();
        }
      } catch (error) {
        console.error('❌ Помилка при виборі файлу:', error);
        this.logError(`Помилка при виборі файлу: ${error}`);
      }
    });

    // Вибір конфігураційного Excel
    const selectConfigBtn = byId('zbd-check-select-config');
    const configFileField = byId<HTMLInputElement>('zbd-check-config-excel');

    selectConfigBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору конфігураційного Excel');
        const result = await window.api?.invoke?.('dialog:select-file', {
          title: 'Оберіть конфігураційний Excel файл',
          filters: [
            { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
            { name: 'All Files', extensions: ['*'] }
          ],
          properties: ['openFile']
        });

        console.log('⚙️ Результат вибору конфігурації:', result);

        if (result && !result.canceled && result.filePaths?.length > 0) {
          this.configExcelFile = result.filePaths[0];
          if (configFileField) {
            configFileField.value = this.configExcelFile;
          }
          this.saveConfigExcelSettings(this.configExcelFile);
          this.logMessage(`⚙️ Вибрано конфігураційний Excel: ${this.configExcelFile}`);
          this.updateButtonStates();
        }
      } catch (error) {
        console.error('❌ Помилка при виборі конфігурації:', error);
        this.logError(`Помилка при виборі конфігурації: ${error}`);
      }
    });

    // Вибір місця збереження
    const selectOutputBtn = byId('zbd-check-select-output');
    const outputFileField = byId<HTMLInputElement>('zbd-check-output-file');

    selectOutputBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору місця збереження');
        const result = await window.api?.invoke?.('dialog:save-file', {
          defaultPath: 'ZBD_Check_Result.xlsx',
          filters: [
            { name: 'Excel Files', extensions: ['xlsx'] },
            { name: 'All Files', extensions: ['*'] }
          ]
        });

        console.log('💾 Результат вибору місця збереження:', result);

        if (result) {
          this.outputFile = result;
          if (outputFileField) {
            outputFileField.value = this.outputFile;
          }
          this.saveOutputFileSettings(this.outputFile);
          this.logMessage(`💾 Обрано місце збереження: ${this.outputFile}`);
          this.updateButtonStates();
        }
      } catch (error) {
        console.error('❌ Помилка при виборі місця збереження:', error);
        this.logError(`Помилка при виборі місця збереження: ${error}`);
      }
    });

    // Кнопка обробки
    const processBtn = byId('zbd-check-process-btn');
    processBtn?.addEventListener('click', () => this.processFile());

    // Кнопка скасування
    const cancelBtn = byId('zbd-check-cancel-btn');
    cancelBtn?.addEventListener('click', () => this.cancelProcessing());

    // Кнопки логу
    const clearLogBtn = byId('zbd-check-clear-log');
    clearLogBtn?.addEventListener('click', () => this.clearLog());

    const copyLogBtn = byId('zbd-check-copy-log');
    copyLogBtn?.addEventListener('click', () => this.copyLog());
  }

  /**
   * Обробка файлу
   */
  private async processFile(): Promise<void> {
    if (this.isProcessing) {
      this.logWarning('⚠️ Обробка вже виконується');
      return;
    }

    if (this.wordFiles.length === 0) {
      this.logError('❌ Не вибрано Word файли ЖБД');
      return;
    }

    if (!this.inputFile) {
      this.logError('❌ Не вибрано вхідний Excel файл');
      return;
    }

    if (!this.outputFile) {
      this.logError('❌ Не вибрано місце збереження результату');
      return;
    }

    this.isProcessing = true;
    this.updateButtonStates();
    this.clearLog();

    try {
      this.logMessage('🚀 Початок перевірки файлів...');
      this.logMessage(`📝 Word файлів: ${this.wordFiles.length}`);
      this.wordFiles.forEach((file, index) => {
        this.logMessage(`  ${index + 1}. ${file.split(/[/\\]/).pop()}`);
      });
      this.logMessage(`📄 Excel файл: ${this.inputFile}`);
      if (this.configExcelFile) {
        this.logMessage(`⚙️ Конфігураційний Excel: ${this.configExcelFile}`);
      }
      this.logMessage(`💾 Результат буде збережено: ${this.outputFile}`);

      const autoOpen = byId<HTMLInputElement>('zbd-check-autoopen')?.checked ?? true;

      // Викликаємо IPC для перевірки
      const result = await window.api?.invoke?.('process:zbd-check', {
        wordFilePaths: this.wordFiles,
        inputFilePath: this.inputFile,
        configExcelPath: this.configExcelFile || null,
        outputFilePath: this.outputFile,
        autoOpen: autoOpen
      });

      if (result?.success) {
        this.logSuccess('✅ Перевірка успішно завершена!');
        this.logMessage(`📊 Результат збережено: ${this.outputFile}`);

        // Виводимо логи з Python
        if (result.logs) {
          this.logMessage('\n📝 Логи обробки:');
          this.logMessage(result.logs);
        }

        if (result.stats) {
          this.logMessage(`📈 Статистика:`);
          if (result.stats.errors !== undefined) {
            this.logMessage(`  - Помилок: ${result.stats.errors}`);
          }
          if (result.stats.warnings !== undefined) {
            this.logMessage(`  - Попереджень: ${result.stats.warnings}`);
          }
        }

        if (autoOpen) {
          this.logMessage('📂 Відкриття результату...');
        }
      } else {
        // Показуємо логи навіть при помилці
        if (result?.logs) {
          this.logMessage('\n📝 Логи обробки:');
          this.logMessage(result.logs);
        }
        throw new Error(result?.error || 'Невідома помилка під час перевірки');
      }
    } catch (error) {
      console.error('❌ Помилка при перевірці:', error);
      this.logError(`❌ Помилка: ${error}`);
    } finally {
      this.isProcessing = false;
      this.updateButtonStates();
    }
  }

  /**
   * Скасування обробки
   */
  private cancelProcessing(): void {
    this.logWarning('🛑 Скасування обробки...');
    this.isProcessing = false;
    this.updateButtonStates();
  }

  /**
   * Оновлення стану кнопок
   */
  private updateButtonStates(): void {
    const processBtn = byId<HTMLButtonElement>('zbd-check-process-btn');
    const cancelBtn = byId<HTMLButtonElement>('zbd-check-cancel-btn');

    if (processBtn) {
      processBtn.disabled = this.isProcessing || this.wordFiles.length === 0 || !this.inputFile || !this.outputFile;
      processBtn.style.display = this.isProcessing ? 'none' : 'block';
    }

    if (cancelBtn) {
      cancelBtn.style.display = this.isProcessing ? 'block' : 'none';
    }
  }

  /**
   * Логування повідомлення
   */
  private logMessage(message: string): void {
    const logBody = byId('zbd-check-log-body');
    if (logBody) {
      const timestamp = new Date().toLocaleTimeString('uk-UA');
      logBody.textContent += `[${timestamp}] ${message}\n`;
      logBody.scrollTop = logBody.scrollHeight;
    }
    console.log(message);
  }

  /**
   * Логування помилки
   */
  private logError(message: string): void {
    this.logMessage(`❌ ${message}`);
  }

  /**
   * Логування попередження
   */
  private logWarning(message: string): void {
    this.logMessage(`⚠️ ${message}`);
  }

  /**
   * Логування успіху
   */
  private logSuccess(message: string): void {
    this.logMessage(`✅ ${message}`);
  }

  /**
   * Очищення логу
   */
  private clearLog(): void {
    const logBody = byId('zbd-check-log-body');
    if (logBody) {
      logBody.textContent = '';
    }
  }

  /**
   * Копіювання логу
   */
  private async copyLog(): Promise<void> {
    const logBody = byId('zbd-check-log-body');
    if (logBody && logBody.textContent) {
      try {
        await navigator.clipboard.writeText(logBody.textContent);
        this.logMessage('📋 Логи скопійовано в буфер обміну');
      } catch (error) {
        this.logError(`Помилка копіювання: ${error}`);
      }
    }
  }

  /**
   * Збереження налаштувань Word файлів
   */
  private saveWordFilesSettings(files: string[]): void {
    try {
      localStorage.setItem('zbdcheck_word_files', JSON.stringify(files));
    } catch (error) {
      console.error('Помилка збереження налаштувань:', error);
    }
  }

  /**
   * Збереження налаштувань вхідного файлу
   */
  private saveInputFileSettings(filePath: string): void {
    try {
      localStorage.setItem('zbdcheck_input_file', filePath);
    } catch (error) {
      console.error('Помилка збереження налаштувань:', error);
    }
  }

  /**
   * Збереження налаштувань вихідного файлу
   */
  private saveOutputFileSettings(filePath: string): void {
    try {
      localStorage.setItem('zbdcheck_output_file', filePath);
    } catch (error) {
      console.error('Помилка збереження налаштувань:', error);
    }
  }

  /**
   * Збереження налаштувань конфігураційного Excel
   */
  private saveConfigExcelSettings(filePath: string): void {
    try {
      localStorage.setItem('zbdcheck_config_excel', filePath);
    } catch (error) {
      console.error('Помилка збереження налаштувань:', error);
    }
  }

  /**
   * Завантаження збережених налаштувань
   */
  private loadSavedSettings(): void {
    try {
      const savedWordFiles = localStorage.getItem('zbdcheck_word_files');
      const savedInputFile = localStorage.getItem('zbdcheck_input_file');
      const savedConfigExcel = localStorage.getItem('zbdcheck_config_excel');
      const savedOutputFile = localStorage.getItem('zbdcheck_output_file');

      if (savedWordFiles) {
        try {
          this.wordFiles = JSON.parse(savedWordFiles);
          const wordFilesField = byId<HTMLInputElement>('zbd-check-word-files');
          if (wordFilesField && this.wordFiles.length > 0) {
            const fileNames = this.wordFiles.map(f => f.split(/[/\\]/).pop()).join(', ');
            wordFilesField.value = `Вибрано файлів: ${this.wordFiles.length} (${fileNames})`;
          }
        } catch (e) {
          console.error('Помилка парсингу збережених Word файлів:', e);
        }
      }

      if (savedInputFile) {
        this.inputFile = savedInputFile;
        const inputFileField = byId<HTMLInputElement>('zbd-check-input-file');
        if (inputFileField) {
          inputFileField.value = savedInputFile;
        }
      }

      if (savedConfigExcel) {
        this.configExcelFile = savedConfigExcel;
        const configFileField = byId<HTMLInputElement>('zbd-check-config-excel');
        if (configFileField) {
          configFileField.value = savedConfigExcel;
        }
      }

      if (savedOutputFile) {
        this.outputFile = savedOutputFile;
        const outputFileField = byId<HTMLInputElement>('zbd-check-output-file');
        if (outputFileField) {
          outputFileField.value = savedOutputFile;
        }
      }
    } catch (error) {
      console.error('Помилка завантаження налаштувань:', error);
    }
  }
}
