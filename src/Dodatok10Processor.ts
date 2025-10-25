/**
 * Dodatok10Processor - Управління UI вкладки "Додаток 10"
 * 
 * Функціонал: Обробка Excel файлів для Додатку 10
 * - Сканування папки з Excel файлами
 * - Вибір файлу призначення
 * - Обробка даних для Додатку 10
 * 
 * @class Dodatok10Processor
 */

import { byId } from './helpers';

export class Dodatok10Processor {
  private inputFolder: string = '';
  private destinationFile: string = '';
  private stroiovkaFile: string = '';
  private correctionsFile: string = '';
  private isProcessing: boolean = false;

  constructor() {
    this.setupEventListeners();
    this.loadSavedSettings();

    this.logMessage('✅ Модуль обробки Додатку 10 ініціалізовано');

    // Оновити стан кнопок після ініціалізації
    setTimeout(() => this.updateButtonStates(), 100);
  }

  /**
   * Налаштування слухачів подій
   */
  private setupEventListeners(): void {
    // Вибір папки з Excel файлами
    const selectFolderBtn = byId('dodatok10-select-folder');
    const inputFolderField = byId<HTMLInputElement>('dodatok10-input-folder');
    
    selectFolderBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору папки (Додаток 10)');
        const result = await window.api?.selectFolder?.();
        console.log('📁 Результат вибору папки:', result);
        if (result?.filePath) {
          this.inputFolder = result.filePath;
          inputFolderField!.value = result.filePath;
          this.logMessage(`📂 Обрана папка: ${result.filePath}`);
          await this.saveInputFolderSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору папки: ${error}`, 'error');
        console.error('❌ Помилка вибору папки:', error);
      }
    });

    // Вибір файлу призначення
    const selectDestinationBtn = byId('dodatok10-select-destination');
    const destinationFileField = byId<HTMLInputElement>('dodatok10-destination-file');
    
    selectDestinationBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору файлу призначення (Додаток 10)');
        const result = await window.api?.selectExcelFile?.();
        console.log('📄 Результат вибору файлу:', result);
        
        if (result) {
          this.destinationFile = result;
          destinationFileField!.value = result;
          this.logMessage(`💾 Обрано файл призначення: ${result}`);
          await this.saveDestinationSettings();
          this.updateButtonStates();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору файлу призначення: ${error}`, 'error');
        console.error('❌ Помилка вибору файлу:', error);
      }
    });

    // Початок обробки
    const startBtn = byId('dodatok10-process-btn');
    startBtn?.addEventListener('click', () => {
      this.handleStartProcessing();
    });

    // Зупинка обробки
    const stopBtn = byId('dodatok10-cancel-btn');
    stopBtn?.addEventListener('click', () => {
      this.handleStopProcessing();
    });

    // Автовідкриття файлу
    const rememberDestinationCheckbox = byId<HTMLInputElement>('dodatok10-remember-destination');
    rememberDestinationCheckbox?.addEventListener('change', () => {
      this.saveDestinationSettings();
    });

    const autoOpenCheckbox = byId<HTMLInputElement>('dodatok10-autoopen');
    autoOpenCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    // Всі інші чекбокси обробки
    const ignoreFormulaColsCheckbox = byId<HTMLInputElement>('dodatok10-ignore-formula-cols');
    ignoreFormulaColsCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    const fnpCheckCheckbox = byId<HTMLInputElement>('dodatok10-fnp-check');
    fnpCheckCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    const updateStatusCheckbox = byId<HTMLInputElement>('dodatok10-update-status-check');
    updateStatusCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    const duplicatesCheckCheckbox = byId<HTMLInputElement>('dodatok10-duplicates-check');
    duplicatesCheckCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    // Чекбокс Стройовка - показ/приховування поля вибору файлу
    const stroiovkaCheckbox = byId<HTMLInputElement>('dodatok10-stroiovka-check');
    const stroiovkaFileSection = byId('dodatok10-stroiovka-file-section');

    stroiovkaCheckbox?.addEventListener('change', () => {
      if (stroiovkaFileSection) {
        stroiovkaFileSection.style.display = stroiovkaCheckbox.checked ? 'block' : 'none';
      }
      this.saveProcessingSettings();
    });

    // Вибір файлу стройовки
    const selectStroiovkaBtn = byId('dodatok10-select-stroiovka');
    const stroiovkaFileField = byId<HTMLInputElement>('dodatok10-stroiovka-file');

    selectStroiovkaBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору файлу стройовки');
        const result = await window.api?.selectExcelFile?.();
        console.log('📄 Результат вибору файлу:', result);

        if (result) {
          this.stroiovkaFile = result;
          stroiovkaFileField!.value = result;
          this.logMessage(`📊 Обрано файл стройовки: ${result}`);
          await this.saveStroiovkaSettings();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору файлу стройовки: ${error}`, 'error');
        console.error('❌ Помилка вибору файлу:', error);
      }
    });

    // Нові чекбокси для виправлень
    const fixRankCheckbox = byId<HTMLInputElement>('dodatok10-fix-rank');
    const fixPositionCheckbox = byId<HTMLInputElement>('dodatok10-fix-position');
    const correctionsFileSection = byId('dodatok10-corrections-file-section');

    // Функція для перевірки чи треба показувати поле вибору файлу
    const updateCorrectionsFileVisibility = () => {
      if (correctionsFileSection) {
        const shouldShow = fixRankCheckbox?.checked || fixPositionCheckbox?.checked;
        correctionsFileSection.style.display = shouldShow ? 'block' : 'none';
      }
    };

    fixRankCheckbox?.addEventListener('change', () => {
      updateCorrectionsFileVisibility();
      this.saveProcessingSettings();
    });

    fixPositionCheckbox?.addEventListener('change', () => {
      updateCorrectionsFileVisibility();
      this.saveProcessingSettings();
    });

    // Вибір файлу з виправленнями
    const selectCorrectionsBtn = byId('dodatok10-select-corrections');
    const correctionsFileField = byId<HTMLInputElement>('dodatok10-corrections-file');

    selectCorrectionsBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору файлу з виправленнями');
        const result = await window.api?.selectExcelFile?.();
        console.log('📄 Результат вибору файлу:', result);

        if (result) {
          this.correctionsFile = result;
          correctionsFileField!.value = result;
          this.logMessage(`📝 Обрано файл з виправленнями: ${result}`);
          await this.saveCorrectionsSettings();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору файлу з виправленнями: ${error}`, 'error');
        console.error('❌ Помилка вибору файлу:', error);
      }
    });

    // Очищення логів
    const clearLogBtn = byId('dodatok10-clear-log');
    clearLogBtn?.addEventListener('click', () => {
      const logBody = byId<HTMLPreElement>('dodatok10-log-body');
      if (logBody) {
        logBody.textContent = '';
        this.logMessage('🧹 Логи очищено');
      }
    });

    // Копіювання логів
    const copyLogBtn = byId('dodatok10-copy-log');
    copyLogBtn?.addEventListener('click', () => {
      const logBody = byId<HTMLPreElement>('dodatok10-log-body');
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
    if (!this.inputFolder) {
      this.logMessage('❌ Оберіть папку з Excel файлами', 'error');
      return;
    }

    if (!this.destinationFile) {
      this.logMessage('❌ Оберіть файл призначення', 'error');
      return;
    }
    
    this.isProcessing = true;
    this.updateButtonStates();

    try {
      this.logMessage('🚀 Початок обробки файлів для Додатку 10...');
      this.logMessage(`📂 Папка: ${this.inputFolder}`);
      this.logMessage(`💾 Призначення: ${this.destinationFile}`);
      this.logMessage('');
      
      console.log('🔄 Викликаємо обробку Додатку 10...');

      // Отримуємо налаштування з UI
      const autoOpenCheckbox = byId<HTMLInputElement>('dodatok10-autoopen');
      const ignoreFormulaColsCheckbox = byId<HTMLInputElement>('dodatok10-ignore-formula-cols');
      const fnpCheckCheckbox = byId<HTMLInputElement>('dodatok10-fnp-check');
      const duplicatesCheckCheckbox = byId<HTMLInputElement>('dodatok10-duplicates-check');
      const stroiovkaCheckCheckbox = byId<HTMLInputElement>('dodatok10-stroiovka-check');
      const updateStatusCheckbox = byId<HTMLInputElement>('dodatok10-update-status-check');
      const fixRankCheckbox = byId<HTMLInputElement>('dodatok10-fix-rank');
      const fixPositionCheckbox = byId<HTMLInputElement>('dodatok10-fix-position');

      const stroiovkaEnabled = stroiovkaCheckCheckbox?.checked || false;
      if (stroiovkaEnabled && !this.stroiovkaFile) {
        this.logMessage('❌ Оберіть файл стройовки для перевірки', 'error');
        this.isProcessing = false;
        this.updateButtonStates();
        return;
      }

      const fixRank = fixRankCheckbox?.checked || false;
      const fixPosition = fixPositionCheckbox?.checked || false;
      if ((fixRank || fixPosition) && !this.correctionsFile) {
        this.logMessage('❌ Додайте файл з виправленнями для звань/посад', 'error');
        this.isProcessing = false;
        this.updateButtonStates();
        return;
      }

      const options = {
        inputFolder: this.inputFolder,
        destinationFile: this.destinationFile,
        autoOpen: autoOpenCheckbox?.checked || false,
        ignoreFormulaCols: ignoreFormulaColsCheckbox?.checked !== false, // default true
        fnpCheck: fnpCheckCheckbox?.checked || false,
        duplicatesCheck: duplicatesCheckCheckbox?.checked || false,
        stroiovkaCheck: stroiovkaEnabled,
        stroiovkaFile: this.stroiovkaFile || '',
        fixRank,
        fixPosition,
        updateStatus: updateStatusCheckbox?.checked || false,
        correctionsFile: this.correctionsFile || ''
      };

      // Викликаємо Python обробку через IPC
      const result = await window.api?.invoke?.('process:dodatok10', options);
      
      if (result?.ok) {
        const stats = result.stats || {};
        this.logMessage('✅ Обробка завершена успішно!');
        this.logMessage(`📊 Файлів оброблено: ${stats.filesProcessed || 0}`);
        this.logMessage(`📊 Рядків записано: ${stats.rowsWritten || 0}`);
        this.logMessage(`📊 Підрозділів знайдено: ${stats.unitsFound || 0}`);
        const fnpErrors = typeof stats.fnpErrors === 'number' ? stats.fnpErrors : undefined;
        if (typeof fnpErrors === 'number') {
          this.logMessage(`📌 FNP помилок: ${fnpErrors}`, fnpErrors > 0 ? 'warn' : 'info');
        }

        const duplicatesErrors = typeof stats.duplicatesErrors === 'number' ? stats.duplicatesErrors : undefined;
        if (typeof duplicatesErrors === 'number') {
          this.logMessage(`📌 Дублів: ${duplicatesErrors}`, duplicatesErrors > 0 ? 'warn' : 'info');
        }

        const stroiovkaErrors = typeof stats.stroiovkaErrors === 'number' ? stats.stroiovkaErrors : undefined;
        if (typeof stroiovkaErrors === 'number') {
          this.logMessage(`📌 Невідповідностей стройовки: ${stroiovkaErrors}`, stroiovkaErrors > 0 ? 'warn' : 'info');
        }

        // Виводимо всі логи з Python (stdout)
        const rawLogs = typeof result.logs === 'string' ? result.logs : '';
        const fallbackLogs = !rawLogs && typeof result.out === 'string' && result.out.includes('\n') ? result.out : '';
        const combinedLogs = (rawLogs || fallbackLogs).trim();
        if (combinedLogs) {
          for (const line of combinedLogs.split(/\r?\n/)) {
            if (line.trim()) {
              this.logMessage(line.trim(), 'info');
            }
          }
        }

        const destinationPath = typeof result.destination === 'string' && result.destination.trim()
          ? result.destination.trim()
          : (typeof result.out === 'string' && result.out.trim() && result.out.trim() !== combinedLogs
            ? result.out.trim()
            : this.destinationFile);
        if (destinationPath) {
          this.logMessage(`💾 Результат збережено: ${destinationPath}`);
        }
        
        // Повідомлення про успіх
        await window.api?.notify?.('Успіх', 
          `Додаток 10 оброблено!\nФайлів: ${stats.filesProcessed || 0}, Рядків: ${stats.rowsWritten || 0}`
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
    const processBtn = byId<HTMLButtonElement>('dodatok10-process-btn');
    const cancelBtn = byId<HTMLButtonElement>('dodatok10-cancel-btn');

    if (!processBtn || !cancelBtn) return;

    const canStart = this.inputFolder && this.destinationFile && !this.isProcessing;

    processBtn.disabled = !canStart;
    processBtn.style.display = this.isProcessing ? 'none' : 'block';
    cancelBtn.style.display = this.isProcessing ? 'block' : 'none';
  }

  /**
   * Логування повідомлень
   */
  private logMessage(message: string, type: 'info' | 'warn' | 'error' = 'info'): void {
    const logBody = byId<HTMLPreElement>('dodatok10-log-body');
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
   * Збереження налаштувань папки вводу
   */
  private async saveInputFolderSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('dodatok10.inputFolder', this.inputFolder);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань папки:', error);
    }
  }

  /**
   * Збереження налаштувань файлу призначення
   */
  private async saveDestinationSettings(): Promise<void> {
    try {
      const rememberCheckbox = byId<HTMLInputElement>('dodatok10-remember-destination');
      if (rememberCheckbox?.checked) {
        await window.api?.setSetting?.('dodatok10.destinationFile', this.destinationFile);
      } else {
        await window.api?.setSetting?.('dodatok10.destinationFile', '');
      }
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань файлу призначення:', error);
    }
  }

  /**
   * Збереження налаштувань обробки
   */
  private async saveProcessingSettings(): Promise<void> {
    try {
      const autoOpenCheckbox = byId<HTMLInputElement>('dodatok10-autoopen');
      const ignoreFormulaColsCheckbox = byId<HTMLInputElement>('dodatok10-ignore-formula-cols');
      const fnpCheckCheckbox = byId<HTMLInputElement>('dodatok10-fnp-check');
      const duplicatesCheckCheckbox = byId<HTMLInputElement>('dodatok10-duplicates-check');
      const stroiovkaCheckCheckbox = byId<HTMLInputElement>('dodatok10-stroiovka-check');
      const fixRankCheckbox = byId<HTMLInputElement>('dodatok10-fix-rank');
      const fixPositionCheckbox = byId<HTMLInputElement>('dodatok10-fix-position');
      const updateStatusCheckbox = byId<HTMLInputElement>('dodatok10-update-status-check');

      await window.api?.setSetting?.('dodatok10.autoOpen', autoOpenCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.ignoreFormulaCols', ignoreFormulaColsCheckbox?.checked !== false);
      await window.api?.setSetting?.('dodatok10.fnpCheck', fnpCheckCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.duplicatesCheck', duplicatesCheckCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.stroiovkaCheck', stroiovkaCheckCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.fixRank', fixRankCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.fixPosition', fixPositionCheckbox?.checked || false);
      await window.api?.setSetting?.('dodatok10.updateStatus', updateStatusCheckbox?.checked || false);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань обробки:', error);
    }
  }

  /**
   * Збереження налаштувань файлу стройовки
   */
  private async saveStroiovkaSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('dodatok10.stroiovkaFile', this.stroiovkaFile);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань файлу стройовки:', error);
    }
  }

  /**
   * Збереження налаштувань файлу з виправленнями
   */
  private async saveCorrectionsSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('dodatok10.correctionsFile', this.correctionsFile);
    } catch (error) {
      console.error('❌ Помилка збереження налаштувань файлу виправлень:', error);
    }
  }

  /**
   * Завантаження збережених налаштувань
   */
  private async loadSavedSettings(): Promise<void> {
    try {
      // Завантаження папки вводу
      const savedInputFolder = await window.api?.getSetting?.('dodatok10.inputFolder');
      if (savedInputFolder) {
        this.inputFolder = savedInputFolder;
        const inputFolderField = byId<HTMLInputElement>('dodatok10-input-folder');
        if (inputFolderField) {
          inputFolderField.value = savedInputFolder;
        }
      }

      // Завантаження файлу призначення
      const savedDestination = await window.api?.getSetting?.('dodatok10.destinationFile');
      if (savedDestination) {
        this.destinationFile = savedDestination;
        const destinationField = byId<HTMLInputElement>('dodatok10-destination-file');
        if (destinationField) {
          destinationField.value = savedDestination;
        }
      }

      // Завантаження налаштувань обробки
      const autoOpen = await window.api?.getSetting?.('dodatok10.autoOpen');
      const autoOpenCheckbox = byId<HTMLInputElement>('dodatok10-autoopen');
      if (autoOpenCheckbox) {
        autoOpenCheckbox.checked = autoOpen !== false;
      }

      const ignoreFormulaCols = await window.api?.getSetting?.('dodatok10.ignoreFormulaCols');
      const ignoreFormulaColsCheckbox = byId<HTMLInputElement>('dodatok10-ignore-formula-cols');
      if (ignoreFormulaColsCheckbox) {
        ignoreFormulaColsCheckbox.checked = ignoreFormulaCols !== false;
      }

      const fnpCheck = await window.api?.getSetting?.('dodatok10.fnpCheck');
      const fnpCheckCheckbox = byId<HTMLInputElement>('dodatok10-fnp-check');
      if (fnpCheckCheckbox) {
        fnpCheckCheckbox.checked = fnpCheck === true;
      }

      const duplicatesCheck = await window.api?.getSetting?.('dodatok10.duplicatesCheck');
      const duplicatesCheckCheckbox = byId<HTMLInputElement>('dodatok10-duplicates-check');
      if (duplicatesCheckCheckbox) {
        duplicatesCheckCheckbox.checked = duplicatesCheck === true;
      }

      const stroiovkaCheck = await window.api?.getSetting?.('dodatok10.stroiovkaCheck');
      const stroiovkaCheckCheckbox = byId<HTMLInputElement>('dodatok10-stroiovka-check');
      if (stroiovkaCheckCheckbox) {
        stroiovkaCheckCheckbox.checked = stroiovkaCheck === true;
        // Показати/сховати секцію вибору файлу стройовки
        const stroiovkaFileSection = byId('dodatok10-stroiovka-file-section');
        if (stroiovkaFileSection) {
          stroiovkaFileSection.style.display = stroiovkaCheck ? 'block' : 'none';
        }
      }

      // Завантаження файлу стройовки
      const savedStroiovka = await window.api?.getSetting?.('dodatok10.stroiovkaFile');
      if (savedStroiovka) {
        this.stroiovkaFile = savedStroiovka;
        const stroiovkaField = byId<HTMLInputElement>('dodatok10-stroiovka-file');
        if (stroiovkaField) {
          stroiovkaField.value = savedStroiovka;
        }
      }

      // Завантаження налаштувань нових чекбоксів
      const fixRank = await window.api?.getSetting?.('dodatok10.fixRank');
      const fixRankCheckbox = byId<HTMLInputElement>('dodatok10-fix-rank');
      if (fixRankCheckbox) {
        fixRankCheckbox.checked = fixRank === true;
      }

      const fixPosition = await window.api?.getSetting?.('dodatok10.fixPosition');
      const fixPositionCheckbox = byId<HTMLInputElement>('dodatok10-fix-position');
      if (fixPositionCheckbox) {
        fixPositionCheckbox.checked = fixPosition === true;
      }

      const updateStatus = await window.api?.getSetting?.('dodatok10.updateStatus');
      const updateStatusCheckbox = byId<HTMLInputElement>('dodatok10-update-status-check');
      if (updateStatusCheckbox) {
        updateStatusCheckbox.checked = updateStatus === true;
      }

      // Завантаження файлу з виправленнями
      const savedCorrections = await window.api?.getSetting?.('dodatok10.correctionsFile');
      if (savedCorrections) {
        this.correctionsFile = savedCorrections;
        const correctionsField = byId<HTMLInputElement>('dodatok10-corrections-file');
        if (correctionsField) {
          correctionsField.value = savedCorrections;
        }
      }

      // Показати/сховати секцію вибору файлу виправлень
      const correctionsFileSection = byId('dodatok10-corrections-file-section');
      if (correctionsFileSection) {
        const shouldShow = fixRank || fixPosition;
        correctionsFileSection.style.display = shouldShow ? 'block' : 'none';
      }

      this.updateButtonStates();
      this.logMessage('✅ Налаштування завантажено');
    } catch (error) {
      console.error('❌ Помилка завантаження налаштувань:', error);
    }
  }
}
