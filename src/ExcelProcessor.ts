/**
 * ExcelProcessor - Управління UI вкладки "Стройовка"
 * 
 * Функціонал: Копіювання даних Excel на основі підрозділів
 * - Сканування папки з Excel файлами
 * - Вибір файлу призначення
 * - Копіювання колонок C:H за ключем з колонки B
 * - Обробка двох аркушів: "ЗС" та "БЗ"
 * 
 * @class ExcelProcessor
 */

import { byId } from './helpers';
import { SubunitMappingProcessor } from './services/SubunitMappingProcessor';
import type { ProcessingStats } from './types/MappingTypes';

export class ExcelProcessor {
  private inputFolder: string = '';
  private destinationFile: string = '';
  private processor: SubunitMappingProcessor;
  private isProcessing: boolean = false;

  constructor() {
    this.processor = new SubunitMappingProcessor();
    this.setupEventListeners();
    this.loadSavedSettings();
    
    this.logMessage('✅ Модуль обробки Excel ініціалізовано');
    
    // 🧪 ТЕСТОВІ ЗНАЧЕННЯ (видалити після тестування)
    // this.inputFolder = 'D:\\TestFolder';
    // this.destinationFile = 'D:\\TestFolder\\destination.xlsx';
    
    // Оновити стан кнопок після ініціалізації
    setTimeout(() => this.updateButtonStates(), 100);
  }

  /**
   * Налаштування слухачів подій
   */
  private setupEventListeners(): void {
    // Вибір папки з Excel файлами
    const selectFolderBtn = byId('excel-select-folder');
    const inputFolderField = byId<HTMLInputElement>('excel-input-folder');
    
    selectFolderBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору папки');
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
    const selectDestinationBtn = byId('excel-select-destination');
    const destinationFileField = byId<HTMLInputElement>('excel-destination-file');
    
    selectDestinationBtn?.addEventListener('click', async () => {
      try {
        console.log('🖱️ Натиснуто кнопку вибору файлу призначення');
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
    const startBtn = byId('excel-process-btn');
    startBtn?.addEventListener('click', () => {
      this.handleStartProcessing();
    });

    // Зупинка обробки
    const stopBtn = byId('excel-cancel-btn');
    stopBtn?.addEventListener('click', () => {
      this.handleStopProcessing();
    });

    // Автовідкриття файлу
    const rememberDestinationCheckbox = byId<HTMLInputElement>('excel-remember-destination');
    rememberDestinationCheckbox?.addEventListener('change', () => {
      this.saveDestinationSettings();
    });

    const sliceCheckbox = byId<HTMLInputElement>('excel-slice-check');
    const mismatchesCheckbox = byId<HTMLInputElement>('excel-mismatches');
    const sanitizerCheckbox = byId<HTMLInputElement>('excel-sanitizer');
    const enable3BSPCheckbox = byId<HTMLInputElement>('enable3BSP');
    const autoOpenCheckbox = byId<HTMLInputElement>('excel-autoopen');
    const duplicatesCheckbox = byId<HTMLInputElement>('excel-duplicates');
    
    sliceCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    mismatchesCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    sanitizerCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    enable3BSPCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    autoOpenCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    duplicatesCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
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
      this.logMessage('🚀 Початок обробки файлів...');
      this.logMessage(`📂 Папка: ${this.inputFolder}`);
      this.logMessage(`💾 Призначення: ${this.destinationFile}`);
      this.logMessage('');
      
      console.log('🔄 Викликаємо processor.process...');

      const stats = await this.processor.process(
        this.inputFolder,
        this.destinationFile,
        (percent: number, message: string) => {
          this.updateProgress(message, percent);
          this.logMessage(message);
        }
      );

      this.displayStats(stats);
      this.logMessage('');
      this.logMessage('✅ Обробка завершена успішно!', 'success');
      
      // Автовідкриття файлу якщо налаштування увімкнене
      const autoOpenCheckbox = byId<HTMLInputElement>('excel-autoopen');
      if (autoOpenCheckbox?.checked) {
        this.logMessage('📂 Відкриваю результат...', 'info');
        try {
          await window.api?.openExternal?.(this.destinationFile);
        } catch (openError) {
          console.error('Failed to open file:', openError);
          this.logMessage('⚠️ Не вдалося автоматично відкрити файл', 'warn');
        }
      }

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      
      // Перевірка на помилку "Permission denied" або "відкритий в іншій програмі"
      if (errorMsg.includes('Permission denied') || 
          errorMsg.includes('відкритий в іншій програмі') ||
          errorMsg.includes('Errno 13')) {
        this.logMessage('❌ ФАЙЛ ЗАБЛОКОВАНИЙ!', 'error');
        this.logMessage('', 'error');
        this.logMessage('Файл призначення відкритий в Excel або іншій програмі.', 'error');
        this.logMessage('', 'error');
        this.logMessage('🔧 Закрийте файл та спробуйте ще раз.', 'error');
        
        // Показуємо alert для користувача
        alert(
          '❌ Файл заблокований!\n\n' +
          'Файл призначення відкритий в Excel або іншій програмі.\n\n' +
          '🔧 Закрийте файл та спробуйте ще раз.'
        );
      } else {
        this.logMessage(`❌ Помилка обробки: ${errorMsg}`, 'error');
      }
      
      console.error('Processing error:', error);
    } finally {
      this.isProcessing = false;
      this.updateButtonStates();
      this.resetProgress();
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
    // TODO: Implement cancellation logic if needed
    this.isProcessing = false;
    this.updateButtonStates();
    this.resetProgress();
  }

  /**
   * Відображення статистики обробки
   */
  private displayStats(stats: ProcessingStats): void {
    this.logMessage('');
    this.logMessage('═══════════════════════════════════════');
    this.logMessage('📊 СТАТИСТИКА ОБРОБКИ');
    this.logMessage('═══════════════════════════════════════');
    this.logMessage(`📁 Оброблено файлів: ${stats.processedFiles} з ${stats.totalFiles}`);
    this.logMessage(`✅ Скопійовано рядків (ЗС): ${stats.totalCopiedRowsZS}`);
    this.logMessage(`✅ Скопійовано рядків (БЗ): ${stats.totalCopiedRowsBZ}`);
    this.logMessage(`⏱️ Час обробки: ${(stats.processingTime / 1000).toFixed(2)} сек`);
    
    if (stats.allMissingSubunits.length > 0) {
      this.logMessage('');
      this.logMessage('⚠️ Підрозділи не знайдені в файлі призначення:');
      const uniqueMissing = [...new Set(stats.allMissingSubunits)];
      uniqueMissing.slice(0, 10).forEach(subunit => {
        this.logMessage(`   • ${subunit}`);
      });
      if (uniqueMissing.length > 10) {
        this.logMessage(`   ... та ще ${uniqueMissing.length - 10} підрозділів`);
      }
    }
    this.logMessage('═══════════════════════════════════════');
  }

  /**
   * Оновлення прогресу обробки
   */
  private updateProgress(phase: string, progress: number): void {
    const progressBar = byId('excel-progress');
    const progressText = byId('excel-progress-text');
    
    if (progressBar) {
      progressBar.style.width = `${progress}%`;
    }
    
    if (progressText) {
      progressText.textContent = `${phase} - ${Math.round(progress)}%`;
    }
  }

  /**
   * Скидання прогресу
   */
  private resetProgress(): void {
    const progressBar = byId('excel-progress');
    const progressText = byId('excel-progress-text');
    
    if (progressBar) {
      progressBar.style.width = '0%';
    }
    
    if (progressText) {
      progressText.textContent = '';
    }
  }

  /**
   * Оновлення стану кнопок
   */
  private updateButtonStates(): void {
    const startBtn = byId<HTMLButtonElement>('excel-process-btn');
    const stopBtn = byId<HTMLButtonElement>('excel-cancel-btn');
    
    if (startBtn) {
      // Кнопка активна тільки якщо обрані папка та файл призначення
      const canStart = !this.isProcessing && this.inputFolder && this.destinationFile;
      startBtn.disabled = !canStart;
      
      // Додаємо підказку
      if (!this.inputFolder || !this.destinationFile) {
        startBtn.title = 'Оберіть папку з файлами та файл призначення';
      } else {
        startBtn.title = 'Розпочати обробку Excel файлів';
      }
    }
    
    if (stopBtn) {
      stopBtn.disabled = !this.isProcessing;
    }
  }

  /**
   * Логування повідомлень
   */
  private logMessage(message: string, level: 'info' | 'warn' | 'error' | 'success' = 'info'): void {
    // Логування видалено з UI
    console.log(`[${level.toUpperCase()}] ${message}`);
  }

  /**
   * Збереження налаштувань папки введення
   */
  private async saveInputFolderSettings(): Promise<void> {
    try {
      await window.api?.setSetting?.('excel.inputFolder', this.inputFolder);
    } catch (error) {
      console.error('Failed to save input folder settings:', error);
    }
  }

  /**
   * Збереження налаштувань файлу призначення
   */
  private async saveDestinationSettings(): Promise<void> {
    const rememberCheckbox = byId<HTMLInputElement>('excel-remember-destination');
    const shouldRemember = rememberCheckbox?.checked ?? false;

    try {
      await window.api?.setSetting?.('excel.rememberDestination', shouldRemember);
      
      if (shouldRemember) {
        await window.api?.setSetting?.('excel.destinationFile', this.destinationFile);
      }
    } catch (error) {
      console.error('Failed to save destination settings:', error);
    }
  }

  /**
   * Збереження налаштувань обробки
   */
  private async saveProcessingSettings(): Promise<void> {
    const sliceCheckbox = byId<HTMLInputElement>('excel-slice-check');
    const mismatchesCheckbox = byId<HTMLInputElement>('excel-mismatches');
    const sanitizerCheckbox = byId<HTMLInputElement>('excel-sanitizer');
    const enable3BSPCheckbox = byId<HTMLInputElement>('enable3BSP');
    const autoOpenCheckbox = byId<HTMLInputElement>('excel-autoopen');
    const duplicatesCheckbox = byId<HTMLInputElement>('excel-duplicates');

    try {
      await window.api?.setSetting?.('excel.enableSliceCheck', sliceCheckbox?.checked ?? false);
      await window.api?.setSetting?.('excel.showMismatches', mismatchesCheckbox?.checked ?? false);
      await window.api?.setSetting?.('excel.enableSanitizer', sanitizerCheckbox?.checked ?? false);
      await window.api?.setSetting?.('enable3BSP', enable3BSPCheckbox?.checked ?? false);
      await window.api?.setSetting?.('excel.autoOpen', autoOpenCheckbox?.checked ?? false);
      await window.api?.setSetting?.('excel.enableDuplicates', duplicatesCheckbox?.checked ?? false);
    } catch (error) {
      console.error('Failed to save processing settings:', error);
    }
  }

  /**
   * Завантаження збережених налаштувань
   */
  private async loadSavedSettings(): Promise<void> {
    try {
      // Завантаження папки введення
      const savedInputFolder = await window.api?.getSetting?.('excel.inputFolder', '');
      if (savedInputFolder) {
        this.inputFolder = savedInputFolder;
        const inputFolderField = byId<HTMLInputElement>('excel-input-folder');
        if (inputFolderField) {
          inputFolderField.value = savedInputFolder;
        }
      }

      // Завантаження файлу призначення
      const shouldRemember = await window.api?.getSetting?.('excel.rememberDestination', false);
      const rememberCheckbox = byId<HTMLInputElement>('excel-remember-destination');
      if (rememberCheckbox) {
        rememberCheckbox.checked = shouldRemember;
      }

      if (shouldRemember) {
        const savedDestination = await window.api?.getSetting?.('excel.destinationFile', '');
        if (savedDestination) {
          this.destinationFile = savedDestination;
          const destinationField = byId<HTMLInputElement>('excel-destination-file');
          if (destinationField) {
            destinationField.value = savedDestination;
          }
        }
      }

      // Завантаження налаштувань обробки
      const sliceCheckbox = byId<HTMLInputElement>('excel-slice-check');
      const mismatchesCheckbox = byId<HTMLInputElement>('excel-mismatches');
      const sanitizerCheckbox = byId<HTMLInputElement>('excel-sanitizer');
      const enable3BSPCheckbox = byId<HTMLInputElement>('enable3BSP');
      const autoOpenCheckbox = byId<HTMLInputElement>('excel-autoopen');
      const duplicatesCheckbox = byId<HTMLInputElement>('excel-duplicates');

      if (sliceCheckbox) {
        sliceCheckbox.checked = await window.api?.getSetting?.('excel.enableSliceCheck', false);
      }
      if (mismatchesCheckbox) {
        mismatchesCheckbox.checked = await window.api?.getSetting?.('excel.showMismatches', false);
      }
      if (sanitizerCheckbox) {
        sanitizerCheckbox.checked = await window.api?.getSetting?.('excel.enableSanitizer', false);
      }
      if (enable3BSPCheckbox) {
        enable3BSPCheckbox.checked = await window.api?.getSetting?.('enable3BSP', false);
      }
      if (autoOpenCheckbox) {
        autoOpenCheckbox.checked = await window.api?.getSetting?.('excel.autoOpen', true);
      }
      if (duplicatesCheckbox) {
        duplicatesCheckbox.checked = await window.api?.getSetting?.('excel.enableDuplicates', false);
      }

      // Оновити стан кнопок після завантаження налаштувань
      this.updateButtonStates();

    } catch (error) {
      console.error('Failed to load saved settings:', error);
    }
  }
}
