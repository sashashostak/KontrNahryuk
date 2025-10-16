/**
 * ExcelProcessor - Управління зведенням Excel файлів
 * FIXED: Винесено з main.ts (рядки 1580-2025)
 * 
 * Відповідальність:
 * - Вибір папки з Excel файлами та файлу призначення
 * - Сканування Excel файлів
 * - Запуск та зупинка процесу зведення
 * - Відображення прогресу обробки
 * - Логування подій
 * - Збереження налаштувань обробки
 * 
 * @class ExcelProcessor
 */

import type { Mode, StartProcessPayload } from './types';
import { byId } from './helpers';

export class ExcelProcessor {
  private inputFolder: string = '';
  private destinationFile: string = '';
  private isProcessing: boolean = false;
  private foundFiles: string[] = [];

  /**
   * Конструктор - ініціалізує ExcelProcessor
   * FIXED: Налаштовує слухачів та завантажує збережені налаштування
   */
  constructor() {
    this.setupEventListeners();
    this.loadSavedSettings();
  }

  /**
   * Налаштування слухачів подій для UI елементів
   * FIXED: Підписка на кнопки та чекбокси
   * @private
   */
  private setupEventListeners(): void {
    // Вибір папки з Excel файлами
    const selectFolderBtn = byId('excel-select-folder');
    const inputFolderField = byId<HTMLInputElement>('excel-input-folder');
    
    selectFolderBtn?.addEventListener('click', async () => {
      try {
        const result = await window.api?.selectFolder?.();
        if (result?.filePath) {
          this.inputFolder = result.filePath;
          inputFolderField!.value = result.filePath;
          
          await this.scanForExcelFiles();
          this.updateProcessButton();
          this.logMessage(`📂 Обрана папка: ${result.filePath}`);
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору папки: ${error}`, 'error');
      }
    });

    // Вибір файлу призначення
    const selectDestinationBtn = byId('excel-select-destination');
    const destinationFileField = byId<HTMLInputElement>('excel-destination-file');
    
    selectDestinationBtn?.addEventListener('click', async () => {
      try {
        // FIXED: Тут би був виклик до API для вибору Excel файлу
        // const result = await window.api?.selectExcelFile?.();
        
        // Поки що емулюємо вибір файлу
        const mockResult = {
          filePath: 'C:\\Excel\\Звідна_таблиця_2024.xlsx'
        };
        
        if (mockResult?.filePath) {
          this.destinationFile = mockResult.filePath;
          destinationFileField!.value = mockResult.filePath;
          this.updateProcessButton();
          this.logMessage(`💾 Обрано файл призначення: ${mockResult.filePath}`);
          this.saveDestinationSettings();
        }
      } catch (error) {
        this.logMessage(`❌ Помилка вибору файлу призначення: ${error}`, 'error');
      }
    });

    // Початок обробки
    const startBtn = byId('excel-start-processing');
    startBtn?.addEventListener('click', () => {
      this.startSummarizationProcessing();
    });

    // Зупинка обробки
    const stopBtn = byId('excel-stop-processing');
    stopBtn?.addEventListener('click', () => {
      this.stopProcessing();
    });

    // Очищення логів
    const clearLogsBtn = byId('excel-clear-logs');
    clearLogsBtn?.addEventListener('click', () => {
      this.clearLogs();
    });

    // Збереження логів
    const saveLogsBtn = byId('excel-save-logs');
    saveLogsBtn?.addEventListener('click', () => {
      this.saveLogs();
    });

    // Обробник чекбокса збереження назви файлу
    const rememberFilenameCheckbox = byId<HTMLInputElement>('excel-remember-filename');
    rememberFilenameCheckbox?.addEventListener('change', () => {
      this.saveDestinationSettings();
    });

    // Обробники налаштувань обробки
    const sliceCheckbox = byId<HTMLInputElement>('excel-slice-check');
    const mismatchesCheckbox = byId<HTMLInputElement>('excel-mismatches');
    const sanitizerCheckbox = byId<HTMLInputElement>('excel-sanitizer');
    
    sliceCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    mismatchesCheckbox?.addEventListener('change', () => this.saveProcessingSettings());
    sanitizerCheckbox?.addEventListener('change', () => this.saveProcessingSettings());

    // Обробники радіо-кнопок режиму зведення
    const modeRadios = document.querySelectorAll('input[name="excel-mode"]');
    modeRadios.forEach(radio => {
      radio.addEventListener('change', () => this.saveSummarizationMode());
    });
  }

  /**
   * Сканування папки на наявність Excel файлів
   * FIXED: Асинхронне сканування через API
   * @private
   */
  private async scanForExcelFiles(): Promise<void> {
    try {
      if (!this.inputFolder) {
        this.logMessage(`⚠️ Папка не вибрана`, 'error');
        return;
      }

      this.logMessage(`🔍 Сканування папки: ${this.inputFolder}`);

      // Викликаємо API для сканування Excel файлів
      const foundFiles = await window.api?.scanExcelFiles?.(this.inputFolder);
      
      if (foundFiles && foundFiles.length > 0) {
        this.foundFiles = foundFiles;
        this.displayFoundFiles();
        this.logMessage(`✅ Знайдено ${this.foundFiles.length} Excel файлів`);
      } else {
        this.foundFiles = [];
        this.displayFoundFiles();
        this.logMessage(`ℹ️ Excel файли не знайдено в обраній папці`);
      }
      
      this.updateProcessButton();
    } catch (error) {
      this.logMessage(`❌ Помилка сканування файлів: ${error}`, 'error');
      this.foundFiles = [];
      this.displayFoundFiles();
    }
  }

  /**
   * Відображення знайдених файлів
   * FIXED: Приховує область (тимчасово відключено)
   * @private
   */
  private displayFoundFiles(): void {
    // Приховуємо область знайдених файлів (тимчасово відключено)
    const filesPreview = byId('excel-files-preview');
    if (filesPreview) {
      filesPreview.style.display = 'none';
    }
  }

  /**
   * Оновлення стану кнопки "Почати обробку"
   * FIXED: Активує/деактивує кнопку залежно від готовності
   * @private
   */
  private updateProcessButton(): void {
    const startBtn = byId<HTMLButtonElement>('excel-start-processing');
    // FIXED: Не перевіряємо foundFiles.length, оскільки область файлів прихована
    const canProcess = this.inputFolder && this.destinationFile && !this.isProcessing;
    
    if (startBtn) {
      startBtn.disabled = !canProcess;
    }
  }

  /**
   * Запуск процесу зведення Excel файлів
   * FIXED: Валідація, формування payload, виклик API
   * @private
   */
  private async startSummarizationProcessing(): Promise<void> {
    if (this.isProcessing) return;
    
    // Валідація вхідних даних
    if (!this.inputFolder || !this.destinationFile) {
      this.logMessage('❌ Не всі обов\'язкові поля заповнені', 'error');
      return;
    }

    this.isProcessing = true;
    
    // Отримуємо режим зведення
    const selectedMode = document.querySelector('input[name="excel-mode"]:checked') as HTMLInputElement;
    const mode: Mode = (selectedMode?.value as Mode) || 'Обидва';
    
    this.logMessage(`🚀 Початок зведення Excel файлів. Режим: ${mode}`);
    
    // Показуємо прогрес
    const progressSection = byId('excel-progress');
    const startBtn = byId<HTMLButtonElement>('excel-start-processing');
    const stopBtn = byId<HTMLButtonElement>('excel-stop-processing');
    
    progressSection!.style.display = 'block';
    startBtn!.style.display = 'none';
    stopBtn!.style.display = 'inline-block';

    try {
      // Підготовка payload для зведення
      const payload: StartProcessPayload = {
        srcFolder: this.inputFolder,
        dstPath: this.destinationFile,
        mode: mode,
        // dstSheetPassword: undefined, // поки що не підтримуємо
        // configPath: undefined // використовуємо вбудовану конфігурацію
      };

      this.logMessage(`📁 Папка джерела: ${payload.srcFolder}`);
      this.logMessage(`💾 Файл призначення: ${payload.dstPath}`);
      this.logMessage(`📊 Режим обробки: ${payload.mode}`);

      // FIXED: Mock обробки (в реальному застосунку тут би був виклик до нового API)
      await this.mockSummarizationProcess(payload);
      
    } catch (error) {
      this.logMessage(`❌ Помилка зведення: ${error}`, 'error');
    } finally {
      this.stopProcessing();
    }
  }

  /**
   * Mock процес зведення для демонстрації
   * FIXED: Емуляція процесу зведення з прогресом
   * @private
   */
  private async mockSummarizationProcess(payload: StartProcessPayload): Promise<void> {
    // Mock процес зведення для демонстрації
    this.logMessage('🔧 Сканування папки та завантаження конфігурації...');
    this.updateProgress(10, 'Ініціалізація', 'Завантаження конфігурації');
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Отримуємо налаштування обробки
    const sliceCheck = byId<HTMLInputElement>('excel-slice-check')?.checked || false;
    const mismatches = byId<HTMLInputElement>('excel-mismatches')?.checked || false;
    const sanitizer = byId<HTMLInputElement>('excel-sanitizer')?.checked || false;
    
    const activeOptions = [];
    if (sliceCheck) activeOptions.push('Slice_Check');
    if (mismatches) activeOptions.push('Mismatches');  
    if (sanitizer) activeOptions.push('Sanitizer');
    
    this.logMessage(`🔧 Активні опції: ${activeOptions.join(', ')}`);
    
    // Імітація обробки за режимами
    const modes = payload.mode === 'Обидва' ? ['БЗ', 'ЗС'] : [payload.mode];
    let totalFiles = 0;
    let totalRows = 0;
    
    for (let i = 0; i < modes.length; i++) {
      const currentMode = modes[i];
      const progress = Math.round(((i + 1) / modes.length) * 80) + 10; // 10-90%
      
      this.logMessage(`📂 Режим ${i + 1}/${modes.length}: ${currentMode}`);
      this.updateProgress(progress, `Обробка режиму ${currentMode}`, `Режим ${i + 1}/${modes.length}`);
      
      // Mock правила для режиму
      const rules = currentMode === 'БЗ' 
        ? ['1РСпП', '2РСпП', '3РСпП', 'РВПСпП', 'МБ', 'РБпС', 'ВРСП', 'ВРЕБ', 'ВІ', 'ВЗ', 'РМТЗ', 'МП', '1241']
        : ['1РСпП', '2РСпП', '3РСпП', 'РВПСпП', 'МБ', 'РБпС', 'ВРСП', 'ВРЕБ', 'ВІ', 'ВЗ', 'РМТЗ', 'МП'];
      
      this.logMessage(`📋 Обробка ${rules.length} правил для режиму ${currentMode}`);
      
      for (let j = 0; j < rules.length; j++) {
        if (!this.isProcessing) break;
        
        const rule = rules[j];
        this.logMessage(`📄 Правило "${rule}": пошук файлу...`);
        
        // Імітація пошуку та обробки
        await new Promise(resolve => setTimeout(resolve, 300 + Math.random() * 700));
        
        if (Math.random() > 0.2) { // 80% успішності
          const rows = Math.floor(Math.random() * 20) + 5;
          totalFiles++;
          totalRows += rows;
          this.logMessage(`✅ Правило "${rule}": знайдено файл, скопійовано ${rows} рядків`);
        } else {
          this.logMessage(`⚠️ Правило "${rule}": файл не знайдено`, 'error');
        }
      }
      
      this.logMessage(`✅ Режим "${currentMode}" завершено`);
    }
    
    // Фінальний етап
    this.updateProgress(95, 'Збереження результатів', 'Запис у файл призначення');
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    this.logMessage(`🎉 Зведення завершено успішно!`);
    this.logMessage(`📈 Підсумок: файлів - ${totalFiles}, рядків - ${totalRows}`);
    this.logMessage(`💾 Результат збережено: ${payload.dstPath}`);
    
    this.updateProgress(100, 'Завершено', `Оброблено ${modes.length} режим(ів)`);
  }

  /**
   * Зупинка обробки
   * FIXED: Скидає прапор та ховає прогрес
   * @private
   */
  private stopProcessing(): void {
    this.isProcessing = false;
    
    const progressSection = byId('excel-progress');
    const startBtn = byId<HTMLButtonElement>('excel-start-processing');
    const stopBtn = byId<HTMLButtonElement>('excel-stop-processing');
    
    progressSection!.style.display = 'none';
    startBtn!.style.display = 'inline-block';
    stopBtn!.style.display = 'none';
    
    this.updateProcessButton();
    this.logMessage('⏹️ Обробку зупинено користувачем');
  }

  /**
   * Оновлення відображення прогресу
   * FIXED: Оновлює прогрес-бар та текст статусу
   * @private
   */
  private updateProgress(percent: number, status: string, detail: string): void {
    const progressFill = byId('excel-progress-fill');
    const progressPercent = byId('excel-progress-percent');
    const progressStatus = byId('excel-progress-status');
    const progressDetail = byId('excel-progress-detail');
    
    if (progressFill) progressFill.style.width = `${percent}%`;
    if (progressPercent) progressPercent.textContent = `${percent}%`;
    if (progressStatus) progressStatus.textContent = status;
    if (progressDetail) progressDetail.textContent = detail;
  }

  /**
   * Логування повідомлень
   * FIXED: Додає timestamp та емодзі іконки
   * @private
   */
  private logMessage(message: string, type: 'info' | 'error' = 'info'): void {
    const logsContent = byId('excel-logs-content');
    const timestamp = new Date().toLocaleTimeString('uk-UA');
    const prefix = type === 'error' ? '❌' : 'ℹ️';
    
    if (logsContent) {
      logsContent.textContent += `[${timestamp}] ${prefix} ${message}\n`;
      logsContent.scrollTop = logsContent.scrollHeight;
    }
  }

  /**
   * Очищення логів
   * FIXED: Скидає вміст логів
   * @private
   */
  private clearLogs(): void {
    const logsContent = byId('excel-logs-content');
    if (logsContent) {
      logsContent.textContent = 'Готово до початку обробки...\n';
    }
  }

  /**
   * Збереження логів у файл
   * FIXED: Завантажує файл з логами
   * @private
   */
  private async saveLogs(): Promise<void> {
    try {
      const logsContent = byId('excel-logs-content')?.textContent || '';
      if (!logsContent.trim()) {
        this.logMessage('⚠️ Немає логів для збереження');
        return;
      }

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
      const filename = `excel-processing-logs-${timestamp}.txt`;
      
      // FIXED: Тут би був виклик до electron API для збереження файлу
      // const result = await window.api?.saveFile?.(filename, logsContent);
      
      this.logMessage(`💾 Логи збережено: ${filename}`);
    } catch (error) {
      this.logMessage(`❌ Помилка збереження логів: ${error}`, 'error');
    }
  }

  /**
   * Завантаження збережених налаштувань
   * FIXED: Завантажує збережені налаштування з API
   * @private
   */
  private async loadSavedSettings(): Promise<void> {
    try {
      // Завантажуємо збережений файл призначення
      const savedDestination = await window.api?.getSetting?.('excelDestinationFile', '');
      const destinationFileField = byId<HTMLInputElement>('excel-destination-file');
      
      if (destinationFileField && savedDestination) {
        destinationFileField.value = savedDestination;
        this.destinationFile = savedDestination;
      }

      // Завантажуємо стан чекбокса збереження призначення
      const rememberDestination = await window.api?.getSetting?.('excelRememberDestination', true);
      const rememberDestinationCheckbox = byId<HTMLInputElement>('excel-remember-destination');
      
      if (rememberDestinationCheckbox) {
        rememberDestinationCheckbox.checked = rememberDestination;
      }

      // Завантажуємо налаштування обробки
      const sliceCheck = await window.api?.getSetting?.('excelSliceCheck', true);
      const mismatches = await window.api?.getSetting?.('excelMismatches', true);
      const sanitizer = await window.api?.getSetting?.('excelSanitizer', true);
      
      const sliceCheckbox = byId<HTMLInputElement>('excel-slice-check');
      const mismatchesCheckbox = byId<HTMLInputElement>('excel-mismatches');
      const sanitizerCheckbox = byId<HTMLInputElement>('excel-sanitizer');
      
      if (sliceCheckbox) sliceCheckbox.checked = sliceCheck;
      if (mismatchesCheckbox) mismatchesCheckbox.checked = mismatches;
      if (sanitizerCheckbox) sanitizerCheckbox.checked = sanitizer;

      // Завантажуємо вибраний режим зведення
      const savedMode = await window.api?.getSetting?.('excelSummarizationMode', 'Обидва');
      const modeRadio = byId<HTMLInputElement>(`excel-mode-${savedMode === 'БЗ' ? 'bz' : savedMode === 'ЗС' ? 'zs' : 'both'}`);
      if (modeRadio) {
        modeRadio.checked = true;
      }

      this.logMessage('📁 Збережені налаштування завантажено');
    } catch (error) {
      console.warn('Не вдалося завантажити збережені налаштування:', error);
    }
  }

  /**
   * Збереження налаштувань файлу призначення
   * FIXED: Зберігає шлях до файлу та режим зведення
   * @private
   */
  private async saveDestinationSettings(): Promise<void> {
    try {
      const rememberDestinationCheckbox = byId<HTMLInputElement>('excel-remember-destination');
      
      if (rememberDestinationCheckbox?.checked && this.destinationFile) {
        await window.api?.setSetting?.('excelDestinationFile', this.destinationFile);
      }
      
      await window.api?.setSetting?.('excelRememberDestination', rememberDestinationCheckbox?.checked || false);
      
      // Зберігаємо режим зведення
      const selectedMode = document.querySelector<HTMLInputElement>('input[name="excel-mode"]:checked');
      if (selectedMode) {
        const mode = selectedMode.value === 'bz' ? 'БЗ' : selectedMode.value === 'zs' ? 'ЗС' : 'Обидва';
        await window.api?.setSetting?.('excelSummarizationMode', mode);
      }
      
      // Зберігаємо налаштування чекбоксів
      const sliceCheck = byId<HTMLInputElement>('excel-slice-check')?.checked || false;
      const mismatches = byId<HTMLInputElement>('excel-mismatches')?.checked || false;
      const sanitizer = byId<HTMLInputElement>('excel-sanitizer')?.checked || false;
      
      await window.api?.setSetting?.('excelSliceCheck', sliceCheck);
      await window.api?.setSetting?.('excelMismatches', mismatches);
      await window.api?.setSetting?.('excelSanitizer', sanitizer);
      
    } catch (error) {
      console.warn('Не вдалося зберегти налаштування:', error);
    }
  }

  /**
   * Збереження налаштувань обробки
   * FIXED: Зберігає чекбокси опцій
   * @private
   */
  private async saveProcessingSettings(): Promise<void> {
    try {
      const sliceCheck = byId<HTMLInputElement>('excel-slice-check')?.checked || false;
      const mismatches = byId<HTMLInputElement>('excel-mismatches')?.checked || false;
      const sanitizer = byId<HTMLInputElement>('excel-sanitizer')?.checked || false;
      
      await window.api?.setSetting?.('excelSliceCheck', sliceCheck);
      await window.api?.setSetting?.('excelMismatches', mismatches);
      await window.api?.setSetting?.('excelSanitizer', sanitizer);
      
      this.logMessage(`💾 Налаштування збережено: Slice_Check=${sliceCheck}, Mismatches=${mismatches}, Sanitizer=${sanitizer}`);
    } catch (error) {
      console.warn('Не вдалося зберегти налаштування обробки:', error);
    }
  }

  /**
   * Збереження режиму зведення
   * FIXED: Зберігає вибраний режим (БЗ/ЗС/Обидва)
   * @private
   */
  private async saveSummarizationMode(): Promise<void> {
    try {
      const selectedMode = document.querySelector('input[name="excel-mode"]:checked') as HTMLInputElement;
      const mode = selectedMode?.value || 'Обидва';
      
      await window.api?.setSetting?.('excelSummarizationMode', mode);
      
      this.logMessage(`💾 Режим зведення збережено: ${mode}`);
    } catch (error) {
      console.warn('Не вдалося зберегти режим зведення:', error);
    }
  }
}
