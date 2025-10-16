/**
 * SourceSelectionManager - Управління вибором джерела файлів
 * FIXED: Винесено з main.ts (рядки 352-432)
 * 
 * Відповідальність:
 * - Перемикання між режимами вибору (один файл / множинні файли / папка)
 * - Відображення відповідних input полів
 * - Виклик діалогів вибору файлів/папок
 * - Надання інтерфейсу для отримання вибраних файлів
 */

import { DOM_IDS, EVENT_TYPES, SOURCE_TYPES } from './constants';

// FIXED: Додано helper функцію byId
const byId = <T extends HTMLElement>(id: string): T | null => 
  document.getElementById(id) as T | null;

// FIXED: Додано helper функцію log
const log = (msg: string): void => {
  const el = byId<HTMLPreElement>(DOM_IDS.LOG_BODY);
  if (!el) return;
  const time = new Date().toLocaleTimeString();
  el.textContent += `[${time}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
};

/**
 * Клас для управління вибором джерела файлів (single/multiple/folder)
 */
export class SourceSelectionManager {
  private sourceRadios: NodeListOf<HTMLInputElement>;
  private singleFileInput: HTMLElement | null;
  private multipleFilesInput: HTMLElement | null;
  private folderInput: HTMLElement | null;

  /**
   * Конструктор - ініціалізує елементи та налаштовує обробники подій
   */
  constructor() {
    this.sourceRadios = document.querySelectorAll<HTMLInputElement>('input[name="source-type"]');
    this.singleFileInput = byId(DOM_IDS.SINGLE_FILE_INPUT);
    this.multipleFilesInput = byId(DOM_IDS.MULTIPLE_FILES_INPUT);
    this.folderInput = byId(DOM_IDS.FOLDER_INPUT);
    
    this.bindEvents();
  }

  /**
   * FIXED: Приватний метод для прив'язки обробників подій
   */
  private bindEvents(): void {
    // Обробник зміни radio buttons
    this.sourceRadios.forEach(radio => {
      radio.addEventListener(EVENT_TYPES.CHANGE, () => {
        this.handleSourceChange(radio.value);
      });
    });

    // Обробник для кнопки вибору папки
    byId(DOM_IDS.CHOOSE_FOLDER)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.selectFolder();
    });
  }

  /**
   * Обробляє зміну типу джерела - показує відповідний input
   * @param sourceType - Тип джерела ('single-file' | 'multiple-files' | 'folder')
   */
  private handleSourceChange(sourceType: string): void {
    // Ховаємо всі inputs
    if (this.singleFileInput) this.singleFileInput.style.display = 'none';
    if (this.multipleFilesInput) this.multipleFilesInput.style.display = 'none';
    if (this.folderInput) this.folderInput.style.display = 'none';

    // Показуємо потрібний input
    switch (sourceType) {
      case SOURCE_TYPES.SINGLE_FILE:
        if (this.singleFileInput) this.singleFileInput.style.display = 'block';
        break;
      case SOURCE_TYPES.MULTIPLE_FILES:
        if (this.multipleFilesInput) this.multipleFilesInput.style.display = 'block';
        break;
      case SOURCE_TYPES.FOLDER:
        if (this.folderInput) this.folderInput.style.display = 'block';
        break;
    }
  }

  /**
   * Викликає діалог вибору папки через IPC
   * @async
   */
  private async selectFolder(): Promise<void> {
    try {
      const folderPath = await window.api?.selectBatchDirectory?.();
      if (folderPath) {
        const folderInput = byId<HTMLInputElement>(DOM_IDS.FOLDER_PATH);
        if (folderInput) {
          folderInput.value = folderPath;
        }
        log(`📁 Обрано папку: ${folderPath}`);
      }
    } catch (error) {
      console.error('Помилка вибору папки:', error);
      log(`❌ Помилка вибору папки: ${error}`);
    }
  }

  /**
   * Повертає поточний вибраний тип джерела
   * @returns Тип джерела або 'single-file' за замовчуванням
   * @public
   */
  public getSelectedSource(): string {
    const checkedRadio = document.querySelector('input[name="source-type"]:checked') as HTMLInputElement;
    return checkedRadio ? checkedRadio.value : SOURCE_TYPES.SINGLE_FILE;
  }

  /**
   * Повертає список вибраних файлів або шлях до папки
   * @returns Масив File об'єктів або масив з шляхом до папки
   * @public
   */
  public getSelectedFiles(): File[] | string[] {
    const sourceType = this.getSelectedSource();
    
    switch (sourceType) {
      case SOURCE_TYPES.SINGLE_FILE:
        const singleFile = byId<HTMLInputElement>(DOM_IDS.WORD_FILE);
        return singleFile?.files ? Array.from(singleFile.files) : [];
        
      case SOURCE_TYPES.MULTIPLE_FILES:
        const multipleFiles = byId<HTMLInputElement>(DOM_IDS.WORD_FILES);
        return multipleFiles?.files ? Array.from(multipleFiles.files) : [];
        
      case SOURCE_TYPES.FOLDER:
        const folderPath = byId<HTMLInputElement>(DOM_IDS.FOLDER_PATH)?.value;
        return folderPath ? [folderPath] : [];
        
      default:
        return [];
    }
  }
}
