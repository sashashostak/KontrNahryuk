/**
 * SectionManager - Управління відображенням секцій форми
 * FIXED: Винесено з main.ts (рядки 437-481)
 * 
 * Відповідальність:
 * - Показ/приховування секції вибору Excel файлу
 * - Прив'язка до чекбокса "Розпорядження"
 * - Виклик діалогу вибору Excel файлу
 */

import { DOM_IDS, EVENT_TYPES } from './constants';

// FIXED: Додано helper функцію byId (використовувалась глобально в main.ts)
const byId = <T extends HTMLElement>(id: string): T | null => 
  document.getElementById(id) as T | null;

// FIXED: Додано helper функцію log (використовувалась глобально в main.ts)
const log = (msg: string): void => {
  const el = byId<HTMLPreElement>(DOM_IDS.LOG_BODY);
  if (!el) return;
  const time = new Date().toLocaleTimeString();
  el.textContent += `[${time}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
};

/**
 * Клас для управління умовним відображенням секцій
 * Показує Excel секцію тільки коли активний режим "Розпорядження"
 */
export class SectionManager {
  private orderCheckbox: HTMLInputElement | null;
  private excelSection: HTMLElement | null;

  /**
   * Конструктор - ініціалізує елементи та налаштовує обробники подій
   */
  constructor() {
    this.orderCheckbox = byId<HTMLInputElement>(DOM_IDS.T_ORDER);
    this.excelSection = byId(DOM_IDS.EXCEL_SECTION);
    this.bindEvents();
  }

  /**
   * FIXED: Приватний метод для прив'язки обробників подій
   */
  private bindEvents(): void {
    // Обробник зміни чекбокса "Розпорядження"
    this.orderCheckbox?.addEventListener(EVENT_TYPES.CHANGE, () => {
      this.toggleExcelSection();
    });

    // Обробник для кнопки вибору Excel файлу
    byId(DOM_IDS.CHOOSE_EXCEL)?.addEventListener(EVENT_TYPES.CLICK, () => {
      this.selectExcelFile();
    });
  }

  /**
   * Показує або ховає Excel секцію залежно від стану чекбокса
   */
  private toggleExcelSection(): void {
    if (this.excelSection) {
      this.excelSection.style.display = this.orderCheckbox?.checked ? 'block' : 'none';
    }
  }

  /**
   * Викликає діалог вибору Excel файлу через IPC
   * @async
   */
  private async selectExcelFile(): Promise<void> {
    try {
      const filePath = await window.api?.selectExcelFile?.();
      if (filePath) {
        const excelInput = byId<HTMLInputElement>(DOM_IDS.EXCEL_PATH);
        if (excelInput) {
          excelInput.value = filePath;
        }
        log(`📊 Обрано Excel файл: ${filePath}`);
      }
    } catch (error) {
      console.error('Помилка вибору Excel файлу:', error);
      log(`❌ Помилка вибору Excel файлу: ${error}`);
    }
  }
}
