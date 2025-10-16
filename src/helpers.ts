/**
 * Helper utilities для роботи з DOM
 * FIXED: Винесено з main.ts для переви використання
 */

import { DOM_IDS, LOCALE } from './constants';

/**
 * FIXED: Helper функція для отримання елемента за ID з типізацією
 * Винесено з main.ts (рядок 87)
 * @param id - ID елемента
 * @returns Елемент з вказаним типом або null
 */
export const byId = <T extends HTMLElement>(id: string): T | null => 
  document.getElementById(id) as T | null;

/**
 * FIXED: Helper функція для логування в UI
 * Винесено з main.ts (рядки 88-94)
 * @param msg - Повідомлення для логування
 */
export const log = (msg: string): void => {
  const el = byId<HTMLPreElement>(DOM_IDS.LOG_BODY);
  if (!el) return;
  const time = new Date().toLocaleTimeString(LOCALE);
  el.textContent += `[${time}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
};

/**
 * FIXED: Додано helper для форматування дати
 * @param date - Дата для форматування
 * @returns Відформатована дата
 */
export const formatDate = (date: Date): string => {
  return date.toISOString().split('T')[0]; // YYYY-MM-DD
};

/**
 * FIXED: Додано helper для форматування часу
 * @param date - Дата для форматування
 * @returns Відформатований час
 */
export const formatTime = (date: Date): string => {
  return date.toTimeString().split(' ')[0].replace(/:/g, '-'); // HH-MM-SS
};

/**
 * FIXED: Додано helper для створення timestamp
 * @returns ISO timestamp without milliseconds
 */
export const getTimestamp = (): string => {
  return new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
};
