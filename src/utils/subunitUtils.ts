/**
 * subunitUtils.ts - Утиліти для роботи з підрозділами та Excel
 * 
 * Функції:
 * - Нормалізація назв підрозділів
 * - Перевірка порожніх рядків
 * - Робота з колонками Excel
 */

import { PROCESSING_OPTIONS } from '../config/constants';

/**
 * Нормалізує назву підрозділу для використання як ключ
 * 
 * Правила:
 * - trim() - видалення пробілів з початку та кінця
 * - toLowerCase() - приведення до нижнього регістру (якщо NORMALIZE_KEYS = true)
 * - null якщо значення порожнє
 * 
 * Приклад:
 * "  1РСпП  " → "1рспп"
 * "ВРСП" → "врсп"
 * "" → null
 * 
 * @param value - Значення комірки (може бути будь-якого типу)
 * @returns Нормалізована строка або null
 */
export function normalizeSubunitKey(value: any): string | null {
  // Перевірка на null/undefined
  if (value === null || value === undefined) {
    return null;
  }
  
  // Приведення до строки та trim
  const str = String(value).trim();
  
  // Перевірка на порожню строку
  if (str === '') {
    return null;
  }
  
  // Нормалізація (toLowerCase) якщо увімкнено
  return PROCESSING_OPTIONS.NORMALIZE_KEYS 
    ? str.toLowerCase() 
    : str;
}

/**
 * Перевіряє чи є рядок Excel порожнім
 * 
 * Рядок вважається порожнім якщо:
 * - Об'єкт row відсутній
 * - Масив values відсутній
 * - Всі комірки null/undefined/пусті строки
 * 
 * @param row - Рядок ExcelJS (Row object)
 * @returns true якщо рядок порожній
 */
export function isEmptyRow(row: any): boolean {
  if (!row) {
    return true;
  }
  
  // Отримуємо values рядка
  const cells = row.values;
  if (!cells) {
    return true;
  }
  
  // Перевіряємо чи всі комірки порожні
  // values[0] зазвичай undefined (індекси починаються з 1)
  return cells.every((cell: any) => 
    cell === null || 
    cell === undefined || 
    String(cell).trim() === ''
  );
}

/**
 * Перетворює назву колонки Excel на номер
 * 
 * Приклади:
 * A = 1
 * B = 2
 * Z = 26
 * AA = 27
 * AB = 28
 * 
 * @param column - Назва колонки (A, B, C, ..., AA, AB, ...)
 * @returns Номер колонки (1-based)
 */
export function columnToNumber(column: string): number {
  let result = 0;
  const upper = column.toUpperCase();
  
  for (let i = 0; i < upper.length; i++) {
    const charCode = upper.charCodeAt(i);
    const value = charCode - 64; // A=1, B=2, ...
    result = result * 26 + value;
  }
  
  return result;
}

/**
 * Перетворює номер колонки на назву
 * 
 * Приклади:
 * 1 = A
 * 2 = B
 * 26 = Z
 * 27 = AA
 * 
 * @param num - Номер колонки (1-based)
 * @returns Назва колонки
 */
export function numberToColumn(num: number): string {
  let column = '';
  let n = num;
  
  while (n > 0) {
    const remainder = (n - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    n = Math.floor((n - 1) / 26);
  }
  
  return column;
}

/**
 * Отримує список колонок для копіювання
 * 
 * Приклад:
 * getColumnsRange('C', 'H') → ['C', 'D', 'E', 'F', 'G', 'H']
 * 
 * @param start - Початкова колонка (наприклад: 'C')
 * @param end - Кінцева колонка (наприклад: 'H')
 * @returns Масив назв колонок
 */
export function getColumnsRange(start: string, end: string): string[] {
  const startNum = columnToNumber(start);
  const endNum = columnToNumber(end);
  const columns: string[] = [];
  
  for (let i = startNum; i <= endNum; i++) {
    columns.push(numberToColumn(i));
  }
  
  return columns;
}

/**
 * Перевіряє чи файл є тимчасовим Excel файлом
 * 
 * Тимчасові файли Excel створюються при відкритті файлу
 * та мають назву ~$filename.xlsx
 * 
 * @param fileName - Назва файлу
 * @returns true якщо це тимчасовий файл
 */
export function isTempFile(fileName: string): boolean {
  return fileName.startsWith('~$');
}

/**
 * Перевіряє чи файл є Excel файлом
 * 
 * @param fileName - Назва файлу
 * @returns true якщо це Excel файл (.xlsx або .xls)
 */
export function isExcelFile(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return lower.endsWith('.xlsx') || lower.endsWith('.xls');
}

/**
 * Валідує діапазон колонок
 * 
 * @param start - Початкова колонка
 * @param end - Кінцева колонка
 * @throws Error якщо діапазон невалідний
 */
export function validateColumnRange(start: string, end: string): void {
  const startNum = columnToNumber(start);
  const endNum = columnToNumber(end);
  
  if (startNum > endNum) {
    throw new Error(`Невалідний діапазон колонок: ${start} > ${end}`);
  }
  
  if (startNum <= 0 || endNum <= 0) {
    throw new Error(`Невалідні колонки: ${start}, ${end}`);
  }
}
