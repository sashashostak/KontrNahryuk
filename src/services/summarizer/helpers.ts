/**
 * Хелпери для роботи з Excel файлами та файловою системою
 */

import * as path from 'path';
import { FileInfo, FileOptions, BlockCoords } from './types';

/**
 * Нормалізує назву файлу для порівняння
 * '_' & '-' → ' ', trim, схлопування пробілів
 */
export function normalizeName(name: string): string {
  return name
    .replace(/[_-]/g, ' ')        // замінити _ та - на пробіли
    .replace(/\s+/g, ' ')         // схлопнути множинні пробіли
    .trim()                       // прибрати пробіли з початку і кінця
    .toLowerCase();               // до нижнього регістру
}

/**
 * Знаходить останній за часом зміни файл, що містить усі токени
 * @param files - список файлів для пошуку
 * @param tokens - токени для пошуку в назвах
 * @returns шлях до найновішого файлу або null
 */
export function findLatestByTokens(files: FileInfo[], tokens: string[]): FileInfo | null {
  if (!files || files.length === 0 || !tokens || tokens.length === 0) {
    return null;
  }

  // Фільтруємо файли, які містять всі токени
  const matchingFiles = files.filter(file => {
    const normalizedName = file.normalized;
    return tokens.every(token => 
      normalizedName.includes(normalizeName(token))
    );
  });

  if (matchingFiles.length === 0) {
    return null;
  }

  // Сортуємо за часом зміни (найновіші першими) і беремо перший
  return matchingFiles.sort((a, b) => 
    b.lastModified.getTime() - a.lastModified.getTime()
  )[0];
}

/**
 * Отримує список Excel файлів з директорії
 * @param dir - шлях до директорії
 * @param options - опції пошуку
 * @returns масив інформації про файли
 */
export async function listExcelFiles(dir: string, options: FileOptions = {}): Promise<FileInfo[]> {
  try {
    // В реальному застосунку тут би був виклик до Node.js fs API
    // Поки що повертаємо mock дані для тестування
    const mockFiles: FileInfo[] = [
      {
        path: path.join(dir, '1РСпП_звіт_грудень_2023.xlsx'),
        name: '1РСпП_звіт_грудень_2023.xlsx',
        normalized: normalizeName('1РСпП_звіт_грудень_2023.xlsx'),
        lastModified: new Date('2023-12-31'),
        size: 50000
      },
      {
        path: path.join(dir, 'МБ-звіт-листопад-2023.xlsx'),
        name: 'МБ-звіт-листопад-2023.xlsx',
        normalized: normalizeName('МБ-звіт-листопад-2023.xlsx'),
        lastModified: new Date('2023-11-30'),
        size: 45000
      },
      {
        path: path.join(dir, 'РВПСпП_звіт_січень_2024.xlsx'),
        name: 'РВПСпП_звіт_січень_2024.xlsx',
        normalized: normalizeName('РВПСпП_звіт_січень_2024.xlsx'),
        lastModified: new Date('2024-01-31'),
        size: 52000
      }
    ];

    return mockFiles.filter(file => {
      const ext = path.extname(file.name).toLowerCase();
      
      // Завжди включаємо .xlsx та .xlsm
      if (ext === '.xlsx' || ext === '.xlsm') {
        return true;
      }
      
      // Опційно включаємо .xls та .xlsb якщо дозволено
      if (options.includeXls && ext === '.xls') {
        return true;
      }
      
      if (options.includeXlsb && ext === '.xlsb') {
        return true;
      }
      
      return false;
    });
  } catch (error) {
    console.error(`Помилка читання директорії ${dir}:`, error);
    return [];
  }
}

/**
 * Читає робочий аркуш Excel файлу
 * @param filePath - шлях до файлу
 * @param sheetName - назва аркуша
 * @returns 2D масив комірок (рядки × колонки)
 */
export async function readWorksheet(filePath: string, sheetName: string): Promise<string[][]> {
  try {
    // В реальному застосунку тут би був виклик до xlsx бібліотеки
    // const workbook = XLSX.readFile(filePath);
    // const worksheet = workbook.Sheets[sheetName];
    // return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    // Mock дані для тестування
    console.log(`Читання аркуша "${sheetName}" з файлу: ${filePath}`);
    
    // Повертаємо mock дані що імітують структуру Excel
    return [
      ['', '', 'Підрозділ', 'Колонка D', 'Колонка E', 'Колонка F', 'Колонка G', 'Колонка H'],
      ['', '', '1РСпП', '100', '200', '300', '400', '500'],
      ['', '', '1РСпП', '110', '210', '310', '410', '510'],
      ['', '', '1РСпП', '120', '220', '320', '420', '520'],
      ['', '', 'МБ', '150', '250', '350', '450', '550'],
      ['', '', 'МБ', '160', '260', '360', '460', '560'],
      ['', '', 'РВПСпП', '180', '280', '380', '480', '580'],
    ];
  } catch (error) {
    console.error(`Помилка читання файлу ${filePath}:`, error);
    throw new Error(`Не вдалося прочитати аркуш "${sheetName}" з файлу "${filePath}": ${error}`);
  }
}

/**
 * Записує зміни в робочий аркуш Excel файлу
 * @param filePath - шлях до файлу
 * @param sheetName - назва аркуша
 * @param updater - функція для оновлення даних
 */
export async function writeWorksheet(
  filePath: string, 
  sheetName: string, 
  updater: (grid: string[][]) => void
): Promise<void> {
  try {
    // В реальному застосунку:
    // 1. Читаємо існуючий файл або створюємо новий
    // 2. Отримуємо аркуш або створюємо його
    // 3. Застосовуємо зміни через updater
    // 4. Зберігаємо файл
    
    console.log(`Запис в аркуш "${sheetName}" файлу: ${filePath}`);
    
    // Mock для тестування
    const mockGrid: string[][] = [];
    updater(mockGrid);
    
    console.log(`Записано ${mockGrid.length} рядків в аркуш "${sheetName}"`);
  } catch (error) {
    console.error(`Помилка запису в файл ${filePath}:`, error);
    throw new Error(`Не вдалося записати в аркуш "${sheetName}" файлу "${filePath}": ${error}`);
  }
}

/**
 * Знаходить перший суцільний блок з ключем в колонці
 * @param grid - 2D масив даних
 * @param keyCol - індекс колонки з ключем (0-based)
 * @param key - значення ключа для пошуку
 * @param left - ліва межа області (0-based)
 * @param right - права межа області (0-based)
 * @returns координати блоку або null
 */
export function findFirstContiguousBlock(
  grid: string[][], 
  keyCol: number, 
  key: string, 
  left: number, 
  right: number
): BlockCoords | null {
  if (!grid || grid.length === 0) {
    return null;
  }

  const normalizedKey = key.trim().toLowerCase();
  let startRow = -1;
  let endRow = -1;

  // Шукаємо перший збіг
  for (let row = 0; row < grid.length; row++) {
    if (!grid[row] || keyCol >= grid[row].length) {
      continue;
    }

    const cellValue = (grid[row][keyCol] || '').trim().toLowerCase();
    
    if (cellValue === normalizedKey) {
      if (startRow === -1) {
        startRow = row;
      }
      endRow = row;
    } else if (startRow !== -1) {
      // Знайшли кінець суцільного блоку
      break;
    }
  }

  if (startRow === -1) {
    return null;
  }

  return {
    r1: startRow,
    r2: endRow,
    c1: left,
    c2: right
  };
}

/**
 * Знаходить всі суцільні блоки з ключем в колонці
 * @param grid - 2D масив даних
 * @param keyCol - індекс колонки з ключем (0-based)
 * @param key - значення ключа для пошуку
 * @param left - ліва межа області (0-based)
 * @param right - права межа області (0-based)
 * @returns масив координат блоків
 */
export function findAllContiguousBlocks(
  grid: string[][], 
  keyCol: number, 
  key: string, 
  left: number, 
  right: number
): BlockCoords[] {
  if (!grid || grid.length === 0) {
    return [];
  }

  const normalizedKey = key.trim().toLowerCase();
  const blocks: BlockCoords[] = [];
  let currentBlock: { start: number; end: number } | null = null;

  for (let row = 0; row < grid.length; row++) {
    if (!grid[row] || keyCol >= grid[row].length) {
      if (currentBlock) {
        blocks.push({
          r1: currentBlock.start,
          r2: currentBlock.end,
          c1: left,
          c2: right
        });
        currentBlock = null;
      }
      continue;
    }

    const cellValue = (grid[row][keyCol] || '').trim().toLowerCase();
    
    if (cellValue === normalizedKey) {
      if (!currentBlock) {
        currentBlock = { start: row, end: row };
      } else {
        currentBlock.end = row;
      }
    } else {
      if (currentBlock) {
        blocks.push({
          r1: currentBlock.start,
          r2: currentBlock.end,
          c1: left,
          c2: right
        });
        currentBlock = null;
      }
    }
  }

  // Додаємо останній блок якщо він існує
  if (currentBlock) {
    blocks.push({
      r1: currentBlock.start,
      r2: currentBlock.end,
      c1: left,
      c2: right
    });
  }

  return blocks;
}

/**
 * Копіює дані з одного блоку в інший і очищає "хвіст"
 * @param srcGrid - джерельна сітка
 * @param srcBlock - координати джерельного блоку
 * @param dstGrid - цільова сітка
 * @param dstBlock - координати цільового блоку
 * @returns кількість скопійованих рядків
 */
export function copySlice(
  srcGrid: string[][], 
  srcBlock: BlockCoords, 
  dstGrid: string[][], 
  dstBlock: BlockCoords
): number {
  let copiedRows = 0;

  // Обчислюємо розміри блоків
  const srcRows = srcBlock.r2 - srcBlock.r1 + 1;
  const srcCols = srcBlock.c2 - srcBlock.c1 + 1;
  const dstRows = dstBlock.r2 - dstBlock.r1 + 1;
  const dstCols = dstBlock.c2 - dstBlock.c1 + 1;

  // Копіюємо дані з джерела в призначення
  for (let row = 0; row < Math.min(srcRows, dstRows); row++) {
    const srcRowIdx = srcBlock.r1 + row;
    const dstRowIdx = dstBlock.r1 + row;

    // Розширюємо масив якщо потрібно
    while (dstGrid.length <= dstRowIdx) {
      dstGrid.push([]);
    }

    for (let col = 0; col < Math.min(srcCols, dstCols); col++) {
      const srcColIdx = srcBlock.c1 + col;
      const dstColIdx = dstBlock.c1 + col;

      // Розширюємо рядок якщо потрібно
      while (dstGrid[dstRowIdx].length <= dstColIdx) {
        dstGrid[dstRowIdx].push('');
      }

      const srcValue = (srcGrid[srcRowIdx] && srcGrid[srcRowIdx][srcColIdx]) || '';
      dstGrid[dstRowIdx][dstColIdx] = srcValue;
    }
    
    copiedRows++;
  }

  // Очищуємо "хвіст" у призначенні якщо воно більше джерела
  if (dstRows > srcRows) {
    for (let row = srcRows; row < dstRows; row++) {
      const dstRowIdx = dstBlock.r1 + row;
      
      if (dstGrid[dstRowIdx]) {
        for (let col = 0; col < dstCols; col++) {
          const dstColIdx = dstBlock.c1 + col;
          if (dstColIdx < dstGrid[dstRowIdx].length) {
            dstGrid[dstRowIdx][dstColIdx] = '';
          }
        }
      }
    }
  }

  return copiedRows;
}

/**
 * Конвертує 1-based індекс колонки (як в Excel) в 0-based
 */
export function col1to0(col1based: number): number {
  return Math.max(0, col1based - 1);
}

/**
 * Конвертує 0-based індекс колонки в 1-based (як в Excel)
 */
export function col0to1(col0based: number): number {
  return col0based + 1;
}

/**
 * Конвертує номер колонки в літерне позначення Excel (A, B, C, ...)
 */
export function colNumberToLetter(colNumber: number): string {
  let result = '';
  while (colNumber > 0) {
    const remainder = (colNumber - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    colNumber = Math.floor((colNumber - 1) / 26);
  }
  return result;
}