/**
 * ExcelReader.ts - Читання Excel файлів та пошук маркера "Придані"
 * Витяг даних з колонок E (ПІБ), F (псевдо) після знайденого маркера
 */

import ExcelJS from 'exceljs'
import { SheetInfo, SheetDateParser } from './SheetDateParser'

export interface ExcelRowData {
  row: number
  pib: string              // колонка E - обов'язкова
  pseudo: string | null    // колонка F
}

export interface SheetProcessResult {
  sheetName: string
  date: Date | null
  markerRow: number | null
  dataRows: ExcelRowData[]
  warnings: string[]
  error?: string
}

export interface FileProcessResult {
  filePath: string
  fileName: string
  processed: boolean
  sheets: SheetProcessResult[]
  totalRows: number
  error?: string
}

export class ExcelReader {
  private static readonly MARKER_TEXT = 'придані';
  private static readonly MARKER_COLUMN = 5; // Колонка E (1-based: A=1, B=2, C=3, D=4, E=5)
  private static readonly PIB_COLUMN = 5;     // Колонка E
  private static readonly PSEUDO_COLUMN = 6;  // Колонка F
  private static readonly STATUS_COLUMN = 8;  // Колонка H для статусів
  private static readonly MAX_EMPTY_ROWS = 10; // Максимум порожніх рядків підряд

  /**
   * Обробляє один Excel файл
   */
  static async processFile(filePath: string): Promise<FileProcessResult> {
    const fileName = filePath.split(/[/\\]/).pop() || '';
    
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      
      const sheetResults: SheetProcessResult[] = [];
      let totalRows = 0;

      // Отримуємо всі видимі листи
      const worksheets = workbook.worksheets.filter(ws => ws.state === 'visible');
      const sheetNames = worksheets.map(ws => ws.name);
      
      // Парсимо дати з назв листів
      const sheetInfos = SheetDateParser.processSheets(sheetNames);
      const sortedSheets = SheetDateParser.sortSheetsByDate(sheetInfos);

      // Обробляємо листи в порядку дат
      for (const sheetInfo of sortedSheets) {
        const worksheet = worksheets.find(ws => ws.name === sheetInfo.name);
        if (!worksheet) continue;

        const sheetResult = await this.processSheet(worksheet, sheetInfo);
        sheetResults.push(sheetResult);
        totalRows += sheetResult.dataRows.length;
      }

      return {
        filePath,
        fileName,
        processed: true,
        sheets: sheetResults,
        totalRows
      };

    } catch (error) {
      return {
        filePath,
        fileName,
        processed: false,
        sheets: [],
        totalRows: 0,
        error: error instanceof Error ? error.message : 'Невідома помилка'
      };
    }
  }

  /**
   * Обробляє один лист Excel
   */
  private static async processSheet(worksheet: ExcelJS.Worksheet, sheetInfo: SheetInfo): Promise<SheetProcessResult> {
    const result: SheetProcessResult = {
      sheetName: sheetInfo.name,
      date: sheetInfo.date,
      markerRow: null,
      dataRows: [],
      warnings: []
    };

    try {
      // Перевіряємо дату тільки якщо вона є
      if (sheetInfo.date && !SheetDateParser.isReasonableDate(sheetInfo.date)) {
        result.error = `Дата ${SheetDateParser.formatDate(sheetInfo.date)} виходить за розумні межі`;
        return result;
      }

      // Шукаємо маркер "Придані" в колонці E
      const markerRow = this.findMarker(worksheet);
      let dataRows: ExcelRowData[] = [];

      if (markerRow) {
        result.markerRow = markerRow;
        // Читаємо дані після маркера (основний спосіб)
        dataRows = this.extractDataRows(worksheet, markerRow + 1);
      }

      // Додатково шукаємо бійців зі статусом "Придані" в колонці H
      const statusRows = this.extractStatusRows(worksheet);
      dataRows = dataRows.concat(statusRows);

      // Видаляємо дублікати (якщо бієць потрапив в обидва списки)
      const uniqueRows = this.removeDuplicateRows(dataRows);
      result.dataRows = uniqueRows;

      // Якщо не знайдено жодного бійця
      if (result.dataRows.length === 0) {
        result.error = 'Не знайдено бійців ні після маркера "Придані" в колонці E, ні зі статусом "Придані" в колонці H';
        return result;
      }

      // Попередження про відсутню дату
      if (!sheetInfo.date) {
        result.warnings.push('Дату листа не розпізнано - бійці знайдені тільки за статусом "Придані" в колонці H');
      }

      // Додаємо попередження якщо знайдено нестандартні ПІБ
      for (const row of dataRows) {
        const pibWords = row.pib.trim().split(/\s+/);
        if (pibWords.length !== 3) {
          result.warnings.push(`Рядок ${row.row}: нестандартний ПІБ "${row.pib}" (${pibWords.length} слов${pibWords.length === 1 ? 'о' : pibWords.length < 5 ? 'а' : ''})`);
        }
      }

      return result;

    } catch (error) {
      result.error = error instanceof Error ? error.message : 'Помилка обробки листа';
      return result;
    }
  }

  /**
   * Шукає маркер "Придані" в колонці E
   */
  private static findMarker(worksheet: ExcelJS.Worksheet): number | null {
    const column = worksheet.getColumn(this.MARKER_COLUMN);
    
    for (let rowNum = 1; rowNum <= worksheet.rowCount; rowNum++) {
      const cell = worksheet.getCell(rowNum, this.MARKER_COLUMN);
      const value = this.getCellText(cell);
      
      if (this.normalizeText(value) === this.MARKER_TEXT) {
        return rowNum;
      }
    }
    
    return null;
  }

  /**
   * Витягує дані з рядків після маркера
   */
  private static extractDataRows(worksheet: ExcelJS.Worksheet, startRow: number): ExcelRowData[] {
    const dataRows: ExcelRowData[] = [];
    let emptyCount = 0;
    
    for (let rowNum = startRow; rowNum <= worksheet.rowCount && emptyCount < this.MAX_EMPTY_ROWS; rowNum++) {
      const pibCell = worksheet.getCell(rowNum, this.PIB_COLUMN);
      const pibValue = this.getCellText(pibCell);
      
      // Якщо ПІБ порожній - збільшуємо лічильник порожніх рядків
      if (!pibValue.trim()) {
        emptyCount++;
        continue;
      }
      
      // Скидаємо лічильник порожніх рядків
      emptyCount = 0;
      
      // Читаємо решту колонок
      const pseudoCell = worksheet.getCell(rowNum, this.PSEUDO_COLUMN);
      
      const pseudoValue = this.getCellText(pseudoCell);
      
      dataRows.push({
        row: rowNum,
        pib: this.normalizePersonName(pibValue),
        pseudo: pseudoValue.trim() || null
      });
    }
    
    return dataRows;
  }

  /**
   * Отримує текстове значення з комірки
   */
  private static getCellText(cell: ExcelJS.Cell): string {
    if (cell.value === null || cell.value === undefined) {
      return '';
    }
    
    // Обробка різних типів значень
    if (typeof cell.value === 'string') {
      return cell.value;
    }
    
    if (typeof cell.value === 'number') {
      return cell.value.toString();
    }
    
    // Обробка rich text
    if (typeof cell.value === 'object' && 'richText' in cell.value) {
      return cell.value.richText
        .map((rt: any) => rt.text || '')
        .join('');
    }
    
    // Обробка формул
    if (typeof cell.value === 'object' && 'result' in cell.value) {
      return String(cell.value.result || '');
    }
    
    return String(cell.value);
  }

  /**
   * Нормалізує текст (для пошуку маркера)
   */
  private static normalizeText(text: string): string {
    return text
      .toLowerCase()
      .trim()
      .replace(/\s+/g, ' ')
      .replace(/[«»"""''`]/g, '"')
      .replace(/[—–−]/g, '-')
      .replace(/…/g, '...');
  }

  /**
   * Нормалізує ПІБ та інші особисті дані
   */
  private static normalizePersonName(text: string): string {
    return text
      .trim()
      .replace(/\s+/g, ' ')
      .replace(/[''`]/g, "'") // Уніфікація апострофів
      .replace(/[—–−]/g, "-") // Уніфікація дефісів
      .replace(/ʼ/g, "'")     // Український апостроф
      .split(' ')
      .map(word => {
        // Капіталізуємо перші літери
        if (word.length > 0) {
          return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
        }
        return word;
      })
      .join(' ');
  }

  /**
   * Створює ключ для індексації ПІБ
   */
  static createPibKey(pib: string): string {
    return this.normalizePersonName(pib)
      .toLowerCase()
      .replace(/[^а-яїієґa-z\s\-']/g, '') // Залишаємо тільки літери, пробіли, дефіси та апострофи
      .trim();
  }

  /**
   * Перевіряє чи є ПІБ валідним
   */
  static isValidPib(pib: string): boolean {
    const normalized = pib.trim();
    if (!normalized) return false;
    
    const words = normalized.split(/\s+/);
    return words.length >= 2 && words.length <= 4 && // 2-4 слова
           words.every(word => word.length >= 2); // Кожне слово мінімум 2 літери
  }

  /**
   * Шукає бійців зі статусом "Придані" в колонці H
   */
  private static extractStatusRows(worksheet: ExcelJS.Worksheet): ExcelRowData[] {
    const rows: ExcelRowData[] = [];
    const maxRows = worksheet.rowCount;

    for (let rowNumber = 1; rowNumber <= maxRows; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // Перевіряємо статус в колонці H
      const statusCell = row.getCell(this.STATUS_COLUMN);
      const statusValue = this.getCellText(statusCell);
      
      if (!statusValue || !statusValue.toLowerCase().includes(this.MARKER_TEXT)) {
        continue;
      }

      // Читаємо ПІБ з колонки E
      const pibCell = row.getCell(this.PIB_COLUMN);
      const pibValue = this.getCellText(pibCell);
      
      if (!pibValue || !this.isValidPib(pibValue)) {
        continue;
      }

      // Читаємо псевдо з колонки F
      const pseudoCell = row.getCell(this.PSEUDO_COLUMN);
      const pseudoValue = this.getCellText(pseudoCell);

      rows.push({
        row: rowNumber,
        pib: this.normalizePersonName(pibValue),
        pseudo: pseudoValue || null
      });
    }

    return rows;
  }

  /**
   * Видаляє дублікати рядків (по ПІБ + псевдо + номеру рядка)
   */
  private static removeDuplicateRows(rows: ExcelRowData[]): ExcelRowData[] {
    const seen = new Set<string>();
    const uniqueRows: ExcelRowData[] = [];

    for (const row of rows) {
      const key = `${this.createPibKey(row.pib)}_${row.pseudo || ''}_${row.row}`;
      if (!seen.has(key)) {
        seen.add(key);
        uniqueRows.push(row);
      }
    }

    return uniqueRows;
  }

  /**
   * Форматує інформацію про лист для логу
   */
  static formatSheetInfo(sheetResult: SheetProcessResult): string {
    const dateStr = sheetResult.date ? SheetDateParser.formatDate(sheetResult.date) : 'невідома дата';
    const rowsCount = sheetResult.dataRows.length;
    const markerInfo = sheetResult.markerRow ? `@E${sheetResult.markerRow}` : 'не знайдено';
    
    return `${sheetResult.sheetName} (${dateStr}, "Придані" ${markerInfo}, рядків: ${rowsCount})`;
  }
}