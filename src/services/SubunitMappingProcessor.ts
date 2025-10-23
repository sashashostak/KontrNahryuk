/**
 * SubunitMappingProcessor.ts - Головний процесор копіювання по ключу підрозділу
 * 
 * Алгоритм обробки:
 * 1. Сканування папки з Excel файлами
 * 2. Відкриття файлу призначення
 * 3. Побудова індексу підрозділів (Map<назва, номер_рядка>)
 * 4. Обробка кожного файлу:
 *    - Читання листів ЗС та БЗ
 *    - Для кожного рядка:
 *      - Читання підрозділу з колонки B
 *      - Пошук у індексі
 *      - Копіювання C:H у відповідний рядок файлу призначення
 * 5. Збереження файлу призначення
 * 6. Підрахунок статистики
 */

import ExcelJS from 'exceljs';
import { 
  SubunitIndex, 
  FileProcessingResult, 
  SheetProcessingResult,
  ProcessingStats,
  CopyOptions,
  ProgressCallback,
  FileInfo
} from '../types/MappingTypes';
import { 
  SHEET_NAMES, 
  COLUMNS, 
  DEFAULT_COPY_OPTIONS,
  PROCESSING_PHASES,
  PROCESSING_OPTIONS,
  SUBUNIT_BLACKLIST
} from '../config/constants';

// 🐍 Типи для Python інтеграції (snake_case відповідно до Python)
export interface PythonProcessConfig {
  destination_file: string;
  source_files: string[];
  sheets: Array<{
    name: string;
    key_column: string;
    data_columns: string[];
    blacklist: string[];
  }>;
}

export interface PythonProcessResult {
  success: boolean;
  total_rows?: number;
  error?: string;
}
import { 
  getColumnsRange,
  isTempFile,
  isExcelFile
} from '../utils/subunitUtils';
import {
  findSingleContiguousBlock,
  findAllContiguousBlocks
} from '../utils/blockUtils';

export class SubunitMappingProcessor {
  private destinationWorkbook: ExcelJS.Workbook | null = null;
  private copyOptions: CopyOptions = DEFAULT_COPY_OPTIONS;
  
  /**
   * Основна функція обробки
   * 
   * @param inputFolder - Папка з вхідними Excel файлами
   * @param destinationFile - Шлях до файлу призначення
   * @param onProgress - Callback для відображення прогресу
   * @returns Статистика обробки
   */
  async process(
    inputFolder: string,
    destinationFile: string,
    onProgress?: ProgressCallback
  ): Promise<ProcessingStats> {
    // 🐍 Використовуємо Python замість ExcelJS (виправляє XML помилки)
    console.log(`\n� ═══ ВИКОРИСТОВУЄТЬСЯ PYTHON EXCEL PROCESSOR ═══\n`);
    return this.processWithPython(inputFolder, destinationFile, onProgress);
  }
  
  /**
   * Фаза 1: Сканування папки на наявність Excel файлів
   * 
   * @param folderPath - Шлях до папки
   * @returns Масив інформації про файли
   */
  private async scanFolder(folderPath: string): Promise<FileInfo[]> {
    try {
      // Виклик Electron API для читання папки
      if (!window.api || !window.api.readDirectory) {
        throw new Error('API для читання директорії недоступне');
      }
      
      const allFiles: FileInfo[] = await window.api.readDirectory(folderPath);
      
      // Фільтрація тільки Excel файлів (без тимчасових)
      const excelFiles = allFiles.filter(file => {
        const isExcel = isExcelFile(file.name);
        const isNotTemp = PROCESSING_OPTIONS.SKIP_TEMP_FILES 
          ? !isTempFile(file.name)
          : true;
        
        return isExcel && isNotTemp;
      });
      
      return excelFiles;
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      throw new Error(`Помилка сканування папки: ${errorMessage}`);
    }
  }
  
  /**
   * @deprecated Цей метод більше не використовується. Вся обробка виконується через Python.
   * Залишено для історичної сумісності.
   */
  // @ts-ignore - Deprecated ExcelJS method, not used in production
  private async processSheet(
    sourceSheet: ExcelJS.Worksheet,
    sheetName: string,
    index: SubunitIndex,
    sheetType: 'ZS' | 'BZ'
  ): Promise<SheetProcessingResult> {
    console.log(`\n🔄 === ОБРОБКА ЛИСТА "${sheetName}" (НОВИЙ АЛГОРИТМ) ===`);
    
    // ✅ ВИБИРАЄМО ПРАВИЛЬНІ КОЛОНКИ
    const columns = COLUMNS[sheetType];
    console.log(`🔧 Використовуємо колонки для ${sheetType}:`);
    console.log(`   • Ключ: ${columns.SUBUNIT_KEY}`);
    console.log(`   • Дані: ${columns.DATA_START}:${columns.DATA_END}`);
    
    const result: SheetProcessingResult = {
      sheetName,
      totalRows: 0,
      copiedRows: 0,
      skippedRows: 0,
      missingSubunits: [],
      errors: []
    };
    
    const destSheet = this.destinationWorkbook?.getWorksheet(sheetName);
    if (!destSheet) {
      const error = `Лист "${sheetName}" не знайдено у файлі призначення`;
      console.error(`❌ ${error}`);
      result.errors.push(error);
      return result;
    }
    
    console.log(`✅ Листи готові до обробки`);
    console.log(`📊 Вхідний: ${sourceSheet.rowCount} рядків`);
    console.log(`📊 Призначення: ${destSheet.rowCount} рядків`);
    
    const columnsToСopy = getColumnsRange(columns.DATA_START, columns.DATA_END);
    console.log(`📋 Колонки: ${columnsToСopy.join(', ')}`);
    
    // Збираємо унікальні ключі з індексу
    const uniqueKeys = Array.from(index.keys());
    
    // ✅ НОВЕ: Фільтрація blacklist
    const blacklist = SUBUNIT_BLACKLIST[sheetType] as readonly string[];
    const filteredKeys = uniqueKeys.filter(key => !blacklist.includes(key));
    
    console.log(`� Підрозділів в індексі: ${uniqueKeys.length}`);
    console.log(`🚫 Виключено (blacklist): ${uniqueKeys.length - filteredKeys.length}`);
    console.log(`✅ До обробки: ${filteredKeys.length}`);
    
    if (blacklist.length > 0) {
      console.log(`   Blacklist для ${sheetType}: ${blacklist.join(', ')}`);
    }
    
    // Обробляємо кожен підрозділ (відфільтровані)
    for (const key of filteredKeys) {
      console.log(`\n--- Підрозділ: "${key}" ---`);
      
      // 1) Знаходимо блок у джерелі
      const srcBlock = findSingleContiguousBlock(sourceSheet, columns.SUBUNIT_KEY, key);
      
      if (!srcBlock) {
        console.log(`⏭️  Пропущено: не знайдено у джерелі`);
        continue;
      }
      
      console.log(`✅ Джерело: рядки ${srcBlock.startRow}-${srcBlock.endRow} (${srcBlock.rowCount} рядків)`);
      result.totalRows += srcBlock.rowCount;
      
      // 2) Знаходимо ВСІ блоки у призначенні
      const dstBlocks = findAllContiguousBlocks(destSheet, columns.SUBUNIT_KEY, key);
      
      if (dstBlocks.length === 0) {
        console.log(`⚠️  Пропущено: не знайдено у призначенні`);
        result.skippedRows += srcBlock.rowCount;
        if (!result.missingSubunits.includes(key)) {
          result.missingSubunits.push(key);
        }
        continue;
      }
      
      console.log(`✅ Призначення: ${dstBlocks.length} блок(ів)`);
      dstBlocks.forEach((block, idx) => {
        console.log(`   Блок ${idx + 1}: рядки ${block.startRow}-${block.endRow} (${block.rowCount} рядків)`);
      });
      
      // 3) Копіюємо дані по частинам
      let srcRowPtr = srcBlock.startRow;
      
      for (let d = 0; d < dstBlocks.length; d++) {
        const dstBlock = dstBlocks[d];
        
        // Скільки рядків залишилось у джерелі
        const rowsLeftSrc = (srcBlock.endRow - srcRowPtr + 1);
        if (rowsLeftSrc <= 0) {
          console.log(`⚠️  Дані джерела закінчились`);
          break;
        }
        
        // Скільки рядків копіювати
        const rowsToCopy = Math.min(rowsLeftSrc, dstBlock.rowCount);
        
        console.log(`\n� Копіювання у блок ${d + 1}:`);
        console.log(`   Джерело: рядки ${srcRowPtr}-${srcRowPtr + rowsToCopy - 1}`);
        console.log(`   Призначення: рядки ${dstBlock.startRow}-${dstBlock.startRow + rowsToCopy - 1}`);
        console.log(`   Кількість: ${rowsToCopy} рядків × ${columnsToСopy.length} колонок`);
        
        try {
          // КЛЮЧОВИЙ МОМЕНТ: Копіюємо БЛОК, а не комірку за коміркою!
          // @ts-ignore - Deprecated method, not used in production
          this.copyBlockValues(
            sourceSheet,
            destSheet,
            srcRowPtr,
            dstBlock.startRow,
            rowsToCopy,
            columnsToСopy
          );
          
          result.copiedRows += rowsToCopy;
          console.log(`   ✅ Скопійовано успішно!`);
          
          // Очищуємо "хвіст" якщо блок призначення більший
          if (rowsToCopy < dstBlock.rowCount) {
            const tailRows = dstBlock.rowCount - rowsToCopy;
            const tailStart = dstBlock.startRow + rowsToCopy;
            
            // @ts-ignore - Deprecated method, not used in production
            this.clearBlockValues(
              destSheet,
              tailStart,
              tailRows,
              columnsToСopy
            );
            
            console.log(`   🧹 Очищено хвіст: ${tailRows} рядків`);
          }
          
          srcRowPtr += rowsToCopy;
          
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          const errorMsg = `Помилка копіювання: ${errorMessage}`;
          result.errors.push(errorMsg);
          console.error(`   ❌ ПОМИЛКА:`, error);
        }
      }
    }
    
    console.log(`\n📊 ПІДСУМОК "${sheetName}":`);
    console.log(`   Всього рядків: ${result.totalRows}`);
    console.log(`   Скопійовано: ${result.copiedRows}`);
    console.log(`   Пропущено: ${result.skippedRows}`);
    console.log(`🔄 === КІНЕЦЬ "${sheetName}" ===\n`);
    
    return result;
  }
  
  /**
   * Підрахунок фінальної статистики
   * 
   * @param results - Результати обробки всіх файлів
   * @param startTime - Час початку обробки (ms)
   * @returns Загальна статистика
   */
  private calculateStats(
    results: FileProcessingResult[], 
    startTime: number
  ): ProcessingStats {
    const stats: ProcessingStats = {
      totalFiles: results.length,
      processedFiles: 0,
      failedFiles: 0,
      totalCopiedRowsZS: 0,
      totalCopiedRowsBZ: 0,
      totalSkippedRowsZS: 0,
      totalSkippedRowsBZ: 0,
      allMissingSubunits: [],
      processingTime: (Date.now() - startTime) / 1000
    };
    
    const missingSet = new Set<string>();
    
    results.forEach(result => {
      if (result.processed) {
        stats.processedFiles++;
        
        // Статистика ЗС
        if (result.zsSheet) {
          stats.totalCopiedRowsZS += result.zsSheet.copiedRows;
          stats.totalSkippedRowsZS += result.zsSheet.skippedRows;
          result.zsSheet.missingSubunits.forEach(s => missingSet.add(s));
        }
        
        // Статистика БЗ
        if (result.bzSheet) {
          stats.totalCopiedRowsBZ += result.bzSheet.copiedRows;
          stats.totalSkippedRowsBZ += result.bzSheet.skippedRows;
          result.bzSheet.missingSubunits.forEach(s => missingSet.add(s));
        }
      } else {
        stats.failedFiles++;
      }
    });
    
    stats.allMissingSubunits = Array.from(missingSet).sort();
    
    return stats;
  }
  
  /**
   * Встановлення налаштувань копіювання
   * 
   * @param options - Налаштування копіювання
   */
  setCopyOptions(options: Partial<CopyOptions>): void {
    this.copyOptions = {
      ...this.copyOptions,
      ...options
    };
  }
  
  /**
   * Отримання поточних налаштувань копіювання
   */
  getCopyOptions(): CopyOptions {
    return { ...this.copyOptions };
  }
  
  /**
   * Допоміжний метод: отримати список літер колонок
   */
  private getColumnsList(start: string, end: string): string[] {
    const columns: string[] = [];
    const startCode = start.charCodeAt(0);
    const endCode = end.charCodeAt(0);
    
    for (let code = startCode; code <= endCode; code++) {
      columns.push(String.fromCharCode(code));
    }
    
    return columns;
  }
  
  /**
   * НОВИЙ МЕТОД: Обробка через Python
   */
  async processWithPython(
    inputFolder: string,
    destinationFile: string,
    onProgress?: ProgressCallback
  ): Promise<ProcessingStats> {
    const startTime = Date.now();
    
    try {
      // Фаза 1: Сканування папки
      onProgress?.(PROCESSING_PHASES.SCANNING.percent, PROCESSING_PHASES.SCANNING.label);
      const files = await this.scanFolder(inputFolder);
      
      console.log(`\n📁 ═══ СКАНУВАННЯ ПАПКИ ═══`);
      console.log(`📍 Шлях: ${inputFolder}`);
      console.log(`📊 Знайдено Excel файлів: ${files.length}`);
      
      if (files.length === 0) {
        console.log(`⚠️ У папці немає файлів Excel`);
        throw new Error('У папці не знайдено Excel файлів');
      }
      
      console.log(`📋 Список файлів для обробки:`);
      files.forEach((file, index) => {
        console.log(`   ${index + 1}. ${file.name}`);
      });
      
      // 🐍 ВИКОРИСТОВУЄМО PYTHON
      console.log(`\n🐍 ═══ ВИКОРИСТОВУЄМО PYTHON EXCEL PROCESSOR ===\n`);
      
      // Отримуємо налаштування 3БСП
      const enable3BSP = await window.api?.getSetting?.('enable3BSP', false) || false;
      console.log(`🔧 Налаштування enable3BSP: ${enable3BSP}`);
      
      const enableSanitizer = await window.api?.getSetting?.('excel.enableSanitizer', false) || false;
      console.log(`🔧 Налаштування enableSanitizer: ${enableSanitizer}`);
      
      const enableMismatches = await window.api?.getSetting?.('excel.showMismatches', false) || false;
      console.log(`🔧 Налаштування enableMismatches: ${enableMismatches}`);
      
      const enableSliceCheck = await window.api?.getSetting?.('excel.enableSliceCheck', false) || false;
      console.log(`🔧 Налаштування enableSliceCheck: ${enableSliceCheck}`);
      
      const enableDuplicates = await window.api?.getSetting?.('excel.enableDuplicates', false) || false;
      console.log(`🔧 Налаштування enableDuplicates: ${enableDuplicates}`);
      
      // Формуємо конфігурацію для Python (snake_case для Python)
      const config: any = {
        destination_file: destinationFile,
        source_files: files.map(f => f.path),
        enable_3bsp: enable3BSP,  // 🆕 Передаємо налаштування 3БСП
        enable_sanitizer: enableSanitizer,  // 🧹 Передаємо налаштування санітизації
        enable_mismatches: enableMismatches,  // 🔍 Передаємо налаштування перевірки невідповідностей
        enable_slice_check: enableSliceCheck,  // 🔪 Передаємо налаштування перевірки «зрізів»
        enable_duplicates: enableDuplicates,  // 🔁 Передаємо налаштування перевірки дублікатів
        sheets: [
          {
            name: SHEET_NAMES.ZS,
            key_column: COLUMNS.ZS.SUBUNIT_KEY,
            data_columns: this.getColumnsList(COLUMNS.ZS.DATA_START, COLUMNS.ZS.DATA_END),
            blacklist: Array.from(SUBUNIT_BLACKLIST.ZS)
          },
          {
            name: SHEET_NAMES.BZ,
            key_column: COLUMNS.BZ.SUBUNIT_KEY,
            data_columns: this.getColumnsList(COLUMNS.BZ.DATA_START, COLUMNS.BZ.DATA_END),
            blacklist: Array.from(SUBUNIT_BLACKLIST.BZ)
          }
        ]
      };
      
      onProgress?.(PROCESSING_PHASES.PROCESSING.percentStart, 'Обробка через Python...');
      
      // 🐍 Викликаємо Python через IPC
      if (!window.api || !window.api.invoke) {
        throw new Error('❌ Electron API не доступний. Переконайтесь що preload.ts підключено правильно.');
      }
      
      const pythonResult = await window.api.invoke('python:process-excel', config) as PythonProcessResult;
      
      if (!pythonResult.success) {
        throw new Error(pythonResult.error || 'Python processing failed');
      }
      
      console.log(`✅ Python обробив ${pythonResult.total_rows} рядків`);
      
      // Формуємо результати для статистики
      const results: FileProcessingResult[] = files.map(file => ({
        fileName: file.name,
        filePath: file.path,
        processed: true,
        zsSheet: { sheetName: SHEET_NAMES.ZS, totalRows: 0, copiedRows: pythonResult.total_rows || 0, skippedRows: 0, missingSubunits: [], errors: [] },
        bzSheet: null
      }));
      
      // Підрахунок статистики
      const stats = this.calculateStats(results, startTime);
      
      onProgress?.(PROCESSING_PHASES.COMPLETE.percent, PROCESSING_PHASES.COMPLETE.label);
      
      console.log(`\n✅ ═══ ОБРОБКА ЗАВЕРШЕНА ===\n`);
      
      return stats;
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      throw new Error(`Помилка обробки: ${errorMessage}`);
    }
  }
}
