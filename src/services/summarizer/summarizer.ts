/**
 * Основний сервіс зведення Excel файлів
 * Реалізує бізнес-логіку обробки блоків даних за правилами
 */

import { 
  SummarizeParams, 
  SummarizeResult, 
  Event, 
  LogLevel, 
  BlockCoords,
  FileOptions,
  ProcessStats
} from './types';
import {
  listExcelFiles,
  findLatestByTokens,
  readWorksheet,
  writeWorksheet,
  findFirstContiguousBlock,
  findAllContiguousBlocks,
  copySlice,
  col1to0,
  colNumberToLetter
} from './helpers';

/**
 * Головна функція зведення даних згідно з параметрами
 */
export async function summarize(
  params: SummarizeParams,
  dstPath: string,
  emit: (event: Event) => void
): Promise<SummarizeResult> {
  
  const stats: ProcessStats = {
    totalFiles: 0,
    processedFiles: 0,
    skippedFiles: 0,
    totalRows: 0,
    copiedRows: 0,
    warnings: 0,
    errors: 0,
    startTime: new Date()
  };

  try {
    emitLog(emit, 'INFO', `Початок обробки режиму ${params.DST_SHEET}`);
    emitLog(emit, 'INFO', `Джерело: ${params.SRC_FOLDER}`);
    emitLog(emit, 'INFO', `Призначення: ${dstPath}`);
    
    // Валідація параметрів
    if (!params.SRC_FOLDER || !dstPath) {
      throw new Error('Не вказано папку джерела або файл призначення');
    }

    // Отримуємо список Excel файлів
    const fileOptions: FileOptions = {
      includeXls: false,    // поки що не підтримуємо .xls
      includeXlsb: false    // поки що не підтримуємо .xlsb
    };
    
    const allFiles = await listExcelFiles(params.SRC_FOLDER, fileOptions);
    stats.totalFiles = allFiles.length;
    
    emitLog(emit, 'INFO', `Знайдено ${allFiles.length} Excel файлів у папці`);

    if (allFiles.length === 0) {
      emitLog(emit, 'WARN', 'Не знайдено жодного Excel файлу для обробки');
      return { foundFiles: 0, copiedRows: 0, warnings: 1 };
    }

    // Обробляємо кожне правило
    let totalCopiedRows = 0;
    const processedRules: string[] = [];

    for (let i = 0; i < params.rules.length; i++) {
      const rule = params.rules[i];
      
      // Відправляємо прогрес
      emit({
        type: 'progress',
        current: i + 1,
        total: params.rules.length,
        note: `Обробка правила: ${rule.key}`
      });

      try {
        // Знаходимо відповідний файл за токенами
        const targetFile = findLatestByTokens(allFiles, rule.tokens);
        
        if (!targetFile) {
          emitLog(emit, 'WARN', `Не знайдено файл для правила "${rule.key}" з токенами: ${rule.tokens.join(', ')}`);
          stats.warnings++;
          continue;
        }

        emitLog(emit, 'INFO', `Правило "${rule.key}": використовую файл ${targetFile.name}`);
        
        // Читаємо джерельний аркуш
        const srcData = await readWorksheet(targetFile.path, params.SRC_SHEET);
        
        if (!srcData || srcData.length === 0) {
          emitLog(emit, 'WARN', `Аркуш "${params.SRC_SHEET}" порожній у файлі ${targetFile.name}`);
          stats.warnings++;
          continue;
        }

        // Знаходимо блок з ключем у джерелі
        const keyColIdx = col1to0(params.COL_SUBUNIT);
        const leftColIdx = col1to0(params.COL_LEFT);
        const rightColIdx = col1to0(params.COL_RIGHT);
        
        const srcBlock = findFirstContiguousBlock(
          srcData, 
          keyColIdx, 
          rule.key, 
          leftColIdx, 
          rightColIdx
        );

        if (!srcBlock) {
          emitLog(emit, 'WARN', `Не знайдено блок з ключем "${rule.key}" в колонці ${colNumberToLetter(params.COL_SUBUNIT)} файлу ${targetFile.name}`);
          stats.warnings++;
          continue;
        }

        const blockSize = srcBlock.r2 - srcBlock.r1 + 1;
        emitLog(emit, 'INFO', `Знайдено блок "${rule.key}": рядки ${srcBlock.r1 + 1}-${srcBlock.r2 + 1}, ${blockSize} записів`);

        // Обробляємо файл призначення
        const copiedRows = await processDestinationFile(
          dstPath,
          params.DST_SHEET,
          rule.key,
          srcData,
          srcBlock,
          params,
          emit
        );

        totalCopiedRows += copiedRows;
        stats.copiedRows += copiedRows;
        stats.processedFiles++;
        processedRules.push(rule.key);
        
        emitLog(emit, 'INFO', `Правило "${rule.key}": скопійовано ${copiedRows} рядків`);
        
      } catch (error) {
        emitLog(emit, 'ERROR', `Помилка обробки правила "${rule.key}": ${error}`);
        stats.errors++;
      }
    }

    stats.endTime = new Date();
    const duration = stats.endTime.getTime() - stats.startTime.getTime();
    
    emitLog(emit, 'INFO', `Завершено обробку режиму ${params.DST_SHEET}`);
    emitLog(emit, 'INFO', `Оброблено правил: ${processedRules.length}/${params.rules.length}`);
    emitLog(emit, 'INFO', `Скопійовано рядків: ${totalCopiedRows}`);
    emitLog(emit, 'INFO', `Попереджень: ${stats.warnings}, помилок: ${stats.errors}`);
    emitLog(emit, 'INFO', `Час виконання: ${Math.round(duration / 1000)} секунд`);

    return {
      foundFiles: stats.processedFiles,
      copiedRows: totalCopiedRows,
      warnings: stats.warnings
    };

  } catch (error) {
    emitLog(emit, 'ERROR', `Критична помилка зведення: ${error}`);
    throw error;
  }
}

/**
 * Обробляє файл призначення: читає, знаходить блоки, копіює дані
 */
async function processDestinationFile(
  dstPath: string,
  dstSheetName: string,
  key: string,
  srcData: string[][],
  srcBlock: BlockCoords,
  params: SummarizeParams,
  emit: (event: Event) => void
): Promise<number> {
  
  let totalCopiedRows = 0;

  try {
    // Обробляємо файл призначення
    await writeWorksheet(dstPath, dstSheetName, (dstData: string[][]) => {
      
      // Знаходимо всі блоки з ключем у призначенні
      const keyColIdx = col1to0(params.COL_SUBUNIT);
      const leftColIdx = col1to0(params.COL_LEFT);
      const rightColIdx = col1to0(params.COL_RIGHT);
      
      const dstBlocks = findAllContiguousBlocks(
        dstData, 
        keyColIdx, 
        key, 
        leftColIdx, 
        rightColIdx
      );

      if (dstBlocks.length === 0) {
        emitLog(emit, 'WARN', `Не знайдено блоків з ключем "${key}" у файлі призначення`);
        return;
      }

      emitLog(emit, 'INFO', `Знайдено ${dstBlocks.length} блок(ів) з ключем "${key}" у файлі призначення`);

      // Копіюємо дані в кожен знайдений блок
      dstBlocks.forEach((dstBlock, index) => {
        const copiedRows = copySlice(srcData, srcBlock, dstData, dstBlock);
        totalCopiedRows += copiedRows;
        
        emitLog(emit, 'INFO', 
          `Блок ${index + 1}/${dstBlocks.length} (рядки ${dstBlock.r1 + 1}-${dstBlock.r2 + 1}): ` +
          `скопійовано ${copiedRows} рядків`
        );
      });
    });

  } catch (error) {
    throw new Error(`Помилка обробки файлу призначення: ${error}`);
  }

  return totalCopiedRows;
}

/**
 * Допоміжна функція для відправлення лог-подій
 */
function emitLog(emit: (event: Event) => void, level: LogLevel, message: string): void {
  const timestamp = new Date().toISOString();
  emit({
    type: 'log',
    level,
    message,
    ts: timestamp
  });
}

/**
 * Валідує параметри зведення
 */
export function validateSummarizeParams(params: SummarizeParams): string[] {
  const errors: string[] = [];

  if (!params.SRC_FOLDER) {
    errors.push('Не вказано папку джерела (SRC_FOLDER)');
  }

  if (!params.SRC_SHEET) {
    errors.push('Не вказано аркуш джерела (SRC_SHEET)');
  }

  if (!params.DST_SHEET) {
    errors.push('Не вказано аркуш призначення (DST_SHEET)');
  }

  if (params.COL_SUBUNIT < 1) {
    errors.push('Невірна колонка підрозділу (COL_SUBUNIT)');
  }

  if (params.COL_LEFT < 1) {
    errors.push('Невірна ліва колонка (COL_LEFT)');
  }

  if (params.COL_RIGHT < 1) {
    errors.push('Невірна права колонка (COL_RIGHT)');
  }

  if (params.COL_LEFT > params.COL_RIGHT) {
    errors.push('Ліва колонка не може бути більше правої');
  }

  if (!params.rules || params.rules.length === 0) {
    errors.push('Не вказано правила обробки');
  } else {
    params.rules.forEach((rule, index) => {
      if (!rule.key) {
        errors.push(`Правило ${index + 1}: не вказано ключ`);
      }
      if (!rule.tokens || rule.tokens.length === 0) {
        errors.push(`Правило ${index + 1}: не вказано токени`);
      }
    });
  }

  return errors;
}