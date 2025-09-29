/**
 * Orchestrator - головний координатор процесу зведення Excel файлів
 * Керує валідацією, завантаженням конфігурації та запуском обробки
 */

import { 
  StartProcessPayload, 
  Event, 
  SummarizeParams,
  Mode
} from './types';
import { loadConfig, getPresetsForMode, validatePreset } from './config';
import { summarize, validateSummarizeParams } from './summarizer';

/**
 * Основна функція запуску процесу зведення
 * @param payload - параметри від UI
 * @param emit - функція для відправлення подій
 */
export async function startProcess(
  payload: StartProcessPayload,
  emit: (event: Event) => void
): Promise<void> {
  try {
    // Валідація вхідних параметрів
    const validationErrors = validateInputs(payload);
    if (validationErrors.length > 0) {
      const errorMessage = `Помилки валідації: ${validationErrors.join('; ')}`;
      emit({ type: 'failed', error: errorMessage });
      return;
    }

    emit({ 
      type: 'log', 
      level: 'INFO', 
      message: `🚀 Початок процесу зведення. Режим: ${payload.mode}`, 
      ts: new Date().toISOString() 
    });

    // Завантажуємо конфігурацію
    const config = await loadConfig(payload.configPath);
    
    // Отримуємо пресети для режиму
    const runModes = getPresetsForMode(config, payload.mode);
    
    if (runModes.length === 0) {
      emit({ type: 'failed', error: `Не знайдено конфігурації для режиму "${payload.mode}"` });
      return;
    }

    // Валідуємо пресети
    for (const preset of runModes) {
      if (!validatePreset(preset)) {
        emit({ type: 'failed', error: `Невірна конфігурація пресету для режиму "${preset.DST_SHEET}"` });
        return;
      }
    }

    emit({ 
      type: 'log', 
      level: 'INFO', 
      message: `📋 Завантажено ${runModes.length} режим(ів) для обробки`, 
      ts: new Date().toISOString() 
    });

    // Попередження про захист аркуша якщо вказано пароль
    if (payload.dstSheetPassword) {
      emit({ 
        type: 'log', 
        level: 'WARN', 
        message: `⚠️ Вказано пароль для аркуша, але зняття захисту не підтримується в поточній версії. Переконайтеся, що аркуш не захищений.`, 
        ts: new Date().toISOString() 
      });
    }

    // Статистика по всіх режимах
    let foundFilesTotal = 0;
    let copiedRowsTotal = 0;
    let warningsTotal = 0;
    
    // Обробляємо кожен режим
    for (let i = 0; i < runModes.length; i++) {
      const preset = runModes[i];
      
      emit({ 
        type: 'log', 
        level: 'INFO', 
        message: `📊 Режим ${i + 1}/${runModes.length}: ${preset.DST_SHEET}`, 
        ts: new Date().toISOString() 
      });

      // Створюємо параметри зведення
      const summarizeParams: SummarizeParams = {
        ...preset,
        SRC_FOLDER: payload.srcFolder,
        DST_SHEET_PASSWORD: payload.dstSheetPassword
      };

      // Додаткова валідація параметрів
      const paramErrors = validateSummarizeParams(summarizeParams);
      if (paramErrors.length > 0) {
        const errorMessage = `Помилки параметрів режиму "${preset.DST_SHEET}": ${paramErrors.join('; ')}`;
        emit({ type: 'log', level: 'ERROR', message: errorMessage, ts: new Date().toISOString() });
        emit({ type: 'failed', error: errorMessage });
        return;
      }

      try {
        // Запускаємо зведення для поточного режиму
        const result = await summarize(summarizeParams, payload.dstPath, emit);
        
        foundFilesTotal += result.foundFiles;
        copiedRowsTotal += result.copiedRows;
        warningsTotal += result.warnings;

        emit({ 
          type: 'log', 
          level: 'INFO', 
          message: `✅ Режим "${preset.DST_SHEET}" завершено. Файлів: ${result.foundFiles}, рядків: ${result.copiedRows}, попереджень: ${result.warnings}`, 
          ts: new Date().toISOString() 
        });

      } catch (error) {
        const errorMessage = `Помилка в режимі "${preset.DST_SHEET}": ${error}`;
        emit({ type: 'log', level: 'ERROR', message: errorMessage, ts: new Date().toISOString() });
        emit({ type: 'failed', error: errorMessage });
        return;
      }
    }

    // Підсумок
    emit({ 
      type: 'log', 
      level: 'INFO', 
      message: `🎉 Всі режими завершено успішно!`, 
      ts: new Date().toISOString() 
    });
    
    emit({ 
      type: 'log', 
      level: 'INFO', 
      message: `📈 Підсумок: файлів - ${foundFilesTotal}, рядків - ${copiedRowsTotal}, попереджень - ${warningsTotal}`, 
      ts: new Date().toISOString() 
    });

    // Відправляємо подію завершення
    emit({
      type: 'done',
      summary: {
        foundFiles: foundFilesTotal,
        copiedRows: copiedRowsTotal,
        warnings: warningsTotal
      }
    });

  } catch (error) {
    const errorMessage = `Критична помилка orchestrator: ${error}`;
    emit({ type: 'log', level: 'ERROR', message: errorMessage, ts: new Date().toISOString() });
    emit({ type: 'failed', error: errorMessage });
  }
}

/**
 * Валідує вхідні параметри від UI
 */
function validateInputs(payload: StartProcessPayload): string[] {
  const errors: string[] = [];

  // Перевірка папки джерела
  if (!payload.srcFolder || payload.srcFolder.trim() === '') {
    errors.push('Не вказано папку джерела');
  }

  // Перевірка файлу призначення
  if (!payload.dstPath || payload.dstPath.trim() === '') {
    errors.push('Не вказано файл призначення');
  } else if (!payload.dstPath.toLowerCase().endsWith('.xlsx')) {
    errors.push('Файл призначення повинен мати розширення .xlsx');
  }

  // Перевірка режиму
  const validModes: Mode[] = ['БЗ', 'ЗС', 'Обидва'];
  if (!validModes.includes(payload.mode)) {
    errors.push(`Невірний режим "${payload.mode}". Дозволені: ${validModes.join(', ')}`);
  }

  // Перевірка шляху до конфігурації (якщо вказано)
  if (payload.configPath && payload.configPath.trim() !== '') {
    if (!payload.configPath.toLowerCase().endsWith('.json')) {
      errors.push('Файл конфігурації повинен мати розширення .json');
    }
  }

  return errors;
}

/**
 * Створює mock stream подій для тестування
 * В реальному застосунку це буде замінено на Observable або EventEmitter
 */
export class ProcessEventEmitter {
  private listeners: ((event: Event) => void)[] = [];

  subscribe(listener: (event: Event) => void): void {
    this.listeners.push(listener);
  }

  unsubscribe(listener: (event: Event) => void): void {
    const index = this.listeners.indexOf(listener);
    if (index >= 0) {
      this.listeners.splice(index, 1);
    }
  }

  emit(event: Event): void {
    this.listeners.forEach(listener => {
      try {
        listener(event);
      } catch (error) {
        console.error('Помилка в listener:', error);
      }
    });
  }

  /**
   * Запускає процес з автоматичною емісією подій
   */
  async startProcess(payload: StartProcessPayload): Promise<void> {
    await startProcess(payload, (event) => this.emit(event));
  }
}

/**
 * Зручна функція для одноразового запуску процесу
 */
export async function runSummarization(
  payload: StartProcessPayload,
  onEvent?: (event: Event) => void
): Promise<{ foundFiles: number; copiedRows: number; warnings: number } | null> {
  
  return new Promise((resolve, reject) => {
    const emitter = new ProcessEventEmitter();
    
    let result: { foundFiles: number; copiedRows: number; warnings: number } | null = null;

    emitter.subscribe((event) => {
      // Передаємо події зовнішньому обробнику
      if (onEvent) {
        onEvent(event);
      }

      // Обробляємо фінальні події
      switch (event.type) {
        case 'done':
          result = event.summary;
          resolve(result);
          break;
        case 'failed':
          reject(new Error(event.error));
          break;
      }
    });

    // Запускаємо процес
    emitter.startProcess(payload).catch(reject);
  });
}