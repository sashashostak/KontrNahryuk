/**
 * MappingTypes.ts - Типи даних для системи копіювання по ключу підрозділу
 * 
 * Описує інтерфейси для:
 * - Індексації підрозділів
 * - Результатів обробки листів та файлів
 * - Загальної статистики
 * - Налаштувань копіювання
 */

/**
 * Індекс підрозділів: Map<назва_підрозділу, номер_рядка>
 * 
 * Приклад:
 * {
 *   "1рспп" => 5,
 *   "2рспп" => 12,
 *   "врсп" => 25
 * }
 */
export type SubunitIndex = Map<string, number>;

/**
 * Результат обробки одного листа
 */
export interface SheetProcessingResult {
  sheetName: string;              // "ЗС" або "БЗ"
  totalRows: number;              // Всього рядків у вхідному файлі
  copiedRows: number;             // Скільки рядків скопійовано
  skippedRows: number;            // Скільки рядків пропущено
  missingSubunits: string[];      // Підрозділи, які не знайдено у файлі призначення
  errors: string[];               // Помилки обробки
}

/**
 * Результат обробки одного файлу
 */
export interface FileProcessingResult {
  fileName: string;               // Ім'я файлу (без шляху)
  filePath: string;               // Повний шлях до файлу
  processed: boolean;             // Чи успішно оброблено
  zsSheet: SheetProcessingResult | null;  // Результат обробки листа ЗС
  bzSheet: SheetProcessingResult | null;  // Результат обробки листа БЗ
  error?: string;                 // Критична помилка (якщо processed = false)
}

/**
 * Загальна статистика обробки
 */
export interface ProcessingStats {
  totalFiles: number;             // Всього файлів знайдено
  processedFiles: number;         // Успішно оброблено
  failedFiles: number;            // Помилок обробки
  totalCopiedRowsZS: number;      // Всього скопійовано рядків ЗС
  totalCopiedRowsBZ: number;      // Всього скопійовано рядків БЗ
  totalSkippedRowsZS: number;     // Всього пропущено рядків ЗС
  totalSkippedRowsBZ: number;     // Всього пропущено рядків БЗ
  allMissingSubunits: string[];   // Унікальні відсутні підрозділи
  processingTime: number;         // Час обробки (секунди)
}

/**
 * Налаштування копіювання даних
 */
export interface CopyOptions {
  values: boolean;        // Копіювати значення комірок
  formulas: boolean;      // Копіювати формули
  formats: boolean;       // Копіювати форматування (стилі, кольори, тощо)
}

/**
 * Callback для відображення прогресу
 * 
 * @param percent - Відсоток виконання (0-100)
 * @param message - Повідомлення про поточну операцію
 */
export type ProgressCallback = (percent: number, message: string) => void;

/**
 * Інформація про файл у директорії
 */
export interface FileInfo {
  name: string;           // Ім'я файлу
  path: string;           // Повний шлях
}
