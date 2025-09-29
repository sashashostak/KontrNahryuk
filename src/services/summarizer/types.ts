/**
 * Типи та інтерфейси для системи зведення Excel файлів
 * Згідно з ТЗ: Інтеграція логіки зведення (БЗ / ЗС)
 */

// Основні типи
export type Mode = 'БЗ' | 'ЗС' | 'Обидва';

export type LogLevel = 'INFO' | 'WARN' | 'ERROR';

// Контракт UI → Orchestrator
export interface StartProcessPayload {
  srcFolder: string;          // вибрана папка з файлами
  dstPath: string;            // шлях до книги призначення (.xlsx)
  mode: Mode;
  dstSheetPassword?: string;  // лише для інформації/попередження
  configPath?: string;        // опційно: шлях до JSON з Rules
}

// Eventi для стріму
export type Event =
  | { type: 'log'; level: LogLevel; message: string; ts: string }
  | { type: 'progress'; current: number; total: number; note?: string }
  | { type: 'done'; summary: { foundFiles: number; copiedRows: number; warnings: number } }
  | { type: 'failed'; error: string };

// Конфігурація правил
export interface Rule {
  key: string;      // значення Підрозділу
  tokens: string[]; // токени для пошуку в назвах файлів
}

export interface Preset {
  SRC_SHEET: string;
  DST_SHEET: string;
  COL_SUBUNIT: number; // 1=A, 2=B, 3=C...
  COL_LEFT: number;
  COL_RIGHT: number;
  rules: Rule[];
}

export interface SummarizeParams extends Preset {
  SRC_FOLDER: string;
  DST_SHEET_PASSWORD?: string;
}

// Результат роботи summarize
export interface SummarizeResult {
  foundFiles: number;
  copiedRows: number;
  warnings: number;
}

// Структура блоку в Excel
export interface BlockCoords {
  r1: number; // початковий рядок (0-based)
  r2: number; // кінцевий рядок (0-based)
  c1: number; // початкова колонка (0-based)
  c2: number; // кінцева колонка (0-based)
}

// Конфігурація з JSON файлу
export interface Config {
  presets: {
    [key in Mode]?: Preset;
  };
}

// Опції для роботи з файлами
export interface FileOptions {
  includeXls?: boolean;    // включати .xls файли (потребує конвертації)
  includeXlsb?: boolean;   // включати .xlsb файли
  autoUnmerge?: boolean;   // автоматично розмерджувати комірки
}

// Інформація про файл
export interface FileInfo {
  path: string;
  name: string;
  normalized: string;  // нормалізована назва
  lastModified: Date;
  size: number;
}

// Статистика обробки
export interface ProcessStats {
  totalFiles: number;
  processedFiles: number;
  skippedFiles: number;
  totalRows: number;
  copiedRows: number;
  warnings: number;
  errors: number;
  startTime: Date;
  endTime?: Date;
}