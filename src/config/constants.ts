/**
 * constants.ts - Константи для системи копіювання по ключу підрозділу
 * 
 * Визначає:
 * - Назви листів Excel
 * - Колонки для читання/запису
 * - Опції обробки
 * - Налаштування копіювання за замовчуванням
 */

import { CopyOptions } from '../types/MappingTypes';

/**
 * Назви листів у Excel файлах
 */
export const SHEET_NAMES = {
  ZS: 'ЗС',     // Загальна служба
  BZ: 'БЗ'      // Бойове забезпечення
} as const;

/**
 * Колонки Excel для роботи
 * НОВА СТРУКТУРА: Різні колонки для різних листів
 */
export const COLUMNS = {
  // Для листа ЗС
  ZS: {
    SUBUNIT_KEY: 'B',      // Колонка з назвою підрозділу (ключ)
    DATA_START: 'C',       // Початок діапазону даних
    DATA_END: 'H'          // Кінець діапазону даних
  },
  
  // Для листа БЗ
  BZ: {
    SUBUNIT_KEY: 'C',      // ⚠️ БЗ шукає по колонці C!
    DATA_START: 'D',       // Початок даних (D:H щоб не копіювати ключ)
    DATA_END: 'H'          // Кінець діапазону даних
  }
} as const;

/**
 * Blacklist підрозділів (не обробляти ці підрозділи)
 * Використовується для виключення підрозділів які постійно є в файлі призначення
 */
export const SUBUNIT_BLACKLIST = {
  ZS: [
    'упр',      // УПР постійно є в файлі призначення
    'п'         // П (можливо також перманентний)
  ],
  BZ: [
    // Додайте сюди підрозділи БЗ які треба ігнорувати (якщо є)
  ]
} as const;

/**
 * Налаштування процесу обробки
 */
export const PROCESSING_OPTIONS = {
  SKIP_TEMP_FILES: true,      // Пропускати тимчасові файли Excel (~$...)
  SKIP_HEADER_ROW: true,      // Пропускати рядок 1 (заголовки)
  NORMALIZE_KEYS: true,       // Нормалізувати ключі підрозділів (trim + toLowerCase)
  MAX_EMPTY_ROWS: 10,         // Максимум порожніх рядків підряд (стоп читання)
  ENABLE_3BSP: false          // 🆕 Обробка листа 3бСпП БЗ (копіювання C4:H231)
} as const;

/**
 * Налаштування копіювання за замовчуванням
 */
export const DEFAULT_COPY_OPTIONS: CopyOptions = {
  values: true,               // ✅ Копіювати ТІЛЬКИ значення комірок
  formulas: false,            // ❌ НЕ копіювати формули
  formats: false              // ❌ НЕ копіювати форматування (зберігаємо стилі файлу призначення)
};

/**
 * Фази обробки (для прогресу)
 */
export const PROCESSING_PHASES = {
  SCANNING: { percent: 10, label: 'Сканування папки з Excel файлами...' },
  OPENING: { percent: 20, label: 'Відкриття файлу призначення...' },
  INDEXING: { percent: 30, label: 'Побудова індексу підрозділів...' },
  PROCESSING: { percentStart: 30, percentEnd: 90, label: 'Обробка файлів...' },
  SAVING: { percent: 95, label: 'Збереження файлу призначення...' },
  COMPLETE: { percent: 100, label: 'Обробка завершена!' }
} as const;

/**
 * Типи логів
 */
export const LOG_TYPES = {
  INFO: 'info',
  ERROR: 'error',
  WARNING: 'warning',
  SUCCESS: 'success'
} as const;
