/**
 * Константи проекту
 * FIXED: Винесено всі магічні числа та хардкод з main.ts
 */

// FIXED: Версія додатку (була захардкоджена в main.ts як '1.2.7')
export const APP_VERSION = '1.3.0';

// FIXED: Ключі налаштувань (використовувались як magic strings)
export const SETTINGS_KEYS = {
  IS_2BSP: 'is2BSP',
  IS_ORDER: 'isOrder',
  AUTO_OPEN: 'autoOpen',
  SHOW_TOKENS: 'showTokens',
  MAX_FILES: 'maxFiles',
  BATCH_NOTIFICATIONS: 'batchNotifications',
  BATCH_LOGS: 'batchLogs',
  POWER_SAVE: 'powerSave',
  STARTUP_CHECK: 'startupCheck',
  MINIMIZE_TRAY: 'minimizeTray',
  THEME: 'theme',
  EXCEL_DESTINATION_FILE: 'excelDestinationFile',
  EXCEL_REMEMBER_DESTINATION: 'excelRememberDestination',
  EXCEL_SLICE_CHECK: 'excelSliceCheck',
  EXCEL_MISMATCHES: 'excelMismatches',
  EXCEL_SANITIZER: 'excelSanitizer',
  EXCEL_SUMMARIZATION_MODE: 'excelSummarizationMode',
  LAST_UPDATE_CHECK: 'lastUpdateCheck'
} as const;

// FIXED: Налаштування за замовчуванням
export const DEFAULT_SETTINGS = {
  [SETTINGS_KEYS.IS_2BSP]: true,
  [SETTINGS_KEYS.IS_ORDER]: false,
  [SETTINGS_KEYS.AUTO_OPEN]: true,
  [SETTINGS_KEYS.SHOW_TOKENS]: true,
  [SETTINGS_KEYS.MAX_FILES]: 100,
  [SETTINGS_KEYS.BATCH_NOTIFICATIONS]: true,
  [SETTINGS_KEYS.BATCH_LOGS]: true,
  [SETTINGS_KEYS.POWER_SAVE]: true,
  [SETTINGS_KEYS.STARTUP_CHECK]: true,
  [SETTINGS_KEYS.MINIMIZE_TRAY]: false,
  [SETTINGS_KEYS.THEME]: 'system',
  [SETTINGS_KEYS.EXCEL_SLICE_CHECK]: true,
  [SETTINGS_KEYS.EXCEL_MISMATCHES]: true,
  [SETTINGS_KEYS.EXCEL_SANITIZER]: true,
  [SETTINGS_KEYS.EXCEL_SUMMARIZATION_MODE]: 'Обидва',
  [SETTINGS_KEYS.EXCEL_REMEMBER_DESTINATION]: true
} as const;

// FIXED: ID елементів DOM (використовувались як magic strings)
export const DOM_IDS = {
  // Навігація та головні секції
  LOG_BODY: 'log-body',
  NOTES_LIST: 'notes-list',
  
  // Кнопки
  BTN_NOTIFY: 'btn-notify',
  BTN_DOCS: 'btn-docs',
  BTN_PROCESS_ORDER: 'btn-process-order',
  BTN_ADD_NOTE: 'btn-add-note',
  BTN_CHECK_UPDATE: 'btn-check-update',
  BTN_CHECK_UPDATES: 'btn-check-updates',
  BTN_CHOOSE_BATCH_FOLDER: 'btn-choose-batch-folder',
  BTN_CHOOSE_BATCH_OUTPUT: 'btn-choose-batch-output',
  BTN_START_BATCH: 'btn-start-batch',
  BTN_CANCEL_BATCH: 'btn-cancel-batch',
  BTN_CLEAR_BATCH_LOG: 'btn-clear-batch-log',
  BTN_SAVE_BATCH_LOG: 'btn-save-batch-log',
  
  // Input поля
  WORD_FILE: 'word-file',
  WORD_FILES: 'word-files',
  EXCEL_DB: 'excel-db',
  RESULT_PATH: 'result-path',
  NOTE_INPUT: 'note-input',
  THEME_SELECT: 'theme-select',
  
  // Чекбокси
  T_2BSP: 't-2bsp',
  T_ORDER: 't-order',
  T_AUTOOPEN: 't-autopen',
  
  // Секції
  SINGLE_FILE_INPUT: 'single-file-input',
  MULTIPLE_FILES_INPUT: 'multiple-files-input',
  FOLDER_INPUT: 'folder-input',
  EXCEL_SECTION: 'excel-section',
  
  // Оновлення
  CURRENT_VERSION: 'current-version',
  UPDATE_STATUS: 'update-status',
  UPDATE_AVAILABLE: 'update-available',
  NEW_VERSION: 'new-version',
  UPDATE_DATE: 'update-date',
  RELEASE_NOTES: 'release-notes',
  LICENSE_KEY_INPUT: 'license-key-input',
  LICENSE_STATUS: 'license-status',
  LICENSE_INPUT_SECTION: 'license-input-section',
  LICENSE_GATE: 'license-gate',
  GATE_LICENSE_INPUT: 'gate-license-input',
  GATE_LICENSE_STATUS: 'gate-license-status',
  
  // Прогрес оновлення
  UPDATE_PROGRESS: 'update-progress',
  PROGRESS_TEXT: 'progress-text',
  PROGRESS_FILL: 'progress-fill',
  PROGRESS_PERCENT: 'progress-percent',
  PROGRESS_SPEED: 'progress-speed',
  PROGRESS_SIZE: 'progress-size',
  BTN_RESTART_AFTER_UPDATE: 'btn-restart-after-update',
  UPDATE_ERROR: 'update-error',
  ERROR_MESSAGE: 'error-message',
  
  // Batch обробка
  BATCH_INPUT_FOLDER: 'batch-input-folder',
  BATCH_OUTPUT_FILE: 'batch-output-file',
  BATCH_LOG_BODY: 'batch-log-body',
  BATCH_PROGRESS: 'batch-progress',
  BATCH_PROGRESS_FILL: 'batch-progress-fill',
  BATCH_PROGRESS_PERCENT: 'batch-progress-percent',
  BATCH_PROGRESS_STATUS: 'batch-progress-status',
  BATCH_PROGRESS_DETAIL: 'batch-progress-detail',
  
  // Excel обробка
  EXCEL_INPUT_FOLDER: 'excel-input-folder',
  EXCEL_DESTINATION_FILE: 'excel-destination-file',
  EXCEL_SELECT_FOLDER: 'excel-select-folder',
  EXCEL_SELECT_DESTINATION: 'excel-select-destination',
  EXCEL_START_PROCESSING: 'excel-start-processing',
  EXCEL_STOP_PROCESSING: 'excel-stop-processing',
  EXCEL_CLEAR_LOGS: 'excel-clear-logs',
  EXCEL_SAVE_LOGS: 'excel-save-logs',
  EXCEL_REMEMBER_FILENAME: 'excel-remember-filename',
  EXCEL_REMEMBER_DESTINATION: 'excel-remember-destination',
  EXCEL_SLICE_CHECK: 'excel-slice-check',
  EXCEL_MISMATCHES: 'excel-mismatches',
  EXCEL_SANITIZER: 'excel-sanitizer',
  EXCEL_LOGS_CONTENT: 'excel-logs-content',
  EXCEL_PROGRESS: 'excel-progress',
  EXCEL_PROGRESS_FILL: 'excel-progress-fill',
  EXCEL_PROGRESS_PERCENT: 'excel-progress-percent',
  EXCEL_PROGRESS_STATUS: 'excel-progress-status',
  EXCEL_PROGRESS_DETAIL: 'excel-progress-detail',
  EXCEL_FILES_PREVIEW: 'excel-files-preview',
  
  // Settings
  SETTINGS_2BSP: 'settings-2bsp',
  SETTINGS_ORDER: 'settings-order',
  SETTINGS_AUTOOPEN: 'settings-autoopen',
  SETTINGS_TOKENS: 'settings-tokens',
  SETTINGS_MAX_FILES: 'settings-max-files',
  SETTINGS_BATCH_NOTIFICATIONS: 'settings-batch-notifications',
  SETTINGS_BATCH_LOGS: 'settings-batch-logs',
  SETTINGS_POWER_SAVE: 'settings-power-save',
  SETTINGS_STARTUP_CHECK: 'settings-startup-check',
  SETTINGS_MINIMIZE_TRAY: 'settings-minimize-tray',
  APP_VERSION_DISPLAY: 'app-version',
  DATA_PATH: 'data-path',
  LAST_UPDATE_CHECK_DISPLAY: 'last-update-check',
  
  // Інші
  CHOOSE_RESULT: 'choose-result',
  CHOOSE_FOLDER: 'choose-folder',
  CHOOSE_EXCEL: 'choose-excel',
  EXCEL_PATH: 'excel-path',
  FOLDER_PATH: 'folder-path',
  BTN_SET_LICENSE: 'btn-set-license',
  GATE_LICENSE_BTN: 'gate-license-btn',
  BTN_AUTO_UPDATE: 'btn-auto-update',
  BTN_MANUAL_DOWNLOAD: 'btn-manual-download',
  BTN_CANCEL_UPDATE: 'btn-cancel-update',
  BTN_RETRY_UPDATE: 'btn-retry-update',
  BTN_SAVE_LOG: 'btn-save-log',
  EXPORT_SETTINGS: 'export-settings',
  IMPORT_SETTINGS: 'import-settings',
  RESET_SETTINGS: 'reset-settings'
} as const;

// FIXED: Таймаути (були захардкоджені як магічні числа)
export const TIMEOUTS = {
  LICENSE_CHECK_DELAY: 1000, // мс (рядок 85 в main.ts)
  THEME_INIT_DELAY: 500, // мс (рядок 271 в main.ts)
  SETTINGS_LOAD_DELAY: 100, // мс (рядок 1495 в main.ts)
  ANIMATION_DELAY: 100, // мс (рядок 1509 в main.ts)
  ANIMATION_STAGGER: 150, // мс (рядок 1546 в main.ts)
  FILE_PICKER_TIMEOUT: 12000 // мс (рядок 115 в main.ts)
} as const;

// FIXED: CSS класи (використовувались як magic strings)
export const CSS_CLASSES = {
  ROUTE: 'route',
  NAV_LINK: 'nav a',
  ACTIVE: 'active',
  ANIMATED: 'animated',
  RIPPLE_EFFECT: 'ripple-effect',
  SETTINGS_SECTION: 'settings-section',
  SETTINGS_BUTTONS: 'settings-buttons',
  BTN: 'btn',
  LIGHT: 'light',
  DARK: 'dark',
  SYSTEM: 'system',
  EMPTY: 'empty',
  FILE_PICKER: 'file-picker',
  FILE_NAME: 'file-name',
  FILE_BTN: 'file-btn',
  LICENSE_STATUS: 'license-status'
} as const;

// FIXED: HTML атрибути (використовувались як magic strings)
export const HTML_ATTRIBUTES = {
  HREF: 'href',
  HIDDEN: 'hidden',
  DATA_ROUTE: 'data-route',
  DATA_PLACEHOLDER: 'data-placeholder',
  DATA_INITIALIZED: 'data-initialized',
  NAME: 'name',
  TYPE: 'type'
} as const;

// FIXED: Типи подій
export const EVENT_TYPES = {
  CLICK: 'click',
  CHANGE: 'change',
  HASH_CHANGE: 'hashchange',
  LOAD: 'load',
  DOM_CONTENT_LOADED: 'DOMContentLoaded'
} as const;

// FIXED: Селектори для radio buttons
export const RADIO_SELECTORS = {
  SOURCE_TYPE: 'input[name="source-type"]',
  SOURCE_TYPE_CHECKED: 'input[name="source-type"]:checked',
  EXCEL_MODE: 'input[name="excel-mode"]',
  EXCEL_MODE_CHECKED: 'input[name="excel-mode"]:checked'
} as const;

// FIXED: Значення для source type
export const SOURCE_TYPES = {
  SINGLE_FILE: 'single-file',
  MULTIPLE_FILES: 'multiple-files',
  FOLDER: 'folder'
} as const;

// FIXED: Значення для Excel mode
export const EXCEL_MODES = {
  BZ: 'БЗ',
  ZS: 'ЗС',
  BOTH: 'Обidва'
} as const;

// FIXED: Локалізація (якщо знадобиться в майбутньому)
export const LOCALE = 'uk-UA';

// FIXED: Стани оновлення
export const UPDATE_STATES = {
  DOWNLOADING: 'downloading',
  VERIFYING: 'verifying',
  INSTALLING: 'installing',
  RESTARTING: 'restarting',
  UP_TO_DATE: 'uptodate',
  FAILED: 'failed'
} as const;

// FIXED: Іконки для рівнів логування
export const LOG_ICONS = {
  info: 'ℹ️',
  warning: '⚠️',
  error: '❌'
} as const;

// FIXED: Розширення файлів
export const FILE_EXTENSIONS = {
  DOCX: '.docx',
  XLSX: '.xlsx',
  JSON: '.json',
  TXT: '.txt'
} as const;

// FIXED: MIME типи
export const MIME_TYPES = {
  JSON: 'application/json',
  TEXT: 'text/plain'
} as const;

// FIXED: Налаштування відображення файлів
export const FILE_PREVIEW = {
  MAX_SKIPPED_TO_SHOW: 3, // максимум пропущених файлів для відображення
  MAX_CONFLICTS_TO_SHOW: 5 // максимум конфліктів для відображення
} as const;
