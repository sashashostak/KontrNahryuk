/**
 * Константи налаштувань
 * FIXED: Винесено магічні рядки з main.ts
 */

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
  LAST_UPDATE_CHECK: 'lastUpdateCheck',
  EXCEL_DESTINATION_FILE: 'excelDestinationFile',
  EXCEL_REMEMBER_DESTINATION: 'excelRememberDestination',
  EXCEL_SLICE_CHECK: 'excelSliceCheck',
  EXCEL_MISMATCHES: 'excelMismatches',
  EXCEL_SANITIZER: 'excelSanitizer',
  EXCEL_SUMMARIZATION_MODE: 'excelSummarizationMode'
} as const;

export const APP_VERSION = '1.3.0';
