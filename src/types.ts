/**
 * Загальні TypeScript типи та інтерфейси для проекту
 * FIXED: Винесено з main.ts для покращення модульності та type safety
 */

// FIXED: Винесено з main.ts (рядок 1)
export type Route = '/functions' | '/excel' | '/dodatok10' | '/shtat-slice' | '/zbd' | '/updates' | '/settings';

// FIXED: Винесено з main.ts (рядок 1568)
export type Mode = 'БЗ' | 'ЗС' | 'Обидва';

// FIXED: Винесено з main.ts (рядки 1570-1575)
export interface StartProcessPayload {
  srcFolder: string;
  dstPath: string;
  mode: Mode;
  dstSheetPassword?: string;
  configPath?: string;
}

// FIXED: Додано інтерфейси для API результатів (неявно використовувались в main.ts)
export interface ApiResult {
  ok: boolean;
  error?: string;
  out?: string;
  stats?: ProcessingStats;
}

export interface ProcessingStats {
  totalDocuments?: number;
  documents?: DocumentStats[];
  paragraphs?: number;
  matched?: number;
}

export interface DocumentStats {
  type: string;
  matched: number;
}

// FIXED: Додано інтерфейс для результату вибору файлу/папки
export interface SelectionResult {
  filePath?: string;
  error?: string;
}

// FIXED: Додано інтерфейс для нотаток (використовується в refreshNotes)
export interface Note {
  text: string;
  createdAt: string | number;
}

// FIXED: Додано інтерфейс для інформації про оновлення
export interface UpdateInfo {
  hasUpdate: boolean;
  latestVersion: string;
  releaseInfo?: ReleaseInfo;
  error?: string;
}

export interface ReleaseInfo {
  html_url: string;
  published_at: string;
  assets?: ReleaseAsset[];
}

export interface ReleaseAsset {
  name: string;
  browser_download_url: string;
}

// FIXED: Додано інтерфейс для результату ліцензії
export interface LicenseResult {
  hasAccess: boolean;
  reason?: string;
  licenseInfo?: LicenseInfo;
}

export interface LicenseInfo {
  plan: string;
  expiresAt?: string;
}

// FIXED: Додано інтерфейс для прогресу оновлення
export interface UpdateProgress {
  percent?: number;
  speedKbps?: number;
  bytesReceived?: number;
  totalBytes?: number;
  percentage?: number;
  message?: string;
}

// FIXED: Додано тип для рівнів логування
export type LogLevel = 'info' | 'warning' | 'error';

// FIXED: Додано інтерфейс для запису логу
export interface LogEntry {
  level: LogLevel;
  message: string;
}

// FIXED: Додано інтерфейс для результату batch обробки
export interface BatchResult {
  success: boolean;
  outputFilePath?: string;
  stats: BatchStats;
  warnings: string[];
  errors: string[];
}

export interface BatchStats {
  filesProcessed: number;
  sheetsProcessed: number;
  fightersFound: number;
  totalOccurrences: number;
  conflicts: number;
  processingTime: number;
}

// FIXED: Додано інтерфейс для прогресу batch обробки
export interface BatchProgress {
  phase: 'scanning' | 'ordering' | 'reading' | 'indexing' | 'writing' | 'complete' | 'error';
  currentFile?: string;
  filesProcessed: number;
  totalFiles: number;
  percentage: number;
  message: string;
  timeElapsed: number;
  estimatedTimeRemaining?: number;
}

// FIXED: Додано інтерфейс для налаштувань
export interface AppSettings {
  isOrder?: boolean;
  autoOpen?: boolean;
  showTokens?: boolean;
  maxFiles?: number;
  batchNotifications?: boolean;
  batchLogs?: boolean;
  powerSave?: boolean;
  startupCheck?: boolean;
  minimizeTray?: boolean;
  theme?: string;
  excelDestinationFile?: string;
  excelRememberDestination?: boolean;
  excelSliceCheck?: boolean;
  excelMismatches?: boolean;
  excelSanitizer?: boolean;
  excelSummarizationMode?: string;
  lastUpdateCheck?: string | null;
}

// FIXED: Додано тип для статусу ліцензії
export type LicenseStatus = 'valid' | 'invalid' | 'pending';

// FIXED: Додано тип для теми
export type Theme = 'light' | 'dark' | 'system';
