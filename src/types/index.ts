/**
 * Загальні типи проекту
 * FIXED: Винесено з main.ts
 */

export type Route = '/functions' | '/excel' | '/batch' | '/notes' | '/updates' | '/settings';

export type Mode = 'БЗ' | 'ЗС' | 'Обидва';

export interface StartProcessPayload {
  srcFolder: string;
  dstPath: string;
  mode: Mode;
  dstSheetPassword?: string;
  configPath?: string;
}
