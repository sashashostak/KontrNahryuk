/**
 * Конфігурація пресетів для зведення Excel файлів
 * БЗ - Бойове забезпечення
 * ЗС - Зберігання спорядження
 */

import { Preset, Rule, Config } from './types';

// Правила для режиму БЗ
const BZ_RULES: Rule[] = [
  { key: '1РСпП', tokens: ['1РСпП'] },
  { key: '2РСпП', tokens: ['2РСпП'] },
  { key: '3РСпП', tokens: ['3РСпП'] },
  { key: 'РВПСпП', tokens: ['РВПСпП'] },
  { key: 'МБ', tokens: ['МБ'] },
  { key: 'РБпС', tokens: ['РБпС'] },
  { key: 'ВРСП', tokens: ['ВРСП'] },
  { key: 'ВРЕБ', tokens: ['ВРЕБ'] },
  { key: 'ВІ', tokens: ['ВІ'] },
  { key: 'ВЗ', tokens: ['ВЗ'] },
  { key: 'РМТЗ', tokens: ['РМТЗ'] },
  { key: 'МП', tokens: ['ОСЄ', 'МП'] },
  { key: '1241', tokens: ['1241'] },
];

// Правила для режиму ЗС
const ZS_RULES: Rule[] = [
  { key: '1РСпП', tokens: ['1РСпП'] },
  { key: '2РСпП', tokens: ['2РСпП'] },
  { key: '3РСпП', tokens: ['3РСпП'] },
  { key: 'РВПСпП', tokens: ['РВПСпП'] },
  { key: 'МБ', tokens: ['МБ'] },
  { key: 'РБпС', tokens: ['РБпС'] },
  { key: 'ВРСП', tokens: ['ВРСП'] },
  { key: 'ВРЕБ', tokens: ['ВРЕБ'] },
  { key: 'ВІ', tokens: ['ВІ'] },
  { key: 'ВЗ', tokens: ['ВЗ'] },
  { key: 'РМТЗ', tokens: ['РМТЗ'] },
  { key: 'МП', tokens: ['ОСЄ', 'МП'] },
];

// Пресет для режиму БЗ
export const Preset_BZ: Preset = {
  SRC_SHEET: 'БЗ',
  DST_SHEET: 'БЗ',
  COL_SUBUNIT: 3, // C (1-based)
  COL_LEFT: 4,    // D (1-based)
  COL_RIGHT: 8,   // H (1-based)
  rules: BZ_RULES
};

// Пресет для режиму ЗС
export const Preset_ZS: Preset = {
  SRC_SHEET: 'ЗС',
  DST_SHEET: 'ЗС',
  COL_SUBUNIT: 2, // B (1-based)
  COL_LEFT: 3,    // C (1-based)
  COL_RIGHT: 8,   // H (1-based)
  rules: ZS_RULES
};

// Вбудована конфігурація
export const DEFAULT_CONFIG: Config = {
  presets: {
    'БЗ': Preset_BZ,
    'ЗС': Preset_ZS
  }
};

/**
 * Завантажує конфігурацію з файлу або повертає вбудовану
 * @param configPath - шлях до JSON файлу конфігурації (опційно)
 * @returns конфігурація пресетів
 */
export async function loadConfig(configPath?: string): Promise<Config> {
  if (!configPath) {
    return DEFAULT_CONFIG;
  }

  try {
    // В реальному застосунку тут би був виклик до file system API
    // const configData = await fs.readFile(configPath, 'utf-8');
    // const config: Config = JSON.parse(configData);
    
    // Поки що повертаємо вбудовану конфігурацію
    console.log(`Завантаження конфігурації з файлу: ${configPath}`);
    return DEFAULT_CONFIG;
  } catch (error) {
    console.warn(`Не вдалося завантажити конфігурацію з ${configPath}, використовую вбудовану`, error);
    return DEFAULT_CONFIG;
  }
}

/**
 * Отримує пресет за режимом
 * @param config - конфігурація
 * @param mode - режим ('БЗ', 'ЗС', 'Обидва')
 * @returns масив пресетів для обробки
 */
export function getPresetsForMode(config: Config, mode: string): Preset[] {
  switch (mode) {
    case 'БЗ':
      return config.presets['БЗ'] ? [config.presets['БЗ']] : [];
    case 'ЗС':
      return config.presets['ЗС'] ? [config.presets['ЗС']] : [];
    case 'Обидва':
      const presets: Preset[] = [];
      if (config.presets['БЗ']) presets.push(config.presets['БЗ']);
      if (config.presets['ЗС']) presets.push(config.presets['ЗС']);
      return presets;
    default:
      throw new Error(`Невідомий режим: ${mode}`);
  }
}

/**
 * Валідує пресет на коректність
 * @param preset - пресет для валідації
 * @returns true якщо валідний
 */
export function validatePreset(preset: Preset): boolean {
  if (!preset.SRC_SHEET || !preset.DST_SHEET) {
    return false;
  }
  
  if (preset.COL_SUBUNIT < 1 || preset.COL_LEFT < 1 || preset.COL_RIGHT < 1) {
    return false;
  }
  
  if (preset.COL_LEFT > preset.COL_RIGHT) {
    return false;
  }
  
  if (!preset.rules || preset.rules.length === 0) {
    return false;
  }
  
  for (const rule of preset.rules) {
    if (!rule.key || !rule.tokens || rule.tokens.length === 0) {
      return false;
    }
  }
  
  return true;
}