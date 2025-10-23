import ExcelJS from 'exceljs';

/**
 * Блок рядків з одним ключем
 */
export interface ContiguousBlock {
  startRow: number;      // Перший рядок блоку
  endRow: number;        // Останній рядок блоку
  rowCount: number;      // Кількість рядків
  key: string;           // Ключ підрозділу
}

/**
 * Знаходить ПЕРШИЙ контігуальний блок у джерелі
 * (всі підряд рядки з одним ключем у колонці B)
 */
export function findSingleContiguousBlock(
  sheet: ExcelJS.Worksheet,
  keyColumn: string,
  key: string
): ContiguousBlock | null {
  
  console.log(`  🔍 Пошук блоку для ключа "${key}"`);
  console.log(`  📍 Колонка: ${keyColumn}`);
  console.log(`  📊 Всього рядків у листі: ${sheet.rowCount}`);
  
  let startRow = 0;
  let endRow = 0;
  let debugInfo: Array<{row: number, raw: any, normalized: string}> = [];
  
  // Збираємо перші 20 значень для діагностики
  let rowsChecked = 0;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      console.log(`  ⏭️  Рядок 1: пропущено (заголовок)`);
      return;
    }
    
    rowsChecked++;
    
    const cell = row.getCell(keyColumn);
    const rawValue = cell.value;
    const normalizedKey = normalizeKey(rawValue);
    
    // Збираємо перші 20 для debug
    if (debugInfo.length < 20) {
      debugInfo.push({
        row: rowNumber,
        raw: rawValue,
        normalized: normalizedKey || '(null)'
      });
    }
    
    // Логування першого значення
    if (rowNumber === 2) {
      console.log(`  🔍 Перший рядок даних (${rowNumber}):`);
      console.log(`     RAW значення:`, rawValue);
      console.log(`     Тип:`, typeof rawValue);
      console.log(`     Після normalize:`, normalizedKey);
    }
    
    if (startRow > 0) return; // Вже знайшли
    
    if (normalizedKey === key.toLowerCase()) {
      startRow = rowNumber;
      console.log(`  ✅ Знайдено початок у рядку ${rowNumber}`);
    }
  });
  
  console.log(`  📊 Перевірено рядків: ${rowsChecked}`);
  
  if (startRow === 0) {
    console.log(`  ❌ Не знайдено початку блоку`);
    console.log(`  📋 Перші 20 значень колонки ${keyColumn}:`);
    debugInfo.forEach(info => {
      console.log(`     Рядок ${info.row}: "${info.raw}" → "${info.normalized}"`);
    });
    return null;
  }
  
  // Шукаємо кінець блоку
  endRow = startRow;
  for (let r = startRow + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const cellValue = row.getCell(keyColumn).value;
    const normalizedKey = normalizeKey(cellValue);
    
    if (normalizedKey === key.toLowerCase()) {
      endRow = r;
    } else {
      break;
    }
  }
  
  console.log(`  ✅ Блок знайдено: рядки ${startRow}-${endRow} (${endRow - startRow + 1} рядків)`);
  
  return {
    startRow,
    endRow,
    rowCount: endRow - startRow + 1,
    key
  };
}

/**
 * Знаходить ВСІ контігуальні блоки у призначенні
 * (можуть бути розкидані по файлу)
 */
export function findAllContiguousBlocks(
  sheet: ExcelJS.Worksheet,
  keyColumn: string,
  key: string
): ContiguousBlock[] {
  const blocks: ContiguousBlock[] = [];
  let inBlock = false;
  let startRow = 0;
  
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Пропускаємо заголовок
    
    const cellValue = row.getCell(keyColumn).value;
    const normalizedKey = normalizeKey(cellValue);
    const matches = normalizedKey === key.toLowerCase();
    
    if (matches && !inBlock) {
      // Початок нового блоку
      inBlock = true;
      startRow = rowNumber;
    } else if (!matches && inBlock) {
      // Кінець блоку
      inBlock = false;
      blocks.push({
        startRow,
        endRow: rowNumber - 1,
        rowCount: rowNumber - startRow,
        key
      });
    }
  });
  
  // Якщо блок триває до кінця листа
  if (inBlock) {
    blocks.push({
      startRow,
      endRow: sheet.rowCount,
      rowCount: sheet.rowCount - startRow + 1,
      key
    });
  }
  
  return blocks;
}

/**
 * Нормалізація ключа з детальним логуванням
 */
function normalizeKey(value: any): string {
  // Перевірка на null/undefined
  if (value === null || value === undefined) {
    return '';
  }
  
  // Якщо це об'єкт (наприклад richText або формула)
  if (typeof value === 'object') {
    // Якщо це richText
    if ('richText' in value) {
      const text = value.richText.map((t: any) => t.text).join('');
      value = text;
    }
    // Якщо це формула з результатом
    else if ('result' in value) {
      value = value.result;
    }
    // Якщо це щось інше - конвертуємо в string
    else {
      value = JSON.stringify(value);
    }
  }
  
  // Конвертуємо в string
  let str = String(value).trim();
  
  // Видаляємо ВСІ пробіли
  str = str.replace(/\s+/g, '');
  
  // toLowerCase
  str = str.toLowerCase();
  
  return str;
}
