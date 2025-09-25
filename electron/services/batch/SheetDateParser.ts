/**
 * SheetDateParser.ts - Парсинг дат з назв листів Excel
 * Підтримувані формати:
 * - dd.MM.yyyy, dd.MM.yy
 * - yyyy-MM-dd  
 * - dd <місяць_словом> yyyy (українська)
 */

export interface SheetInfo {
  name: string
  date: Date | null
  originalName: string
  parseError?: string
}

export class SheetDateParser {
  // Словник українських місяців
  private static readonly UKRAINIAN_MONTHS: { [key: string]: number } = {
    'січень': 0, 'січня': 0,
    'лютий': 1, 'лютого': 1,
    'березень': 2, 'березня': 2,
    'квітень': 3, 'квітня': 3,
    'травень': 4, 'травня': 4,
    'червень': 5, 'червня': 5,
    'липень': 6, 'липня': 6,
    'серпень': 7, 'серпня': 7,
    'вересень': 8, 'вересня': 8,
    'жовтень': 9, 'жовтня': 9,
    'листопад': 10, 'листопада': 10,
    'грудень': 11, 'грудня': 11
  };

  /**
   * Очищає назву листа від префіксів та зайвих пробілів
   */
  private static cleanSheetName(sheetName: string): string {
    let cleaned = sheetName.trim();
    
    // Видалити загальні префікси (БЗ, ЗТ, СП тощо) з пробілами
    cleaned = cleaned.replace(/^[А-ЯІЇЄ]{1,4}\s+/i, '');
    
    // Видалити зайві пробіли
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    return cleaned;
  }

  /**
   * Визначає рік на основі місяця
   * Листопад та грудень відносяться до 2024 року
   */
  private static getYearForMonth(month: number): number {
    // month є 0-based (0 = січень, 11 = грудень)
    // Листопад (10) та грудень (11) = 2024
    if (month === 10 || month === 11) { // листопад або грудень
      return 2024;
    }
    
    // Решта місяців - поточний рік
    return new Date().getFullYear();
  }

  /**
   * Парсить дату з назви листа
   */
  static parseDate(sheetName: string): Date | null {
    const trimmed = sheetName.trim();
    
    // Видалити префікси та зайві пробіли
    const cleanedName = this.cleanSheetName(trimmed);

    // Формат: dd.MM.yyyy або dd.MM.yy
    const dotFormat = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/;
    const dotMatch = cleanedName.match(dotFormat);
    if (dotMatch) {
      const day = parseInt(dotMatch[1]);
      const month = parseInt(dotMatch[2]) - 1; // Місяці у Date 0-based
      let year = parseInt(dotMatch[3]);
      
      // Обробка 2-значного року
      if (year < 100) {
        year += year < 50 ? 2000 : 1900; // 00-49 -> 2000-2049, 50-99 -> 1950-1999
      }
      
      const date = new Date(year, month, day);
      if (this.isValidDate(date, day, month, year)) {
        return date;
      }
    }

    // Формат: dd.MM (без року) - визначаємо рік по місяцю
    const shortDotFormat = /^(\d{1,2})\.(\d{1,2})$/;
    const shortDotMatch = cleanedName.match(shortDotFormat);
    if (shortDotMatch) {
      const day = parseInt(shortDotMatch[1]);
      const month = parseInt(shortDotMatch[2]) - 1; // Місяці у Date 0-based
      const year = this.getYearForMonth(month);
      
      const date = new Date(year, month, day);
      if (this.isValidDate(date, day, month, year)) {
        return date;
      }
    }

    // Формат: yyyy-MM-dd
    const isoFormat = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
    const isoMatch = cleanedName.match(isoFormat);
    if (isoMatch) {
      const year = parseInt(isoMatch[1]);
      const month = parseInt(isoMatch[2]) - 1; // Місяці у Date 0-based
      const day = parseInt(isoMatch[3]);
      
      const date = new Date(year, month, day);
      if (this.isValidDate(date, day, month, year)) {
        return date;
      }
    }

    // Формат: dd <місяць_словом> yyyy (українська)
    const wordFormat = /^(\d{1,2})\s+([а-яіїєґА-ЯІЇЄҐ]+)\s+(\d{4})$/i;
    const wordMatch = cleanedName.match(wordFormat);
    if (wordMatch) {
      const day = parseInt(wordMatch[1]);
      const monthWord = wordMatch[2].toLowerCase();
      const year = parseInt(wordMatch[3]);
      
      const monthNumber = this.UKRAINIAN_MONTHS[monthWord];
      if (monthNumber !== undefined) {
        const date = new Date(year, monthNumber, day);
        if (this.isValidDate(date, day, monthNumber, year)) {
          return date;
        }
      }
    }

    // Формат: dd <місяць_словом> (без року) - визначаємо рік по місяцю
    const wordFormatNoYear = /^(\d{1,2})\s+([а-яіїєґА-ЯІЇЄҐ]+)$/i;
    const wordNoYearMatch = cleanedName.match(wordFormatNoYear);
    if (wordNoYearMatch) {
      const day = parseInt(wordNoYearMatch[1]);
      const monthWord = wordNoYearMatch[2].toLowerCase();
      
      const monthNumber = this.UKRAINIAN_MONTHS[monthWord];
      if (monthNumber !== undefined) {
        const year = this.getYearForMonth(monthNumber);
        const date = new Date(year, monthNumber, day);
        if (this.isValidDate(date, day, monthNumber, year)) {
          return date;
        }
      }
    }

    // Формат: тільки <місяць_словом> (без дня і року) - визначаємо як 1 число місяця
    const monthOnlyFormat = /^([а-яіїєґА-ЯІЇЄҐ]+)$/i;
    const monthOnlyMatch = cleanedName.match(monthOnlyFormat);
    if (monthOnlyMatch) {
      const monthWord = monthOnlyMatch[1].toLowerCase();
      
      const monthNumber = this.UKRAINIAN_MONTHS[monthWord];
      if (monthNumber !== undefined) {
        const year = this.getYearForMonth(monthNumber);
        const date = new Date(year, monthNumber, 1);
        if (this.isValidDate(date, 1, monthNumber, year)) {
          return date;
        }
      }
    }

    return null;
  }

  /**
   * Перевіряє валідність створеної дати
   */
  private static isValidDate(date: Date, expectedDay: number, expectedMonth: number, expectedYear: number): boolean {
    return date.getDate() === expectedDay && 
           date.getMonth() === expectedMonth && 
           date.getFullYear() === expectedYear &&
           expectedYear >= 1900 && 
           expectedYear <= 2100; // Розумні межі
  }

  /**
   * Обробляє масив назв листів і повертає інформацію з датами
   */
  static processSheets(sheetNames: string[]): SheetInfo[] {
    return sheetNames.map(name => {
      const date = this.parseDate(name);
      const sheetInfo: SheetInfo = {
        name,
        date,
        originalName: name
      };

      if (!date) {
        sheetInfo.parseError = 'Не вдалось розпізнати дату';
      }

      return sheetInfo;
    });
  }

  /**
   * Сортує листи за датою (зростаюче)
   */
  static sortSheetsByDate(sheets: SheetInfo[]): SheetInfo[] {
    return sheets.sort((a, b) => {
      // Листи з валідними датами йдуть спочатку
      if (a.date && b.date) {
        return a.date.getTime() - b.date.getTime();
      }
      if (a.date && !b.date) return -1;
      if (!a.date && b.date) return 1;
      
      // Якщо обидва без дати - за порядком назв
      return a.name.localeCompare(b.name, 'uk');
    });
  }

  /**
   * Фільтрує тільки листи з валідними датами
   */
  static getValidDateSheets(sheets: SheetInfo[]): SheetInfo[] {
    return sheets.filter(sheet => sheet.date !== null);
  }

  /**
   * Форматує дату для відображення
   */
  static formatDate(date: Date): string {
    return date.toLocaleDateString('uk-UA', {
      day: '2-digit',
      month: '2-digit', 
      year: 'numeric'
    });
  }

  /**
   * Форматує дату для збереження в Excel (dd.MM.yyyy)
   */
  static formatDateForExcel(date: Date): string {
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
  }

  /**
   * Перевіряє чи дата входить у розумні межі для військових звітів
   */
  static isReasonableDate(date: Date): boolean {
    const now = new Date();
    const minDate = new Date(2014, 0, 1); // Початок конфлікту
    const maxDate = new Date(now.getFullYear() + 1, 11, 31); // Рік вперед
    
    return date >= minDate && date <= maxDate;
  }

  /**
   * Знаходить дублікати дат серед листів
   */
  static findDuplicateDates(sheets: SheetInfo[]): { [dateStr: string]: SheetInfo[] } {
    const dateGroups: { [dateStr: string]: SheetInfo[] } = {};
    
    for (const sheet of sheets) {
      if (sheet.date) {
        const dateStr = this.formatDateForExcel(sheet.date);
        if (!dateGroups[dateStr]) {
          dateGroups[dateStr] = [];
        }
        dateGroups[dateStr].push(sheet);
      }
    }
    
    // Повертаємо тільки групи з більш ніж одним листом
    const duplicates: { [dateStr: string]: SheetInfo[] } = {};
    for (const [dateStr, group] of Object.entries(dateGroups)) {
      if (group.length > 1) {
        duplicates[dateStr] = group;
      }
    }
    
    return duplicates;
  }

  /**
   * Створює діапазон дат для перевірки послідовності
   */
  static getDateRange(sheets: SheetInfo[]): { min: Date | null, max: Date | null } {
    const validDates = sheets
      .map(s => s.date)
      .filter((date): date is Date => date !== null);
    
    if (validDates.length === 0) {
      return { min: null, max: null };
    }
    
    const timestamps = validDates.map(d => d.getTime());
    return {
      min: new Date(Math.min(...timestamps)),
      max: new Date(Math.max(...timestamps))
    };
  }
}