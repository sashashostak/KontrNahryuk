/**
 * MonthFileOrder.ts - Розпізнавання та сортування файлів за місяцями
 * Жорстко заданий порядок: листопад → грудень → січень → лютий → березень → квітень → травень → червень → липень → серпень
 */

export interface MonthFile {
  filePath: string
  fileName: string
  monthIndex: number
  monthName: string
  modifiedTime: Date
}

export class MonthFileOrder {
  // Словник українських місяців з варіантами відмінків
  private static readonly MONTH_DICTIONARY: { [key: string]: { index: number, name: string } } = {
    // Листопад
    'листопад': { index: 0, name: 'листопад' },
    'листопада': { index: 0, name: 'листопад' },
    
    // Грудень  
    'грудень': { index: 1, name: 'грудень' },
    'грудня': { index: 1, name: 'грудень' },
    
    // Січень
    'січень': { index: 2, name: 'січень' },
    'січня': { index: 2, name: 'січень' },
    
    // Лютий
    'лютий': { index: 3, name: 'лютий' },
    'лютого': { index: 3, name: 'лютий' },
    
    // Березень
    'березень': { index: 4, name: 'березень' },
    'березня': { index: 4, name: 'березень' },
    
    // Квітень
    'квітень': { index: 5, name: 'квітень' },
    'квітня': { index: 5, name: 'квітень' },
    
    // Травень
    'травень': { index: 6, name: 'травень' },
    'травня': { index: 6, name: 'травень' },
    
    // Червень
    'червень': { index: 7, name: 'червень' },
    'червня': { index: 7, name: 'червень' },
    
    // Липень
    'липень': { index: 8, name: 'липень' },
    'липня': { index: 8, name: 'липень' },
    
    // Серпень
    'серпень': { index: 9, name: 'серпень' },
    'серпня': { index: 9, name: 'серпень' }
  };

  private static readonly MONTH_ORDER = [
    'листопад', 'грудень', 'січень', 'лютий', 'березень', 
    'квітень', 'травень', 'червень', 'липень', 'серпень'
  ];

  /**
   * Розпізнає місяць у назві файлу
   */
  static recognizeMonth(fileName: string): { index: number, name: string } | null {
    const normalized = fileName.toLowerCase().replace(/\.[^/.]+$/, ''); // видаляємо розширення
    
    // Шукаємо збіги у словнику
    for (const [monthWord, monthInfo] of Object.entries(this.MONTH_DICTIONARY)) {
      if (normalized.includes(monthWord)) {
        return monthInfo;
      }
    }
    
    return null;
  }

  /**
   * Сортує файли за заданим порядком місяців
   */
  static sortFilesByMonth(files: MonthFile[]): MonthFile[] {
    return files.sort((a, b) => {
      // Спочатку за індексом місяця
      if (a.monthIndex !== b.monthIndex) {
        return a.monthIndex - b.monthIndex;
      }
      
      // Якщо однакові місяці - за датою модифікації (старіші спочатку)
      return a.modifiedTime.getTime() - b.modifiedTime.getTime();
    });
  }

  /**
   * Обробляє масив шляхів до файлів і повертає відсортований список із розпізнаними місяцями
   */
  static async processFiles(filePaths: string[], getFileStats: (path: string) => Promise<{ mtime: Date }>): Promise<{
    validFiles: MonthFile[],
    skippedFiles: { path: string, reason: string }[]
  }> {
    const validFiles: MonthFile[] = [];
    const skippedFiles: { path: string, reason: string }[] = [];

    for (const filePath of filePaths) {
      try {
        const fileName = filePath.split(/[/\\]/).pop() || '';
        
        // Перевірка розширення
        if (!fileName.toLowerCase().endsWith('.xlsx')) {
          skippedFiles.push({ path: filePath, reason: 'Не Excel файл (.xlsx)' });
          continue;
        }

        // Розпізнавання місяця
        const monthInfo = this.recognizeMonth(fileName);
        if (!monthInfo) {
          skippedFiles.push({ path: filePath, reason: 'Не вдалось розпізнати місяць у назві' });
          continue;
        }

        // Отримання дати модифікації
        const stats = await getFileStats(filePath);
        
        validFiles.push({
          filePath,
          fileName,
          monthIndex: monthInfo.index,
          monthName: monthInfo.name,
          modifiedTime: stats.mtime
        });

      } catch (error) {
        skippedFiles.push({ 
          path: filePath, 
          reason: `Помилка обробки: ${error instanceof Error ? error.message : 'Невідома помилка'}` 
        });
      }
    }

    return {
      validFiles: this.sortFilesByMonth(validFiles),
      skippedFiles
    };
  }

  /**
   * Отримує порядковий номер місяця за назвою
   */
  static getMonthOrder(): string[] {
    return [...this.MONTH_ORDER];
  }

  /**
   * Перевіряє чи є дублікати місяців
   */
  static findDuplicateMonths(files: MonthFile[]): { [monthName: string]: MonthFile[] } {
    const monthGroups: { [monthName: string]: MonthFile[] } = {};
    
    for (const file of files) {
      if (!monthGroups[file.monthName]) {
        monthGroups[file.monthName] = [];
      }
      monthGroups[file.monthName].push(file);
    }
    
    // Повертаємо тільки групи з більш ніж одним файлом
    const duplicates: { [monthName: string]: MonthFile[] } = {};
    for (const [monthName, group] of Object.entries(monthGroups)) {
      if (group.length > 1) {
        duplicates[monthName] = group;
      }
    }
    
    return duplicates;
  }

  /**
   * Форматує інформацію про файл для логу
   */
  static formatFileInfo(file: MonthFile): string {
    return `${file.fileName} (${file.monthName}, ${file.modifiedTime.toLocaleDateString('uk-UA')})`;
  }
}