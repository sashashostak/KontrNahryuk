/**
 * BojecIndex.ts - Індексація та агрегація даних про бійців
 * Збирає дані з усіх файлів, відстежує псевдоніми та входження
 * 
 * ⚠️ ВАЖЛИВО: Цей модуль НЕ ВИКОРИСТОВУЄТЬСЯ у поточній версії
 * Бізнес-логіка вкладки "Стройовка" відключена згідно з ТЗ (Фаза 4: Індексування та аналіз)
 * 
 * Модуль збережено для можливого відновлення функціоналу.
 */

import { FileProcessResult } from './ExcelReader'

export interface ServicePeriod {
  arrival: Date      // Дата прибуття (входження)
  departure: Date    // Дата вибуття (останнє входження + 1 день до наступного періоду або кінець)
  duration: number   // Тривалість у днях
}

export interface BojecRecord {
  pibKey: string           // Нормалізований ПІБ як ключ
  originalPib: string      // Оригінальний ПІБ з файлу
  pseudonimy: Set<string>  // Унікальні псевдоніми
  totalOccurrences: number // Загальна кількість входжень
  uniqueFiles: Set<string> // Унікальні файли де зустрічається
  servicePeriods: ServicePeriod[] // Періоди служби (може бути декілька)
  dateRange: {             // Загальний діапазон дат
    earliest: Date | null
    latest: Date | null
  }
  occurrences: Array<{     // Детальна інформація про входження
    filePath: string
    fileName: string
    sheetName: string
    date: Date | null
    row: number
    pseudo: string | null
  }>
}

export interface ConflictInfo {
  pibKey: string
  type: 'pseudo' | 'name'
  values: string[]
  occurrences: any[]
}

export interface IndexStats {
  totalFighters: number
  totalOccurrences: number
  uniqueFiles: number
  dateRange: {
    earliest: Date | null
    latest: Date | null
  }
  conflicts: {
    pseudoConflicts: number
  }
  validation: {
    emptyPseudos: number
    invalidPibs: number
  }
}

export class BojecIndex {
  private fighters: Map<string, BojecRecord> = new Map()
  private conflicts: ConflictInfo[] = []

  /**
   * Нормалізує ПІБ для використання як ключ
   */
  private normalizePib(pib: string): string {
    return pib
      .toLowerCase()
      .trim()
      .replace(/\s+/g, ' ')      // Множинні пробіли → один пробіл
      .replace(/['']/g, "'")      // Різні апострофи → стандартний
      .replace(/[""]/g, '"')      // Різні лапки → стандартні
  }

  /**
   * Додає дані з одного файлу
   */
  addFileData(fileResult: FileProcessResult): void {
    if (!fileResult.processed) return

    for (const sheet of fileResult.sheets) {
      if (!sheet.dataRows || sheet.dataRows.length === 0) continue

      for (const row of sheet.dataRows) {
        if (!row.pib || row.pib.trim().length === 0) continue

        const pibKey = this.normalizePib(row.pib)
        const pseudo = row.pseudo?.trim() || null

        // Отримати або створити запис
        if (!this.fighters.has(pibKey)) {
          this.fighters.set(pibKey, {
            pibKey,
            originalPib: row.pib,
            pseudonimy: new Set(),
            totalOccurrences: 0,
            uniqueFiles: new Set(),
            servicePeriods: [],
            dateRange: { earliest: null, latest: null },
            occurrences: []
          })
        }

        const fighter = this.fighters.get(pibKey)!

        // Додати псевдонім
        if (pseudo && pseudo.length > 0) {
          fighter.pseudonimy.add(pseudo)
        }

        // Збільшити лічильник
        fighter.totalOccurrences++

        // Додати файл
        fighter.uniqueFiles.add(fileResult.fileName)

        // Оновити діапазон дат
        if (sheet.date) {
          if (!fighter.dateRange.earliest || sheet.date < fighter.dateRange.earliest) {
            fighter.dateRange.earliest = sheet.date
          }
          if (!fighter.dateRange.latest || sheet.date > fighter.dateRange.latest) {
            fighter.dateRange.latest = sheet.date
          }
        }

        // Додати деталі входження
        fighter.occurrences.push({
          filePath: fileResult.filePath,
          fileName: fileResult.fileName,
          sheetName: sheet.sheetName,
          date: sheet.date,
          row: row.row,
          pseudo
        })
      }
    }
  }

  /**
   * Обчислює періоди служби для всіх бійців
   */
  calculateServicePeriods(): void {
    for (const fighter of this.fighters.values()) {
      fighter.servicePeriods = this.calculateFighterServicePeriods(fighter)
    }
  }

  /**
   * Обчислює періоди служби для одного бійця
   */
  private calculateFighterServicePeriods(fighter: BojecRecord): ServicePeriod[] {
    // Отримуємо унікальні дати з входжень, відфільтровуємо null та сортуємо
    const uniqueDates = Array.from(new Set(
      fighter.occurrences
        .map(occ => occ.date)
        .filter((date): date is Date => date !== null)
        .map(date => date.getTime())
    )).map(time => new Date(time))
      .sort((a, b) => a.getTime() - b.getTime())
    
    if (uniqueDates.length === 0) return []
    if (uniqueDates.length === 1) {
      return [{
        arrival: uniqueDates[0],
        departure: uniqueDates[0], 
        duration: 1
      }]
    }

    const periods: ServicePeriod[] = []
    let currentPeriodStart = uniqueDates[0]
    let previousDate = uniqueDates[0]

    for (let i = 1; i < uniqueDates.length; i++) {
      const currentDate = uniqueDates[i]
      const daysDiff = Math.floor((currentDate.getTime() - previousDate.getTime()) / (24 * 60 * 60 * 1000))

      // Якщо різниця більше 7 днів, закінчуємо попередній період і починаємо новий
      if (daysDiff > 7) {
        // Закінчуємо попередній період
        periods.push({
          arrival: currentPeriodStart,
          departure: previousDate,
          duration: Math.floor((previousDate.getTime() - currentPeriodStart.getTime()) / (24 * 60 * 60 * 1000)) + 1
        })

        // Починаємо новий період
        currentPeriodStart = currentDate
      }

      previousDate = currentDate
    }

    // Додаємо останній період
    periods.push({
      arrival: currentPeriodStart,
      departure: previousDate,
      duration: Math.floor((previousDate.getTime() - currentPeriodStart.getTime()) / (24 * 60 * 60 * 1000)) + 1
    })

    return periods
  }

  /**
   * Аналізує конфлікти (поки що заглушка)
   */
  analyzeConflicts(): void {
    // Поки що не реалізовано
  }

  /**
   * Вирішує конфлікти (поки що заглушка)
   */
  resolveConflicts(): void {
    // Поки що не реалізовано
  }

  /**
   * Отримати бійця за ключем
   */
  getFighter(pibKey: string): BojecRecord | undefined {
    return this.fighters.get(pibKey)
  }

  /**
   * Отримати всіх бійців
   */
  getAllFighters(): BojecRecord[] {
    return Array.from(this.fighters.values())
  }

  /**
   * Отримати конфлікти
   */
  getConflicts(): ConflictInfo[] {
    return this.conflicts
  }

  /**
   * Отримати статистику
   */
  getStats(): IndexStats {
    let totalOccurrences = 0
    let earliestDate: Date | null = null
    let latestDate: Date | null = null
    const uniqueFiles = new Set<string>()
    let emptyPseudos = 0
    let invalidPibs = 0

    for (const fighter of this.fighters.values()) {
      totalOccurrences += fighter.totalOccurrences

      // Збір унікальних файлів
      for (const file of fighter.uniqueFiles) {
        uniqueFiles.add(file)
      }

      // Діапазон дат
      if (fighter.dateRange.earliest) {
        if (!earliestDate || fighter.dateRange.earliest < earliestDate) {
          earliestDate = fighter.dateRange.earliest
        }
      }
      if (fighter.dateRange.latest) {
        if (!latestDate || fighter.dateRange.latest > latestDate) {
          latestDate = fighter.dateRange.latest
        }
      }

      // Валідація
      if (fighter.pseudonimy.size === 0) {
        emptyPseudos++
      }
    }

    return {
      totalFighters: this.fighters.size,
      totalOccurrences,
      uniqueFiles: uniqueFiles.size,
      dateRange: {
        earliest: earliestDate,
        latest: latestDate
      },
      conflicts: {
        pseudoConflicts: this.conflicts.length
      },
      validation: {
        emptyPseudos,
        invalidPibs
      }
    }
  }
}