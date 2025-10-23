/**
 * ExcelWriter.ts - Генератор результуючого Excel файлу з агрегованими даними
 * Створює файл з індексом бійців, статистикою та конфліктами
 * 
 * ⚠️ ВАЖЛИВО: Цей модуль НЕ ВИКОРИСТОВУЄТЬСЯ у поточній версії
 * Бізнес-логіка вкладки "Стройовка" відключена згідно з ТЗ (Фаза 5: Запис результату)
 * 
 * Модуль збережено для можливого відновлення функціоналу.
 */

import ExcelJS from 'exceljs'
import { BojecIndex, BojecRecord, ConflictInfo, IndexStats } from './BojecIndex'

export interface ExcelWriterOptions {
  outputPath: string
  includeStats: boolean
  includeConflicts: boolean
  includeOccurrences: boolean
  sheetNames?: {
    main: string
    stats: string
    conflicts: string
    occurrences: string
  }
}

export class ExcelWriter {
  private static readonly DEFAULT_SHEET_NAMES = {
    main: 'Індекс бійців',
    stats: 'Статистика',
    conflicts: 'Конфлікти',
    occurrences: 'Всі входження'
  }

  /**
   * Створює Excel файл з результатами
   */
  static async writeIndex(
    bojecIndex: BojecIndex,
    options: ExcelWriterOptions
  ): Promise<void> {
    const workbook = new ExcelJS.Workbook()
    const sheetNames = { ...this.DEFAULT_SHEET_NAMES, ...options.sheetNames }
    
    // Отримуємо дані
    const fighters = bojecIndex.getAllFighters()
    const stats = bojecIndex.getStats()
    const conflicts = bojecIndex.getConflicts()

    // Основний лист з індексом бійців
    await this.createMainSheet(workbook, sheetNames.main, fighters)

    // Лист статистики
    if (options.includeStats) {
      await this.createStatsSheet(workbook, sheetNames.stats, stats)
    }

    // Лист конфліктів
    if (options.includeConflicts && conflicts.length > 0) {
      await this.createConflictsSheet(workbook, sheetNames.conflicts, conflicts, fighters)
    }

    // Лист всіх входжень
    if (options.includeOccurrences) {
      await this.createOccurrencesSheet(workbook, sheetNames.occurrences, fighters)
    }

    // Зберігаємо файл
    await workbook.xlsx.writeFile(options.outputPath)
  }

  /**
   * Створює основний лист з індексом бійців
   */
  private static async createMainSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    fighters: BojecRecord[]
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName)
    
    // Визначаємо максимальну кількість періодів служби
    const maxPeriods = Math.max(...fighters.map(f => f.servicePeriods.length), 1)
    
    // Створюємо заголовки з динамічними колонками періодів
    const headers = [
      'ПІБ',
      'Псевдоніми',
      'Кількість входжень',
      'Унікальних файлів'
    ]
    
    // Додаємо колонки для кожного періоду служби
    for (let i = 1; i <= maxPeriods; i++) {
      headers.push(`Прибуття ${i}`)
      headers.push(`Вибуття ${i}`)
      headers.push(`Днів ${i}`)
    }
    
    const headerRow = worksheet.addRow(headers)
    
    // Стиль заголовків
    headerRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true }
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      }
      cell.alignment = { horizontal: 'center' }
    })

    // Данні бійців
    for (const fighter of fighters) {
      // Базові дані
      const rowData = [
        fighter.originalPib,
        Array.from(fighter.pseudonimy).join(', ') || '(не вказано)',
        fighter.totalOccurrences,
        fighter.uniqueFiles.size
      ]
      
      // Додаємо дані по кожному періоду служби
      for (let i = 0; i < maxPeriods; i++) {
        if (i < fighter.servicePeriods.length) {
          const period = fighter.servicePeriods[i]
          rowData.push(this.formatDate(period.arrival))
          rowData.push(this.formatDate(period.departure))
          rowData.push(period.duration.toString())
        } else {
          // Порожні колонки для бійців з меншою кількістю періодів
          rowData.push('')
          rowData.push('')
          rowData.push('')
        }
      }
      
      const row = worksheet.addRow(rowData)

      // Стиль рядків
      row.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
        
        // Вирівнювання для базових колонок
        if (colNumber >= 3 && colNumber <= 4) {
          cell.alignment = { horizontal: 'center' }
        }
        
        // Вирівнювання для колонок періодів служби (після 4-ої колонки)
        if (colNumber > 4) {
          const periodColumnIndex = (colNumber - 5) % 3 // 0=прибуття, 1=вибуття, 2=днів
          if (periodColumnIndex === 0 || periodColumnIndex === 1) {
            // Колонки прибуття/вибуття - по центру
            cell.alignment = { horizontal: 'center' }
          } else {
            // Колонки кількість днів - по центру
            cell.alignment = { horizontal: 'center' }
          }
        }
      })
    }

    // Налаштування ширини колонок
    worksheet.columns.forEach((column, index) => {
      const colNumber = index + 1
      
      if (colNumber <= 4) {
        // Базові колонки - автоширина
        let maxWidth = 10
        if (column.eachCell) {
          column.eachCell({ includeEmpty: false }, (cell) => {
            const cellLength = cell.value ? cell.value.toString().length : 0
            maxWidth = Math.max(maxWidth, Math.min(cellLength + 2, 50))
          })
        }
        column.width = maxWidth
      } else {
        // Колонки періодів служби
        const periodColumnIndex = (colNumber - 5) % 3 // 0=прибуття, 1=вибуття, 2=днів
        if (periodColumnIndex === 0 || periodColumnIndex === 1) {
          // Колонки прибуття/вибуття - ширина 11
          column.width = 11
        } else {
          // Колонки кількість днів - ширина 7.5
          column.width = 7.5
        }
      }
    })

    // Заморозка заголовку
    worksheet.views = [{ state: 'frozen', ySplit: 1 }]
  }

  /**
   * Створює лист статистики
   */
  private static async createStatsSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    stats: IndexStats
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName)
    
    // Загальна статистика
    worksheet.addRow(['ЗАГАЛЬНА СТАТИСТИКА']).font = { bold: true, size: 14 }
    worksheet.addRow([])
    
    worksheet.addRow(['Загальна кількість бійців:', stats.totalFighters])
    worksheet.addRow(['Загальна кількість входжень:', stats.totalOccurrences])
    worksheet.addRow(['Унікальних файлів оброблено:', stats.uniqueFiles])
    
    if (stats.dateRange.earliest && stats.dateRange.latest) {
      worksheet.addRow(['Період даних:', 
        `${this.formatDate(stats.dateRange.earliest)} - ${this.formatDate(stats.dateRange.latest)}`
      ])
    }
    
    worksheet.addRow([])
    
    // Конфлікти
    worksheet.addRow(['КОНФЛІКТИ']).font = { bold: true, size: 12 }
    // Конфлікти частин видалено
    worksheet.addRow(['Конфлікти псевдонімів:', stats.conflicts.pseudoConflicts])
    
    worksheet.addRow([])
    
    // Проблеми валідації
    worksheet.addRow(['ВАЛІДАЦІЯ ДАНИХ']).font = { bold: true, size: 12 }
    worksheet.addRow(['Невалідних ПІБ:', stats.validation.invalidPibs])
    // Порожні частини видалено
    worksheet.addRow(['Порожніх псевдонімів:', stats.validation.emptyPseudos])
    
    // Стилізація
    worksheet.getColumn(1).width = 30
    worksheet.getColumn(2).width = 20
    
    worksheet.eachRow((row, rowNumber) => {
      if (row.getCell(1).font?.bold) {
        row.getCell(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD0D0D0' }
        }
      }
    })
  }

  /**
   * Створює лист конфліктів
   */
  private static async createConflictsSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    conflicts: ConflictInfo[],
    fighters: BojecRecord[]
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName)
    
    // Заголовки
    const headers = [
      'ПІБ',
      'Тип конфлікту',
      'Варіанти',
      'Найчастіший варіант',
      'Кількість файлів',
      'Загальна кількість входжень'
    ]
    
    const headerRow = worksheet.addRow(headers)
    headerRow.font = { bold: true }
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFCCCC' }
    }

    // Конфлікти
    for (const conflict of conflicts) {
      const fighter = fighters.find(f => f.pibKey === conflict.pibKey)
      if (!fighter) continue

      const mostFrequent = conflict.occurrences.reduce((prev: any, curr: any) => 
        curr.count > prev.count ? curr : prev
      )

      worksheet.addRow([
        fighter.originalPib,
        'Псевдоніми', // Тільки псевдоніми залишилися
        conflict.values.join(', '),
        `${mostFrequent.value} (${mostFrequent.count} разів)`,
        fighter.uniqueFiles.size,
        fighter.totalOccurrences
      ])
    }

    // Автоширина
    worksheet.columns.forEach((column) => {
      let maxWidth = 15
      if (column.eachCell) {
        column.eachCell({ includeEmpty: false }, (cell) => {
          const cellLength = cell.value ? cell.value.toString().length : 0
          maxWidth = Math.max(maxWidth, Math.min(cellLength + 2, 60))
        })
      }
      column.width = maxWidth
    })

    worksheet.views = [{ state: 'frozen', ySplit: 1 }]
  }

  /**
   * Створює лист з усіма входженнями
   */
  private static async createOccurrencesSheet(
    workbook: ExcelJS.Workbook,
    sheetName: string,
    fighters: BojecRecord[]
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(sheetName)
    
    // Заголовки
    const headers = [
      'ПІБ',
      'Псевдонім',
      'Дата',
      'Файл',
      'Лист',
      'Рядок'
    ]
    
    const headerRow = worksheet.addRow(headers)
    headerRow.font = { bold: true }
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0F0E0' }
    }

    // Всі входження
    const allOccurrences: Array<{
      pib: string
      occurrence: typeof fighters[0]['occurrences'][0]
    }> = []

    for (const fighter of fighters) {
      for (const occurrence of fighter.occurrences) {
        allOccurrences.push({
          pib: fighter.originalPib,
          occurrence
        })
      }
    }

    // Сортуємо по даті (враховуємо null дати)
    allOccurrences.sort((a, b) => {
      const dateA = a.occurrence.date?.getTime() ?? 0
      const dateB = b.occurrence.date?.getTime() ?? 0
      return dateA - dateB
    })

    // Додаємо рядки
    for (const { pib, occurrence } of allOccurrences) {
      worksheet.addRow([
        pib,
        occurrence.pseudo || '(не вказано)',
        occurrence.date ? this.formatDate(occurrence.date) : '(не відомо)',
        occurrence.fileName,
        occurrence.sheetName,
        occurrence.row
      ])
    }

    // Автоширина
    worksheet.columns.forEach((column, index) => {
      let maxWidth = 12
      if (index === 0) maxWidth = 25 // ПІБ
      if (index === 4) maxWidth = 30 // Файл
      if (index === 5) maxWidth = 20 // Лист
      
      if (column.eachCell) {
        column.eachCell({ includeEmpty: false }, (cell) => {
          const cellLength = cell.value ? cell.value.toString().length : 0
          maxWidth = Math.max(maxWidth, Math.min(cellLength + 2, 50))
        })
      }
      column.width = maxWidth
    })

    worksheet.views = [{ state: 'frozen', ySplit: 1 }]
  }

  /**
   * Форматує дату
   */
  private static formatDate(date: Date): string {
    return date.toLocaleDateString('uk-UA', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    })
  }

  /**
   * Обчислює період між датами
   */
  private static calculatePeriod(start: Date, end: Date): string {
    if (start.getTime() === end.getTime()) {
      return '1 день'
    }
    
    const diffTime = end.getTime() - start.getTime()
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1
    
    if (diffDays < 30) {
      return `${diffDays} днів`
    }
    
    const diffMonths = Math.round(diffDays / 30)
    if (diffMonths < 12) {
      return `${diffMonths} місяців`
    }
    
    const diffYears = Math.round(diffDays / 365)
    return `${diffYears} років`
  }

  /**
   * Створює базові налаштування для Excel файлу
   */
  static createDefaultOptions(outputPath: string): ExcelWriterOptions {
    return {
      outputPath,
      includeStats: true,
      includeConflicts: true,
      includeOccurrences: false, // За замовчуванням відключено через великий розмір
      sheetNames: this.DEFAULT_SHEET_NAMES
    }
  }

  /**
   * Валідує шлях до файлу
   */
  static validateOutputPath(outputPath: string): boolean {
    return outputPath.endsWith('.xlsx') && outputPath.length > 5
  }

  /**
   * Генерує пропоноване ім'я файлу
   */
  static generateFileName(inputDir: string): string {
    const now = new Date()
    const dateStr = now.toISOString().split('T')[0] // YYYY-MM-DD
    const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-') // HH-MM-SS
    
    const dirName = inputDir.split(/[/\\]/).pop() || 'batch'
    return `Індекс_бійців_${dirName}_${dateStr}_${timeStr}.xlsx`
  }
}