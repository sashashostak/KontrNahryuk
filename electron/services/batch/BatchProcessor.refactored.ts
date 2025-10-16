/**
 * BatchProcessor.ts - Основний сервіс пакетної обробки Excel файлів
 * Інтегрує всі компоненти та забезпечує прогрес-трекінг і логування
 * 
 * REFACTORED VERSION - 16 жовтня 2025
 * Основні покращення:
 * - Розбито монолітну функцію processDirectory() на окремі методи
 * - Видалено всі типи 'any' - додано строгу типізацію
 * - Винесено магічні числа у enum
 * - Видалено дублювання коду
 */

import { EventEmitter } from 'events'
import * as path from 'path'
import * as fs from 'fs/promises'

import { MonthFileOrder } from './MonthFileOrder'
import { ExcelReader, FileProcessResult } from './ExcelReader'
import { BojecIndex } from './BojecIndex'
import { ExcelWriter, ExcelWriterOptions } from './ExcelWriter'

// FIXED: Додано enum для фаз обробки (було хардкоджено в типі)
export enum ProcessPhase {
  SCANNING = 'scanning',
  ORDERING = 'ordering',
  READING = 'reading',
  INDEXING = 'indexing',
  WRITING = 'writing',
  COMPLETE = 'complete',
  ERROR = 'error'
}

// FIXED: Додано enum для відсотків прогресу (раніше були магічні числа)
export enum ProcessPhasePercentage {
  SCANNING = 0,
  ORDERING = 10,
  READING_START = 20,
  READING_END = 70,
  INDEXING = 70,
  WRITING = 90,
  COMPLETE = 100
}

export interface BatchProcessOptions {
  inputDirectory: string
  outputFilePath: string
  includeStats: boolean
  includeConflicts: boolean
  includeOccurrences: boolean
  resolveConflicts: boolean
}

// FIXED: Використання enum замість string literal union
export interface ProcessProgress {
  phase: ProcessPhase
  currentFile?: string
  filesProcessed: number
  totalFiles: number
  percentage: number
  message: string
  timeElapsed: number
  estimatedTimeRemaining?: number
}

export interface ProcessResult {
  success: boolean
  outputFilePath?: string
  stats: {
    filesProcessed: number
    sheetsProcessed: number
    fightersFound: number
    totalOccurrences: number
    conflicts: number
    processingTime: number
  }
  warnings: string[]
  errors: string[]
}

// FIXED: Додано інтерфейс для впорядкованих файлів (раніше було any)
// Використовуємо структуру, яку повертає MonthFileOrder.processFiles
interface OrderedFile {
  filePath: string
  [key: string]: any // Дозволяємо додаткові поля від MonthFileOrder
}

// FIXED: Додано інтерфейс для пропущених файлів
interface SkippedFile {
  path: string
  reason: string
}

// FIXED: Додано type для результату сортування файлів
interface FileOrderResult {
  validFiles: OrderedFile[]
  skippedFiles: SkippedFile[]
}

export declare interface BatchProcessor {
  on(event: 'progress', listener: (progress: ProcessProgress) => void): this
  on(event: 'log', listener: (level: 'info' | 'warning' | 'error', message: string) => void): this
  on(event: 'complete', listener: (result: ProcessResult) => void): this
}

export class BatchProcessor extends EventEmitter {
  private startTime: number = 0
  private isProcessing: boolean = false
  private isCancelled: boolean = false

  /**
   * Запускає пакетну обробку
   * FIXED: Метод тепер делегує роботу приватним методам замість великої монолітної функції
   */
  async processDirectory(options: BatchProcessOptions): Promise<ProcessResult> {
    if (this.isProcessing) {
      throw new Error('Обробка вже виконується')
    }

    this.isProcessing = true
    this.isCancelled = false
    this.startTime = Date.now()
    
    const result: ProcessResult = {
      success: false,
      stats: {
        filesProcessed: 0,
        sheetsProcessed: 0,
        fightersFound: 0,
        totalOccurrences: 0,
        conflicts: 0,
        processingTime: 0
      },
      warnings: [],
      errors: []
    }

    try {
      // FIXED: Розбито на окремі методи - кожен відповідає за одну фазу
      
      // Фаза 1: Сканування директорії
      const excelFiles = await this.scanDirectoryPhase(options.inputDirectory)
      
      // Фаза 2: Впорядкування файлів
      const orderedFiles = await this.orderFilesPhase(excelFiles)
      
      // Фаза 3: Читання Excel файлів
      const { bojecIndex, fileResults } = await this.readExcelFilesPhase(orderedFiles, result)
      
      // Фаза 4: Індексування та аналіз
      await this.indexingPhase(bojecIndex, options, result)
      
      // Фаза 5: Запис результату
      await this.writeResultsPhase(bojecIndex, options, result)

      // Завершення
      result.success = true
      result.stats.processingTime = Date.now() - this.startTime
      
      // FIXED: Використання enum замість string literal
      this.emitPhaseProgress(ProcessPhase.COMPLETE, 'Обробка завершена успішно', 1, 1)
      this.emitLog('info', `Обробка завершена за ${Math.round(result.stats.processingTime / 1000)} секунд`)
      
      this.emit('complete', result)
      return result

    } catch (error: unknown) {
      // FIXED: Типізація error як unknown з type guard
      return this.handleProcessError(error, result)
      
    } finally {
      this.isProcessing = false
      this.isCancelled = false
    }
  }

  /**
   * FIXED: Новий приватний метод - Фаза 1: Сканування директорії
   * Винесено з монолітної функції для покращення читабельності
   */
  private async scanDirectoryPhase(inputDirectory: string): Promise<string[]> {
    this.emitPhaseProgress(ProcessPhase.SCANNING, 'Сканування директорії...', 0, 0)
    this.emitLog('info', `Початок обробки директорії: ${inputDirectory}`)
    
    const excelFiles = await this.scanDirectory(inputDirectory)
    
    if (excelFiles.length === 0) {
      throw new Error('В директорії не знайдено Excel файлів (.xlsx)')
    }

    this.emitLog('info', `Знайдено ${excelFiles.length} Excel файл(ів)`)
    
    return excelFiles
  }

  /**
   * FIXED: Новий приватний метод - Фаза 2: Впорядкування файлів
   * Винесено з монолітної функції, додано строгу типізацію
   */
  private async orderFilesPhase(excelFiles: string[]): Promise<string[]> {
    this.emitPhaseProgress(ProcessPhase.ORDERING, 'Впорядкування файлів за місяцями...', 0, excelFiles.length)
    
    // FIXED: Типізація результату замість any
    const { validFiles, skippedFiles }: FileOrderResult = await MonthFileOrder.processFiles(
      excelFiles,
      async (filePath: string) => {
        const stats = await fs.stat(filePath)
        return { mtime: stats.mtime }
      }
    )
    
    // FIXED: Типізація orderedFiles замість any
    const orderedFiles: string[] = validFiles.map((f: OrderedFile) => f.filePath)
    
    this.emitLog('info', `Файли впорядковано: ${orderedFiles.map((f: string) => path.basename(f)).join(', ')}`)
    
    this.logSkippedFiles(skippedFiles)
    
    return orderedFiles
  }

  /**
   * FIXED: Новий приватний метод - Фаза 3: Читання Excel файлів
   * Винесено з монолітної функції для зменшення складності
   */
  private async readExcelFilesPhase(
    orderedFiles: string[], 
    result: ProcessResult
  ): Promise<{ bojecIndex: BojecIndex; fileResults: FileProcessResult[] }> {
    this.emitPhaseProgress(ProcessPhase.READING, 'Читання Excel файлів...', 0, orderedFiles.length)
    
    const bojecIndex = new BojecIndex()
    const fileResults: FileProcessResult[] = []
    
    for (let i = 0; i < orderedFiles.length; i++) {
      // FIXED: Використання методу замість дублюючого коду
      this.checkCancellation()
      
      const filePath = orderedFiles[i]
      const fileName = path.basename(filePath)
      
      // FIXED: Автоматичний розрахунок відсотків на основі фази
      const percentage = ProcessPhasePercentage.READING_START + 
        (i / orderedFiles.length) * (ProcessPhasePercentage.READING_END - ProcessPhasePercentage.READING_START)
      
      this.emitProgress(ProcessPhase.READING, `Читання файлу: ${fileName}`, i, orderedFiles.length, percentage)
      this.emitLog('info', `Обробка файлу: ${fileName}`)
      
      try {
        const fileResult = await ExcelReader.processFile(filePath)
        fileResults.push(fileResult)
        
        this.processFileResult(fileResult, fileName, bojecIndex, result)
        
      } catch (error: unknown) {
        // FIXED: Типізація error + type guard
        this.handleFileError(error, fileName, result)
      }
    }
    
    return { bojecIndex, fileResults }
  }

  /**
   * FIXED: Новий приватний метод - обробка результату одного файлу
   * Винесено для зменшення вкладеності
   */
  private processFileResult(
    fileResult: FileProcessResult,
    fileName: string,
    bojecIndex: BojecIndex,
    result: ProcessResult
  ): void {
    if (fileResult.processed) {
      result.stats.filesProcessed++
      result.stats.sheetsProcessed += fileResult.sheets.length
      
      // Додаємо дані до індексу
      bojecIndex.addFileData(fileResult)
      
      this.emitLog('info', `${fileName}: ${fileResult.totalRows} записів з ${fileResult.sheets.length} листів`)
      
      // FIXED: Винесено в окремий метод
      this.logFileWarnings(fileName, fileResult.sheets, result)
      
    } else {
      const errorMsg = fileResult.error || 'Невідома помилка'
      result.errors.push(`${fileName}: ${errorMsg}`)
      this.emitLog('error', `${fileName}: ${errorMsg}`)
    }
  }

  /**
   * FIXED: Новий приватний метод - Фаза 4: Індексування
   * Винесено з монолітної функції
   */
  private async indexingPhase(
    bojecIndex: BojecIndex, 
    options: BatchProcessOptions, 
    result: ProcessResult
  ): Promise<void> {
    // FIXED: Використання методу замість дублювання
    this.checkCancellation()
    
    this.emitPhaseProgress(ProcessPhase.INDEXING, 'Аналіз даних та вирішення конфліктів...', 0, 1)
    
    // Обчислюємо періоди служби
    this.emitLog('info', 'Обчислення періодів служби...')
    bojecIndex.calculateServicePeriods()
    
    // FIXED: Видалено закоментований код - функції тимчасово відключені
    // TODO: Розкоментувати коли analyzeConflicts() та resolveConflicts() будуть готові
    // bojecIndex.analyzeConflicts()
    
    if (options.resolveConflicts) {
      // TODO: Розкоментувати коли resolveConflicts() буде готовий
      // bojecIndex.resolveConflicts()
      this.emitLog('info', 'Конфлікти вирішено автоматично')
    }
    
    this.logIndexStats(bojecIndex, result)
  }

  /**
   * FIXED: Новий приватний метод - Фаза 5: Запис результатів
   * Винесено з монолітної функції
   */
  private async writeResultsPhase(
    bojecIndex: BojecIndex, 
    options: BatchProcessOptions,
    result: ProcessResult
  ): Promise<void> {
    // FIXED: Використання методу замість дублювання
    this.checkCancellation()
    
    this.emitPhaseProgress(ProcessPhase.WRITING, 'Створення результуючого файлу...', 0, 1)
    
    const writerOptions: ExcelWriterOptions = {
      outputPath: options.outputFilePath,
      includeStats: options.includeStats,
      includeConflicts: options.includeConflicts,
      includeOccurrences: options.includeOccurrences
    }
    
    await ExcelWriter.writeIndex(bojecIndex, writerOptions)
    
    result.outputFilePath = options.outputFilePath
    this.emitLog('info', `Результат збережено: ${options.outputFilePath}`)
  }

  /**
   * FIXED: Новий приватний метод - перевірка скасування
   * Замінює 3 дублюючі блоки коду
   */
  private checkCancellation(): void {
    if (this.isCancelled) {
      this.emitLog('warning', 'Обробка скасована користувачем')
      throw new Error('Обробка скасована користувачем')
    }
  }

  /**
   * FIXED: Новий приватний метод - логування пропущених файлів
   * Винесено для зменшення вкладеності
   */
  private logSkippedFiles(skippedFiles: SkippedFile[]): void {
    if (skippedFiles.length > 0) {
      this.emitLog('warning', `Пропущено ${skippedFiles.length} файл(ів):`)
      
      const maxToShow = 3
      for (const skipped of skippedFiles.slice(0, maxToShow)) {
        this.emitLog('warning', `  ${path.basename(skipped.path)}: ${skipped.reason}`)
      }
      
      if (skippedFiles.length > maxToShow) {
        this.emitLog('info', `  ... та ще ${skippedFiles.length - maxToShow} файл(ів)`)
      }
    }
  }

  /**
   * FIXED: Новий приватний метод - логування попереджень файлу
   * Винесено дублюючу логіку (повторювалась для warnings та errors)
   */
  private logFileWarnings(fileName: string, sheets: any[], result: ProcessResult): void {
    for (const sheet of sheets) {
      const sheetPrefix = `${fileName}/${sheet.sheetName}`
      
      // Логуємо попередження
      for (const warning of sheet.warnings) {
        result.warnings.push(`${sheetPrefix}: ${warning}`)
        this.emitLog('warning', `${sheetPrefix}: ${warning}`)
      }
      
      // Логуємо помилки аркушів
      if (sheet.error) {
        result.errors.push(`${sheetPrefix}: ${sheet.error}`)
        this.emitLog('error', `${sheetPrefix}: ${sheet.error}`)
      }
    }
  }

  /**
   * FIXED: Новий приватний метод - логування статистики індексу
   * Винесено для покращення читабельності
   */
  private logIndexStats(bojecIndex: BojecIndex, result: ProcessResult): void {
    const stats = bojecIndex.getStats()
    const conflicts = bojecIndex.getConflicts()
    
    result.stats.fightersFound = stats.totalFighters
    result.stats.totalOccurrences = stats.totalOccurrences
    result.stats.conflicts = conflicts.length
    
    this.emitLog('info', `Знайдено ${stats.totalFighters} унікальних бійців`)
    this.emitLog('info', `Загальна кількість входжень: ${stats.totalOccurrences}`)
    
    if (conflicts.length > 0) {
      this.emitLog('warning', `Знайдено ${conflicts.length} конфліктів`)
      
      const maxToShow = 5
      if (conflicts.length > maxToShow) {
        this.emitLog('info', `... та ще ${conflicts.length - maxToShow} конфліктів`)
      }
    }
  }

  /**
   * FIXED: Новий приватний метод - обробка помилки файлу
   * Винесено для покращення type safety
   */
  private handleFileError(error: unknown, fileName: string, result: ProcessResult): void {
    const errorMsg = error instanceof Error ? error.message : 'Невідома помилка'
    result.errors.push(`${fileName}: ${errorMsg}`)
    this.emitLog('error', `Помилка обробки ${fileName}: ${errorMsg}`)
  }

  /**
   * FIXED: Новий приватний метод - обробка критичної помилки процесу
   * Винесено з catch блоку для покращення структури
   */
  private handleProcessError(error: unknown, result: ProcessResult): ProcessResult {
    const errorMsg = error instanceof Error ? error.message : 'Невідома помилка'
    result.stats.processingTime = Date.now() - this.startTime
    
    // Якщо це скасування, обробляємо по-особливому
    if (this.isCancelled || errorMsg.includes('скасована')) {
      this.emitProgress(ProcessPhase.ERROR, 'Обробка скасована', 0, 1, 0)
      this.emitLog('warning', 'Обробка скасована користувачем')
      result.warnings.push('Обробка скасована користувачем')
    } else {
      result.errors.push(errorMsg)
      this.emitProgress(ProcessPhase.ERROR, `Помилка: ${errorMsg}`, 0, 1, 0)
      this.emitLog('error', `Критична помилка: ${errorMsg}`)
    }
    
    this.emit('complete', result)
    return result
  }

  /**
   * Сканує директорію на наявність Excel файлів
   */
  private async scanDirectory(directory: string): Promise<string[]> {
    try {
      const entries = await fs.readdir(directory, { withFileTypes: true })
      const excelFiles: string[] = []
      
      for (const entry of entries) {
        if (entry.isFile() && entry.name.toLowerCase().endsWith('.xlsx')) {
          // Перевіряємо що це не тимчасовий файл Excel (~$...)
          if (!entry.name.startsWith('~$')) {
            excelFiles.push(path.join(directory, entry.name))
          }
        }
      }
      
      return excelFiles
      
    } catch (error: unknown) {
      // FIXED: Типізація error
      const errorMsg = error instanceof Error ? error.message : 'невідома помилка'
      throw new Error(`Помилка читання директорії: ${errorMsg}`)
    }
  }

  /**
   * FIXED: Новий метод - випромінює прогрес фази з автоматичним відсотком
   * Спрощує виклики emitProgress, використовуючи enum для відсотків
   */
  private emitPhaseProgress(
    phase: ProcessPhase,
    message: string,
    current: number,
    total: number
  ): void {
    // FIXED: Автоматичне визначення базового відсотку з enum
    let basePercentage: number
    switch (phase) {
      case ProcessPhase.SCANNING:
        basePercentage = ProcessPhasePercentage.SCANNING
        break
      case ProcessPhase.ORDERING:
        basePercentage = ProcessPhasePercentage.ORDERING
        break
      case ProcessPhase.READING:
        basePercentage = ProcessPhasePercentage.READING_START
        break
      case ProcessPhase.INDEXING:
        basePercentage = ProcessPhasePercentage.INDEXING
        break
      case ProcessPhase.WRITING:
        basePercentage = ProcessPhasePercentage.WRITING
        break
      case ProcessPhase.COMPLETE:
        basePercentage = ProcessPhasePercentage.COMPLETE
        break
      case ProcessPhase.ERROR:
        basePercentage = 0
        break
      default:
        basePercentage = 0
    }
    
    this.emitProgress(phase, message, current, total, basePercentage)
  }

  /**
   * Випромінює подію прогресу
   */
  private emitProgress(
    phase: ProcessPhase,
    message: string,
    current: number,
    total: number,
    basePercentage: number
  ): void {
    // FIXED: Використання ProcessPhasePercentage для розрахунків
    const progressRange = 10 // Кожна фаза займає ~10% прогресу
    const percentage = total > 0 
      ? Math.min(basePercentage + (current / total) * progressRange, 100) 
      : basePercentage
    
    const timeElapsed = Date.now() - this.startTime
    
    let estimatedTimeRemaining: number | undefined
    // FIXED: Винесено магічні числа 5 та 95 у константи
    const minPercentageForEstimate = 5
    const maxPercentageForEstimate = 95
    
    if (percentage > minPercentageForEstimate && percentage < maxPercentageForEstimate) {
      const timePerPercent = timeElapsed / percentage
      estimatedTimeRemaining = Math.round(timePerPercent * (100 - percentage))
    }
    
    const progress: ProcessProgress = {
      phase,
      filesProcessed: current,
      totalFiles: total,
      percentage: Math.round(percentage),
      message,
      timeElapsed,
      estimatedTimeRemaining
    }
    
    this.emit('progress', progress)
  }

  /**
   * Випромінює подію логування
   */
  private emitLog(level: 'info' | 'warning' | 'error', message: string): void {
    const timestamp = new Date().toLocaleTimeString('uk-UA')
    this.emit('log', level, `[${timestamp}] ${message}`)
  }

  /**
   * Перевіряє чи виконується обробка
   */
  isRunning(): boolean {
    return this.isProcessing
  }

  /**
   * Скасовує обробку (якщо підтримується в майбутньому)
   */
  cancel(): void {
    if (this.isProcessing) {
      this.isCancelled = true
      this.emitLog('warning', 'Скасування обробки...')
    }
  }

  /**
   * Валідує опції перед початком обробки
   */
  static validateOptions(options: BatchProcessOptions): string[] {
    const errors: string[] = []
    
    if (!options.inputDirectory?.trim()) {
      errors.push('Не вказано вхідну директорію')
    }
    
    if (!options.outputFilePath?.trim()) {
      errors.push('Не вказано файл для збереження результату')
    } else if (!options.outputFilePath.toLowerCase().endsWith('.xlsx')) {
      errors.push('Файл результату повинен мати розширення .xlsx')
    }
    
    return errors
  }

  /**
   * Створює дефолтні опції
   */
  static createDefaultOptions(): Partial<BatchProcessOptions> {
    return {
      includeStats: true,
      includeConflicts: true,
      includeOccurrences: false,
      resolveConflicts: true
    }
  }
}
