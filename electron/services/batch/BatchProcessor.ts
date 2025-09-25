/**
 * BatchProcessor.ts - Основний сервіс пакетної обробки Excel файлів
 * Інтегрує всі компоненти та забезпечує прогрес-трекінг і логування
 */

import { EventEmitter } from 'events'
import * as path from 'path'
import * as fs from 'fs/promises'

import { MonthFileOrder } from './MonthFileOrder'
import { ExcelReader, FileProcessResult } from './ExcelReader'
import { BojecIndex } from './BojecIndex'
import { ExcelWriter, ExcelWriterOptions } from './ExcelWriter'

export interface BatchProcessOptions {
  inputDirectory: string
  outputFilePath: string
  includeStats: boolean
  includeConflicts: boolean
  includeOccurrences: boolean
  resolveConflicts: boolean
}

export interface ProcessProgress {
  phase: 'scanning' | 'ordering' | 'reading' | 'indexing' | 'writing' | 'complete' | 'error'
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
      // Фаза 1: Сканування директорії
      this.emitProgress('scanning', 'Сканування директорії...', 0, 0, 0)
      this.emitLog('info', `Початок обробки директорії: ${options.inputDirectory}`)
      
      const excelFiles = await this.scanDirectory(options.inputDirectory)
      if (excelFiles.length === 0) {
        throw new Error('В директорії не знайдено Excel файлів (.xlsx)')
      }

      this.emitLog('info', `Знайдено ${excelFiles.length} Excel файл(ів)`)

      // Фаза 2: Впорядкування файлів
      this.emitProgress('ordering', 'Впорядкування файлів за місяцями...', 0, excelFiles.length, 10)
      
      const { validFiles, skippedFiles } = await MonthFileOrder.processFiles(
        excelFiles,
        async (filePath: string) => {
          const stats = await fs.stat(filePath)
          return { mtime: stats.mtime }
        }
      )
      
      const orderedFiles = validFiles.map((f: any) => f.filePath)
      
      this.emitLog('info', `Файли впорядковано: ${orderedFiles.map((f: string) => path.basename(f)).join(', ')}`)
      
      if (skippedFiles.length > 0) {
        this.emitLog('warning', `Пропущено ${skippedFiles.length} файл(ів):`)
        for (const skipped of skippedFiles.slice(0, 3)) {
          this.emitLog('warning', `  ${path.basename(skipped.path)}: ${skipped.reason}`)
        }
        if (skippedFiles.length > 3) {
          this.emitLog('info', `  ... та ще ${skippedFiles.length - 3} файл(ів)`)
        }
      }

      // Фаза 3: Читання Excel файлів
      this.emitProgress('reading', 'Читання Excel файлів...', 0, orderedFiles.length, 20)
      
      const bojecIndex = new BojecIndex()
      const fileResults: FileProcessResult[] = []
      
      for (let i = 0; i < orderedFiles.length; i++) {
        // Перевіряємо скасування
        if (this.isCancelled) {
          this.emitLog('warning', 'Обробка скасована користувачем')
          throw new Error('Обробка скасована користувачем')
        }
        
        const filePath = orderedFiles[i]
        const fileName = path.basename(filePath)
        
        this.emitProgress('reading', `Читання файлу: ${fileName}`, i, orderedFiles.length, 20 + (i / orderedFiles.length) * 50)
        this.emitLog('info', `Обробка файлу: ${fileName}`)
        
        try {
          const fileResult = await ExcelReader.processFile(filePath)
          fileResults.push(fileResult)
          
          if (fileResult.processed) {
            result.stats.filesProcessed++
            result.stats.sheetsProcessed += fileResult.sheets.length
            
            // Додаємо дані до індексу
            bojecIndex.addFileData(fileResult)
            
            this.emitLog('info', `${fileName}: ${fileResult.totalRows} записів з ${fileResult.sheets.length} листів`)
            
            // Логуємо попередження
            for (const sheet of fileResult.sheets) {
              for (const warning of sheet.warnings) {
                result.warnings.push(`${fileName}/${sheet.sheetName}: ${warning}`)
                this.emitLog('warning', `${fileName}/${sheet.sheetName}: ${warning}`)
              }
              if (sheet.error) {
                result.errors.push(`${fileName}/${sheet.sheetName}: ${sheet.error}`)
                this.emitLog('error', `${fileName}/${sheet.sheetName}: ${sheet.error}`)
              }
            }
          } else {
            result.errors.push(`${fileName}: ${fileResult.error}`)
            this.emitLog('error', `${fileName}: ${fileResult.error}`)
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : 'Невідома помилка'
          result.errors.push(`${fileName}: ${errorMsg}`)
          this.emitLog('error', `Помилка обробки ${fileName}: ${errorMsg}`)
        }
      }

      // Фаза 4: Індексування та аналіз конфліктів
      // Перевіряємо скасування перед індексуванням
      if (this.isCancelled) {
        this.emitLog('warning', 'Обробка скасована користувачем')
        throw new Error('Обробка скасована користувачем')
      }
      
      this.emitProgress('indexing', 'Аналіз даних та вирішення конфліктів...', 0, 1, 70)
      
      // Обчислюємо періоди служби
      this.emitLog('info', 'Обчислення періодів служби...')
      bojecIndex.calculateServicePeriods()
      
      // bojecIndex.analyzeConflicts() // Тимчасово відключено
      
      if (options.resolveConflicts) {
        // bojecIndex.resolveConflicts() // Тимчасово відключено
        this.emitLog('info', 'Конфлікти вирішено автоматично')
      }
      
      const stats = bojecIndex.getStats()
      const conflicts = bojecIndex.getConflicts()
      
      result.stats.fightersFound = stats.totalFighters
      result.stats.totalOccurrences = stats.totalOccurrences
      result.stats.conflicts = conflicts.length
      
      this.emitLog('info', `Знайдено ${stats.totalFighters} унікальних бійців`)
      this.emitLog('info', `Загальна кількість входжень: ${stats.totalOccurrences}`)
      
      if (conflicts.length > 0) {
        this.emitLog('warning', `Знайдено ${conflicts.length} конфліктів`)
        // Conflict processing temporarily disabled
        if (conflicts.length > 5) {
          this.emitLog('info', `... та ще ${conflicts.length - 5} конфліктів`)
        }
      }

      // Фаза 5: Запис результату
      // Перевіряємо скасування перед записом
      if (this.isCancelled) {
        this.emitLog('warning', 'Обробка скасована користувачем')
        throw new Error('Обробка скасована користувачем')
      }
      
      this.emitProgress('writing', 'Створення результуючого файлу...', 0, 1, 90)
      
      const writerOptions: ExcelWriterOptions = {
        outputPath: options.outputFilePath,
        includeStats: options.includeStats,
        includeConflicts: options.includeConflicts,
        includeOccurrences: options.includeOccurrences
      }
      
      await ExcelWriter.writeIndex(bojecIndex, writerOptions)
      
      result.outputFilePath = options.outputFilePath
      this.emitLog('info', `Результат збережено: ${options.outputFilePath}`)

      // Завершення
      result.success = true
      result.stats.processingTime = Date.now() - this.startTime
      
      this.emitProgress('complete', 'Обробка завершена успішно', 1, 1, 100)
      this.emitLog('info', `Обробка завершена за ${Math.round(result.stats.processingTime / 1000)} секунд`)
      
      this.emit('complete', result)
      return result

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Невідома помилка'
      result.stats.processingTime = Date.now() - this.startTime
      
      // Якщо це скасування, обробляємо по-особливому
      if (this.isCancelled || errorMsg.includes('скасована')) {
        this.emitProgress('error', 'Обробка скасована', 0, 1, 0)
        this.emitLog('warning', 'Обробка скасована користувачем')
        result.warnings.push('Обробка скасована користувачем')
      } else {
        result.errors.push(errorMsg)
        this.emitProgress('error', `Помилка: ${errorMsg}`, 0, 1, 0)
        this.emitLog('error', `Критична помилка: ${errorMsg}`)
      }
      
      this.emit('complete', result)
      return result
      
    } finally {
      this.isProcessing = false
      this.isCancelled = false
    }
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
      
    } catch (error) {
      throw new Error(`Помилка читання директорії: ${error instanceof Error ? error.message : 'невідома помилка'}`)
    }
  }

  /**
   * Випромінює подію прогресу
   */
  private emitProgress(
    phase: ProcessProgress['phase'],
    message: string,
    current: number,
    total: number,
    basePercentage: number
  ): void {
    const percentage = total > 0 ? Math.min(basePercentage + (current / total) * 10, 100) : basePercentage
    const timeElapsed = Date.now() - this.startTime
    
    let estimatedTimeRemaining: number | undefined
    if (percentage > 5 && percentage < 95) {
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