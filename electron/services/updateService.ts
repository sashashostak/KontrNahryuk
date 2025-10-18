/**
 * UpdateService - Advanced Auto-Update System for KontrNahryuk
 * Based on TZ_Auto_Update_IMPROVED.txt
 * 
 * Features:
 * - Real-time progress bar with download speed
 * - Automatic ZIP extraction
 * - Automatic backup system
 * - Automatic file replacement
 * - Auto-restart after update
 * 
 * @version 3.0.0 - Advanced with Progress Bar & Auto-Install
 */

import { EventEmitter } from 'events'
import { app, shell } from 'electron'
import * as path from 'path'
import * as fs from 'fs'
import * as os from 'os'
import AdmZip from 'adm-zip'
import { spawn } from 'child_process'


// ============================================================================
// TYPES & INTERFACES
// ============================================================================

/**
 * Інформація про перевірку оновлень
 */
export interface UpdateInfo {
  hasUpdate: boolean          // Чи є доступне оновлення
  latestVersion: string        // Остання версія на GitHub
  currentVersion: string       // Поточна версія додатку
  releaseInfo: any | null      // Повна інформація про реліз з GitHub API
  error: string | null         // Помилка (якщо є)
}

/**
 * Прогрес завантаження оновлення
 */
export interface DownloadProgress {
  percent: number              // Відсоток завантаження (0-100)
  downloadedBytes: number      // Скільки байтів завантажено
  totalBytes: number           // Загальний розмір файлу
  bytesPerSecond: number       // Швидкість завантаження
}

/**
 * Результат завантаження оновлення
 */
export interface DownloadResult {
  success: boolean             // Чи успішно завантажено
  path: string | false         // Шлях до файлу або false
  error?: string               // Помилка (якщо є)
}

// ============================================================================
// UPDATE SERVICE CLASS
// ============================================================================

class UpdateService extends EventEmitter {
  private currentVersion: string
  private updateCheckInProgress: boolean = false
  private downloadInProgress: boolean = false
  private updateBasePath: string
  private backupPath: string
  private readonly GITHUB_REPO = 'sashashostak/KontrNahryuk'

  constructor() {
    super()
    this.currentVersion = app.getVersion()
    
    // Папки для оновлень
    const localAppData = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local')
    this.updateBasePath = path.join(localAppData, 'KontrNahryuk', 'Updates')
    this.backupPath = path.join(localAppData, 'KontrNahryuk', 'Backup')
    
    // Створити папки якщо не існують
    this.ensureDirectories()
    
    this.log(`Ініціалізовано. Поточна версія: ${this.currentVersion}`)
    this.log(`Папка оновлень: ${this.updateBasePath}`)
  }

  /**
   * Створити необхідні директорії
   */
  private ensureDirectories(): void {
    if (!fs.existsSync(this.updateBasePath)) {
      fs.mkdirSync(this.updateBasePath, { recursive: true })
    }
    if (!fs.existsSync(this.backupPath)) {
      fs.mkdirSync(this.backupPath, { recursive: true })
    }
  }

  // ==========================================================================
  // PUBLIC METHODS
  // ==========================================================================

  /**
   * Отримати поточну версію додатку
   */
  getCurrentVersion(): string {
    return this.currentVersion
  }

  /**
   * Перевірити наявність оновлень через GitHub Releases API
   * 
   * @returns {Promise<UpdateInfo>} Інформація про оновлення
   */
  async checkForUpdates(): Promise<UpdateInfo> {
    // Prevent multiple simultaneous checks
    if (this.updateCheckInProgress) {
      this.log('⚠️ Перевірка оновлень вже виконується')
      return {
        hasUpdate: false,
        latestVersion: this.currentVersion,
        currentVersion: this.currentVersion,
        releaseInfo: null,
        error: 'Перевірка вже виконується'
      }
    }

    this.updateCheckInProgress = true

    try {
      this.log('🔍 Перевірка оновлень на GitHub...')

      // Fetch latest release from GitHub API
      const url = `https://api.github.com/repos/${this.GITHUB_REPO}/releases/latest`
      
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'User-Agent': `KontrNahryuk/${this.currentVersion}`,
          'Accept': 'application/vnd.github.v3+json'
        }
      })

      if (!response.ok) {
        throw new Error(`GitHub API помилка: ${response.status} ${response.statusText}`)
      }

      const release = await response.json()
      const latestVersion = release.tag_name || release.name || 'unknown'
      const hasUpdate = this.isNewerVersion(latestVersion, this.currentVersion)

      this.log(`📦 Поточна версія: ${this.currentVersion}`)
      this.log(`📦 Остання версія: ${latestVersion}`)
      this.log(`${hasUpdate ? '✅ Доступне оновлення!' : '✓ Програма актуальна'}`)

      this.updateCheckInProgress = false

      return {
        hasUpdate,
        latestVersion,
        currentVersion: this.currentVersion,
        releaseInfo: hasUpdate ? release : null,
        error: null
      }

    } catch (error) {
      this.log(`❌ Помилка перевірки оновлень: ${error instanceof Error ? error.message : String(error)}`)
      this.updateCheckInProgress = false

      return {
        hasUpdate: false,
        latestVersion: this.currentVersion,
        currentVersion: this.currentVersion,
        releaseInfo: null,
        error: error instanceof Error ? error.message : 'Невідома помилка'
      }
    }
  }

  /**
   * Завантажити та встановити оновлення з прогрес-баром та автоустановкою
   * 
   * @param {UpdateInfo} updateInfo - Інформація про оновлення
   * @returns {Promise<DownloadResult>} Результат завантаження та установки
   */
  async downloadUpdate(updateInfo: UpdateInfo): Promise<DownloadResult> {
    if (this.downloadInProgress) {
      return {
        success: false,
        path: false,
        error: 'Завантаження вже виконується'
      }
    }

    this.downloadInProgress = true

    try {
      // Крок 1: Знайти portable файл
      const release = updateInfo.releaseInfo
      if (!release) {
        throw new Error('Інформація про реліз відсутня')
      }

      const portableAsset = release.assets?.find((asset: any) =>
        asset.name.toLowerCase().includes('portable') && 
        asset.name.endsWith('.zip')
      )

      if (!portableAsset) {
        this.log('⚠️ Portable файл не знайдено, відкриваю GitHub')
        shell.openExternal(release.html_url)
        throw new Error('Автоматичне оновлення недоступне')
      }

      this.log(`📥 Завантаження файлу: ${portableAsset.name}`)
      
      // Крок 2: Завантажити з прогресом
      const downloadPath = path.join(this.updateBasePath, portableAsset.name)
      await this.downloadWithProgress(portableAsset.browser_download_url, downloadPath)

      this.log(`✅ Файл завантажено: ${downloadPath}`)

      // Крок 3: Розпакувати
      this.emit('status', { message: 'Розпакування файлів...' })
      const extractPath = path.join(this.updateBasePath, 'extracted')
      await this.extractZip(downloadPath, extractPath)

      this.log(`✅ Файли розпаковано: ${extractPath}`)

      // Крок 4: Створити backup поточної версії
      this.emit('status', { message: 'Створення резервної копії...' })
      await this.createBackup()

      this.log(`✅ Backup створено`)

      // Крок 5: Замінити файли
      this.emit('status', { message: 'Оновлення файлів...' })
      await this.replaceFiles(extractPath)

      this.log(`✅ Файли оновлено`)

      // Крок 6: Очистити тимчасові файли
      await this.cleanupTempFiles(downloadPath, extractPath)

      // Крок 7: Перезапустити додаток
      this.emit('status', { message: 'Перезапуск додатку...' })
      setTimeout(() => {
        app.relaunch()
        app.exit(0)
      }, 1000)

      return {
        success: true,
        path: extractPath
      }

    } catch (error) {
      console.error('[UpdateService] Помилка оновлення:', error)
      this.emit('error', { 
        message: error instanceof Error ? error.message : 'Невідома помилка' 
      })
      return {
        success: false,
        path: false,
        error: error instanceof Error ? error.message : 'Невідома помилка'
      }
    } finally {
      this.downloadInProgress = false
    }
  }

  // ==========================================================================
  // PRIVATE HELPER METHODS
  // ==========================================================================

  /**
   * Завантажити файл з відстеженням прогресу
   */
  private async downloadWithProgress(url: string, savePath: string): Promise<void> {
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'User-Agent': `KontrNahryuk/${this.currentVersion}`,
        'Accept': 'application/octet-stream'
      }
    })

    if (!response.ok) {
      throw new Error(`Помилка завантаження: ${response.status}`)
    }

    const totalBytes = parseInt(response.headers.get('content-length') || '0', 10)
    
    if (!response.body) {
      throw new Error('Response body відсутній')
    }

    const reader = response.body.getReader()
    const chunks: Uint8Array[] = []
    let downloadedBytes = 0
    let startTime = Date.now()
    let lastEmitTime = Date.now()

    this.log(`📦 Розмір файлу: ${this.formatBytes(totalBytes)}`)

    while (true) {
      const { done, value } = await reader.read()

      if (done) break

      chunks.push(value)
      downloadedBytes += value.length

      // Оновлювати прогрес кожні 100ms
      const now = Date.now()
      if (now - lastEmitTime >= 100) {
        const elapsed = (now - startTime) / 1000
        const bytesPerSecond = downloadedBytes / elapsed
        const percent = totalBytes > 0 ? (downloadedBytes / totalBytes) * 100 : 0

        this.emit('download-progress', {
          percent: Math.round(percent * 100) / 100,
          downloadedBytes,
          totalBytes,
          bytesPerSecond: Math.round(bytesPerSecond)
        })

        this.log(
          `📊 Прогрес: ${percent.toFixed(1)}% ` +
          `(${this.formatBytes(downloadedBytes)} / ${this.formatBytes(totalBytes)}) ` +
          `${this.formatBytes(bytesPerSecond)}/s`
        )

        lastEmitTime = now
      }
    }

    // Фінальний прогрес
    this.emit('download-progress', {
      percent: 100,
      downloadedBytes: totalBytes,
      totalBytes,
      bytesPerSecond: 0
    })

    // Зберегти файл
    const buffer = Buffer.concat(chunks)
    fs.writeFileSync(savePath, buffer)
  }

  /**
   * Розпакувати ZIP архів
   */
  private async extractZip(zipPath: string, extractPath: string): Promise<void> {
    return new Promise((resolve, reject) => {
      try {
        // Видалити стару папку якщо існує
        if (fs.existsSync(extractPath)) {
          fs.rmSync(extractPath, { recursive: true, force: true })
        }

        // Створити нову папку
        fs.mkdirSync(extractPath, { recursive: true })

        // Розпакувати
        const zip = new AdmZip(zipPath)
        zip.extractAllTo(extractPath, true)

        this.log(`📂 Розпаковано файлів: ${zip.getEntries().length}`)

        resolve()
      } catch (error) {
        reject(new Error(`Помилка розпакування: ${error instanceof Error ? error.message : 'Unknown'}`))
      }
    })
  }

  /**
   * Створити backup поточної версії
   */
  private async createBackup(): Promise<void> {
    try {
      const currentExePath = app.getPath('exe')
      const currentDir = path.dirname(currentExePath)
      const backupDir = path.join(this.backupPath, `v${this.currentVersion}_${Date.now()}`)

      if (!fs.existsSync(backupDir)) {
        fs.mkdirSync(backupDir, { recursive: true })
      }

      // Копіювати .exe файл
      const exeFileName = path.basename(currentExePath)
      const backupExePath = path.join(backupDir, exeFileName)
      fs.copyFileSync(currentExePath, backupExePath)

      // Копіювати resources якщо є
      const resourcesPath = path.join(currentDir, 'resources')
      if (fs.existsSync(resourcesPath)) {
        const backupResourcesPath = path.join(backupDir, 'resources')
        this.copyDirectory(resourcesPath, backupResourcesPath)
      }

      this.log(`💾 Backup створено: ${backupDir}`)

      // Видалити старі backups (залишити тільки останні 3)
      this.cleanupOldBackups(3)

    } catch (error) {
      console.warn('[UpdateService] Помилка створення backup:', error)
      // Не блокуємо оновлення якщо backup не створився
    }
  }

  /**
   * Замінити файли новою версією
   */
  private async replaceFiles(extractPath: string): Promise<void> {
    const currentExePath = app.getPath('exe')
    const currentDir = path.dirname(currentExePath)

    // Знайти підпапку з розпакованими файлами
    let actualExtractPath = extractPath
    const entries = fs.readdirSync(extractPath)
    
    // Якщо є тільки одна папка, файли можуть бути всередині
    if (entries.length === 1 && fs.statSync(path.join(extractPath, entries[0])).isDirectory()) {
      actualExtractPath = path.join(extractPath, entries[0])
    }

    // Знайти .exe файл рекурсивно
    const findExeFile = (dir: string): string | null => {
      const files = fs.readdirSync(dir)
      for (const file of files) {
        const filePath = path.join(dir, file)
        if (file.endsWith('.exe')) {
          return filePath
        }
        if (fs.statSync(filePath).isDirectory()) {
          const found = findExeFile(filePath)
          if (found) return found
        }
      }
      return null
    }

    const newExePath = findExeFile(actualExtractPath)

    if (!newExePath) {
      throw new Error('Не знайдено .exe файл в оновленні')
    }

    const tempExePath = path.join(currentDir, `${path.basename(currentExePath)}.new`)

    // Скопіювати новий .exe як тимчасовий
    fs.copyFileSync(newExePath, tempExePath)

    // Створити .bat скрипт для заміни після закриття
    const batScript = `
@echo off
echo Оновлення KontrNahryuk...
timeout /t 2 /nobreak > nul

:retry
del /f /q "${currentExePath}"
if exist "${currentExePath}" (
  timeout /t 1 /nobreak > nul
  goto retry
)

move /y "${tempExePath}" "${currentExePath}"

echo Оновлення завершено!
start "" "${currentExePath}"
del "%~f0"
`.trim()

    const batPath = path.join(currentDir, 'update.bat')
    fs.writeFileSync(batPath, batScript)

    // Запустити .bat скрипт
    spawn('cmd.exe', ['/c', batPath], {
      detached: true,
      stdio: 'ignore'
    }).unref()
  }

  /**
   * Очистити тимчасові файли
   */
  private async cleanupTempFiles(zipPath: string, extractPath: string): Promise<void> {
    try {
      if (fs.existsSync(zipPath)) {
        fs.unlinkSync(zipPath)
      }
      if (fs.existsSync(extractPath)) {
        fs.rmSync(extractPath, { recursive: true, force: true })
      }
    } catch (error) {
      console.warn('[UpdateService] Помилка очищення:', error)
    }
  }

  /**
   * Видалити старі backups
   */
  private cleanupOldBackups(keepCount: number): void {
    try {
      const backups = fs.readdirSync(this.backupPath)
        .filter(name => name.startsWith('v'))
        .map(name => ({
          name,
          path: path.join(this.backupPath, name),
          time: fs.statSync(path.join(this.backupPath, name)).mtime.getTime()
        }))
        .sort((a, b) => b.time - a.time)

      backups.slice(keepCount).forEach(backup => {
        fs.rmSync(backup.path, { recursive: true, force: true })
        this.log(`🗑️ Видалено старий backup: ${backup.name}`)
      })
    } catch (error) {
      console.warn('[UpdateService] Помилка очищення backups:', error)
    }
  }

  /**
   * Копіювати директорію рекурсивно
   */
  private copyDirectory(src: string, dest: string): void {
    if (!fs.existsSync(dest)) {
      fs.mkdirSync(dest, { recursive: true })
    }

    const entries = fs.readdirSync(src, { withFileTypes: true })

    for (const entry of entries) {
      const srcPath = path.join(src, entry.name)
      const destPath = path.join(dest, entry.name)

      if (entry.isDirectory()) {
        this.copyDirectory(srcPath, destPath)
      } else {
        fs.copyFileSync(srcPath, destPath)
      }
    }
  }

  /**
   * Форматувати байти в читабельний вигляд
   */
  private formatBytes(bytes: number): string {
    if (bytes === 0) return '0 Bytes'

    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))

    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i]
  }

  // ==========================================================================
  // PRIVATE METHODS
  // ==========================================================================

  /**
   * Порівняти версії (semantic versioning: major.minor.patch)
   * 
   * @param {string} remoteVersion - Віддалена версія (з GitHub)
   * @param {string} currentVersion - Поточна версія
   * @returns {boolean} true якщо remoteVersion новіша
   */
  private isNewerVersion(remoteVersion: string, currentVersion: string): boolean {
    // Видалити префікс 'v' або 'V' якщо є
    const cleanRemote = remoteVersion.replace(/^[vV]/, '').trim()
    const cleanCurrent = currentVersion.replace(/^[vV]/, '').trim()

    // Розбити на частини: major.minor.patch
    const remoteParts = cleanRemote.split('.').map(part => 
      parseInt(part.replace(/[^\d]/g, ''), 10) || 0
    )
    const currentParts = cleanCurrent.split('.').map(part => 
      parseInt(part.replace(/[^\d]/g, ''), 10) || 0
    )

    // Доповнити до 3 частин (якщо версія неповна)
    while (remoteParts.length < 3) remoteParts.push(0)
    while (currentParts.length < 3) currentParts.push(0)

    this.log(`🔢 Порівняння версій: [${remoteParts.join('.')}] vs [${currentParts.join('.')}]`)

    // Порівняти по частинах: major -> minor -> patch
    for (let i = 0; i < 3; i++) {
      if (remoteParts[i] > currentParts[i]) {
        this.log(`   ↑ Remote версія новіша на рівні ${i === 0 ? 'major' : i === 1 ? 'minor' : 'patch'}`)
        return true
      }
      if (remoteParts[i] < currentParts[i]) {
        this.log(`   ↓ Remote версія старіша`)
        return false
      }
    }

    this.log(`   = Версії однакові`)
    return false
  }

  /**
   * Логування з префіксом
   */
  private log(message: string): void {
    console.log(`[UpdateService] ${message}`)
  }
}

// ============================================================================
// SINGLETON EXPORT
// ============================================================================

export const updateService = new UpdateService()
export { UpdateService }
