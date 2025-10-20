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
      // Крок 1: Знайти patch або portable файл
      const release = updateInfo.releaseInfo
      if (!release) {
        throw new Error('Інформація про реліз відсутня')
      }

      // Спочатку шукаємо patch файл (пріоритет для економії трафіку)
      const currentVersion = app.getVersion()
      
      // Шукаємо прямий патч для поточної версії
      let patchAsset = release.assets?.find((asset: any) =>
        asset.name.toLowerCase().includes('patch') && 
        asset.name.includes(currentVersion) &&
        asset.name.endsWith('.zip')
      )

      // Якщо прямий патч не знайдено, шукаємо ланцюжок патчів
      if (!patchAsset) {
        this.log(`⚠️ Прямий патч ${currentVersion} → ${updateInfo.latestVersion} не знайдено`)
        this.log(`🔗 Шукаю ланцюжок патчів через проміжні версії...`)
        
        const patchChain = await this.findPatchChain(currentVersion, updateInfo.latestVersion)
        
        if (patchChain.length > 0) {
          this.log(`✅ Знайдено ланцюжок оновлень: ${patchChain.map(p => p.version).join(' → ')}`)
          this.log(`📦 Буде завантажено ${patchChain.length} патчів (замість portable)`)
          
          // Завантажуємо та встановлюємо патчі по черзі
          return await this.downloadAndApplyPatchChain(patchChain)
        } else {
          this.log(`⚠️ Ланцюжок патчів не знайдено`)
        }
      }

      // Якщо patch не знайдено, шукаємо portable
      const portableAsset = release.assets?.find((asset: any) =>
        asset.name.toLowerCase().includes('portable') && 
        asset.name.endsWith('.zip')
      )

      const updateAsset = patchAsset || portableAsset

      if (!updateAsset) {
        this.log('⚠️ Файли оновлення не знайдено, відкриваю GitHub')
        shell.openExternal(release.html_url)
        throw new Error('Автоматичне оновлення недоступне')
      }

      const isPatch = !!patchAsset
      this.log(`📥 Завантаження ${isPatch ? 'patch' : 'portable'} файлу: ${updateAsset.name}`)
      
      // Крок 2: Завантажити з прогресом
      const downloadPath = path.join(this.updateBasePath, updateAsset.name)
      await this.downloadWithProgress(updateAsset.browser_download_url, downloadPath)

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

    // Знайти .exe файл рекурсивно (для full portable)
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

    // Якщо знайдено .exe - це full portable, замінюємо через .bat
    if (newExePath) {
      this.log('📦 Full portable виявлено, оновлення через .bat скрипт')
      
      const tempExePath = path.join(currentDir, `${path.basename(currentExePath)}.new`)

      // Скопіювати новий .exe як тимчасовий
      fs.copyFileSync(newExePath, tempExePath)

      // Створити .bat скрипт для заміни після закриття (з підтримкою Unicode шляхів)
      const batScript = `@echo off
chcp 65001 > nul
echo Оновлення KontrNahryuk...
timeout /t 3 /nobreak > nul

:retry
del /f /q "%~dp0${path.basename(currentExePath)}"
if exist "%~dp0${path.basename(currentExePath)}" (
  timeout /t 1 /nobreak > nul
  goto retry
)

move /y "%~dp0${path.basename(tempExePath)}" "%~dp0${path.basename(currentExePath)}"
if errorlevel 1 (
  echo Помилка переміщення файлу!
  pause
  exit /b 1
)

echo Оновлення завершено! Запуск додатку...
timeout /t 1 /nobreak > nul
start "" "%~dp0${path.basename(currentExePath)}"

timeout /t 2 /nobreak > nul
del "%~f0"
exit
`

      const batPath = path.join(currentDir, 'update.bat')
      fs.writeFileSync(batPath, batScript, { encoding: 'utf8' })

      // Запустити .bat скрипт
      spawn('cmd.exe', ['/c', batPath], {
        detached: true,
        stdio: 'ignore',
        cwd: currentDir
      }).unref()
      
      // КРИТИЧНО: Закрити додаток ПІСЛЯ запуску bat
      setTimeout(() => {
        app.quit()
      }, 500)
      
    } else {
      // Це patch файл - замінюємо тільки файли в resources/app/
      this.log('🔄 Patch виявлено, оновлення файлів resources/app/')
      
      // Шлях до поточної папки resources/app
      const currentResourcesPath = path.join(currentDir, 'resources', 'app')
      
      if (!fs.existsSync(currentResourcesPath)) {
        throw new Error('Поточна папка resources/app/ не знайдена')
      }

      // Перевіряємо дві можливі структури патчу:
      // 1. Нова структура: dist/ + package.json (безпосередньо в actualExtractPath)
      // 2. Стара структура: resources/app/ (всередині actualExtractPath)
      
      const distFolder = path.join(actualExtractPath, 'dist')
      const packageJson = path.join(actualExtractPath, 'package.json')
      const resourcesPath = path.join(actualExtractPath, 'resources', 'app')
      
      // Копіювання рекурсивне
      const copyRecursive = (src: string, dest: string, basePath?: string) => {
        if (!fs.existsSync(src)) return
        
        const stat = fs.statSync(src)
        if (stat.isDirectory()) {
          if (!fs.existsSync(dest)) {
            fs.mkdirSync(dest, { recursive: true })
          }
          const files = fs.readdirSync(src)
          for (const file of files) {
            copyRecursive(path.join(src, file), path.join(dest, file), basePath || src)
          }
        } else {
          fs.copyFileSync(src, dest)
          const relativePath = basePath ? path.relative(basePath, src) : path.basename(src)
          this.log(`  ✓ ${relativePath}`)
        }
      }
      
      // Нова структура: dist/ + package.json
      if (fs.existsSync(distFolder) && fs.existsSync(packageJson)) {
        this.log('📦 Patch нової структури (dist/ + package.json)')
        
        // Копіюємо dist/
        const targetDist = path.join(currentResourcesPath, 'dist')
        copyRecursive(distFolder, targetDist, distFolder)
        
        // Копіюємо package.json
        fs.copyFileSync(packageJson, path.join(currentResourcesPath, 'package.json'))
        this.log('  ✓ package.json')
        
        this.log('✅ Patch файли замінено (нова структура), перезапуск...')
        
      // Стара структура: resources/app/
      } else if (fs.existsSync(resourcesPath)) {
        this.log('📦 Patch старої структури (resources/app/)')
        
        copyRecursive(resourcesPath, currentResourcesPath, resourcesPath)
        
        this.log('✅ Patch файли замінено (стара структура), перезапуск...')
        
      } else {
        throw new Error('Patch не містить ні dist/ + package.json, ні resources/app/')
      }
      
      // Для patch перезапускаємо просто через relaunch
      app.relaunch()
      app.exit(0)
    }
  }  /**
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
   * Знайти ланцюжок патчів від поточної версії до цільової
   * 
   * @param {string} fromVersion - Початкова версія (поточна)
   * @param {string} toVersion - Цільова версія (нова)
   * @returns {Promise<Array<{version: string, release: any, patchAsset: any}>>} Ланцюжок патчів
   */
  private async findPatchChain(fromVersion: string, toVersion: string): Promise<Array<{version: string, release: any, patchAsset: any}>> {
    try {
      // Отримати всі релізи з GitHub
      const url = `https://api.github.com/repos/${this.GITHUB_REPO}/releases`
      const response = await fetch(url, {
        method: 'GET',
        headers: {
          'User-Agent': `KontrNahryuk/${this.currentVersion}`,
          'Accept': 'application/vnd.github.v3+json'
        }
      })

      if (!response.ok) {
        throw new Error(`GitHub API помилка: ${response.status}`)
      }

      const releases = await response.json()
      
      // Парсимо версії
      const cleanFrom = fromVersion.replace(/^[vV]/, '')
      const cleanTo = toVersion.replace(/^[vV]/, '')
      
      const fromParts = cleanFrom.split('.').map(p => parseInt(p, 10) || 0)
      const toParts = cleanTo.split('.').map(p => parseInt(p, 10) || 0)
      
      // Знаходимо всі проміжні версії
      const chain: Array<{version: string, release: any, patchAsset: any}> = []
      let currentVer = cleanFrom
      
      // Генеруємо послідовність версій (тільки patch increment)
      // Наприклад: 1.4.5 -> 1.4.6 -> 1.4.7 -> 1.4.8 -> 1.4.9
      const [major, minor, patchStart] = fromParts
      const patchEnd = toParts[2]
      
      for (let patch = patchStart + 1; patch <= patchEnd; patch++) {
        const targetVersion = `${major}.${minor}.${patch}`
        
        // Знаходимо реліз для цієї версії
        const release = releases.find((r: any) => {
          const releaseVer = (r.tag_name || r.name || '').replace(/^[vV]/, '')
          return releaseVer === targetVersion
        })
        
        if (!release) {
          this.log(`   ⚠️ Реліз v${targetVersion} не знайдено на GitHub`)
          return [] // Ланцюжок розірваний
        }
        
        // Шукаємо патч від попередньої версії до цієї
        const prevVersion = `${major}.${minor}.${patch - 1}`
        const patchAsset = release.assets?.find((asset: any) =>
          asset.name.toLowerCase().includes('patch') &&
          asset.name.includes(prevVersion) &&
          asset.name.endsWith('.zip')
        )
        
        if (!patchAsset) {
          this.log(`   ⚠️ Патч v${prevVersion} → v${targetVersion} не знайдено`)
          return [] // Ланцюжок розірваний
        }
        
        chain.push({
          version: targetVersion,
          release: release,
          patchAsset: patchAsset
        })
        
        currentVer = targetVersion
      }
      
      return chain
      
    } catch (error) {
      this.log(`❌ Помилка пошуку ланцюжка патчів: ${error}`)
      return []
    }
  }

  /**
   * Завантажити та застосувати ланцюжок патчів
   * 
   * @param {Array} patchChain - Масив патчів для застосування
   * @returns {Promise<DownloadResult>} Результат оновлення
   */
  private async downloadAndApplyPatchChain(patchChain: Array<{version: string, release: any, patchAsset: any}>): Promise<DownloadResult> {
    try {
      this.log(`📦 Застосування ланцюжка з ${patchChain.length} патчів...`)
      
      for (let i = 0; i < patchChain.length; i++) {
        const patch = patchChain[i]
        const progress = `[${i + 1}/${patchChain.length}]`
        
        this.log(`${progress} Оновлення до v${patch.version}...`)
        this.emit('status', { message: `Завантаження патчу ${i + 1}/${patchChain.length}...` })
        
        // Завантажити патч
        const downloadPath = path.join(this.updateBasePath, patch.patchAsset.name)
        await this.downloadWithProgress(patch.patchAsset.browser_download_url, downloadPath)
        
        this.log(`${progress} ✅ Завантажено: ${patch.patchAsset.name}`)
        
        // Розпакувати
        this.emit('status', { message: `Встановлення патчу ${i + 1}/${patchChain.length}...` })
        const extractPath = path.join(this.updateBasePath, `extracted-${patch.version}`)
        await this.extractZip(downloadPath, extractPath)
        
        // Застосувати патч (копіювання файлів)
        this.emit('status', { message: `Застосування патчу ${i + 1}/${patchChain.length}...` })
        await this.copyPatchFiles(extractPath)
        
        this.log(`${progress} ✅ Патч застосовано`)
        
        // Видалити тимчасові файли
        try {
          fs.unlinkSync(downloadPath)
          fs.rmSync(extractPath, { recursive: true, force: true })
        } catch (cleanupError) {
          this.log(`⚠️ Помилка очищення: ${cleanupError}`)
        }
      }
      
      this.log(`✅ Всі патчі застосовано успішно!`)
      this.emit('status', { message: 'Оновлення завершено. Перезапуск...' })
      
      // Перезапустити після всіх патчів
      app.relaunch()
      app.exit(0)
      
      return {
        success: true,
        path: this.updateBasePath
      }
      
    } catch (error) {
      this.log(`❌ Помилка застосування ланцюжка патчів: ${error}`)
      return {
        success: false,
        path: false,
        error: error instanceof Error ? error.message : String(error)
      }
    } finally {
      this.downloadInProgress = false
    }
  }

  /**
   * Скопіювати файли патчу до resources/app
   * 
   * @param {string} extractPath - Шлях до розпакованого патчу
   */
  private async copyPatchFiles(extractPath: string): Promise<void> {
    // Поточна папка app (де виконується electron)
    const currentDir = path.dirname(app.getPath('exe'))
    
    // Перевіряємо структуру патчу
    let actualExtractPath = extractPath
    const distFolder = path.join(extractPath, 'dist')
    const packageJson = path.join(extractPath, 'package.json')
    
    // Якщо є dist/, працюємо з цією структурою
    if (fs.existsSync(distFolder) && fs.existsSync(packageJson)) {
      // Патч містить dist/ + package.json
      const targetPath = path.join(currentDir, 'resources', 'app')
      
      if (!fs.existsSync(targetPath)) {
        throw new Error('Папка resources/app/ не знайдена')
      }
      
      // Копіюємо dist/
      const targetDist = path.join(targetPath, 'dist')
      this.copyRecursive(distFolder, targetDist)
      
      // Копіюємо package.json
      fs.copyFileSync(packageJson, path.join(targetPath, 'package.json'))
      this.log('  ✓ package.json')
      
      this.log('✅ Файли патчу скопійовано')
      
    } else {
      throw new Error('Невідома структура патчу')
    }
  }

  /**
   * Рекурсивне копіювання файлів
   */
  private copyRecursive(src: string, dest: string): void {
    if (!fs.existsSync(src)) return
    
    const stat = fs.statSync(src)
    if (stat.isDirectory()) {
      if (!fs.existsSync(dest)) {
        fs.mkdirSync(dest, { recursive: true })
      }
      const files = fs.readdirSync(src)
      for (const file of files) {
        this.copyRecursive(path.join(src, file), path.join(dest, file))
      }
    } else {
      fs.copyFileSync(src, dest)
      this.log(`  ✓ ${path.relative(src, dest)}`)
    }
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
