import { app, dialog, shell } from 'electron'
import { createHash, createVerify } from 'crypto'
import { promises as fs } from 'fs'
import { join, dirname } from 'path'
import fetch, { Response } from 'node-fetch'
import { createWriteStream, createReadStream } from 'fs'
import { Extract } from 'unzipper'
import { spawn } from 'child_process'
import { EventEmitter } from 'events'

// Типи для маніфесту оновлень
export interface UpdateAsset {
  url: string
  size: number
  sha256: string
  sig: string // RSA підпис
}

export interface UpdateManifest {
  version: string
  channel: string
  published_utc: string
  notes_url: string
  asset: UpdateAsset
  min_supported: string
  mandatory: boolean
}

export interface UpdateInfo {
  latest: UpdateManifest
}

export interface UpdateProgress {
  percent: number
  bytesReceived: number
  totalBytes: number
  speedKbps: number
}

export enum UpdateState {
  Idle = 'idle',
  Checking = 'checking',
  UpToDate = 'uptodate',
  UpdateAvailable = 'available',
  MandatoryUpdate = 'mandatory',
  Downloading = 'downloading',
  Verifying = 'verifying',
  Installing = 'installing',
  Restarting = 'restarting',
  Failed = 'failed'
}

export interface UpdateCheckResult {
  state: UpdateState
  manifest?: UpdateManifest
  error?: string
  currentVersion: string
}

export interface LicenseInfo {
  key: string
  plan: string
  userId: string
  email?: string
  expiresAt?: string
  permissions: string[]
}

export interface UpdateAccessResult {
  hasAccess: boolean
  reason?: string
  licenseInfo?: LicenseInfo
}

class UpdateService extends EventEmitter {
  private currentVersion: string
  private updateBasePath: string
  private manifestUrl: string
  private publicKey: string
  private state: UpdateState = UpdateState.Idle
  private downloadProgress: UpdateProgress | null = null
  private licenseKey: string | null = null
  private updateServerUrl: string
  private githubToken: string | null = null
  
  constructor() {
    super()
    this.currentVersion = '1.1.1'
    
    // Структура папок для оновлень у %LocalAppData%
    const localAppData = process.env.LOCALAPPDATA || process.env.APPDATA || ''
    this.updateBasePath = join(localAppData, 'UkrainianDocumentProcessor')
    
    // URL сервера оновлень з контролем доступу
    this.updateServerUrl = 'https://your-update-server.com/api'
    this.manifestUrl = `${this.updateServerUrl}/updates/check`
    
    // Публічний RSA ключ для верифікації підписів
    this.publicKey = `-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA9h0CuJb1tR1yktMfb0PG
BlG9WBNyU1Q4pCkqOk7nJEnqq8Z5A+Ovy8MQwVTBJ8JuhkH3B5PtuLbvUrT0Pv5n
wzWxCvBGaE6V6CAVDgi252xTPMwTc/mVtIAld8GYgDlFL65Ka81cDxJwInGqdMyc
uOGoxRshI4l3bR7ANQBqvK43LSKFCIFl4dfhEhKPg30G5t0T7I+KBeeICA3CFQqB
yDrRKJDj5sTJVMGZS4yrH1ZeDnz9Q1kBtEX/Cc291EY9I4KibTYSJfuqI3ejbHp4
ZDv+u1zvow5wrCjsFBrrwXPa+56VlfR4nNPFuQFkV8msTMKaV44/jdQFFR9DwgC+
fwIDAQAB
-----END PUBLIC KEY-----`
    
    this.ensureDirectoryStructure()
    // Завантажуємо ключ з затримкою, щоб переконатись що storage вже ініціалізовано
    setTimeout(() => this.loadLicenseKey(), 100)
  }

  // Публічний метод для ініціалізації ліцензії з main процесу
  public async initializeLicense(): Promise<void> {
    await this.loadLicenseKey()
  }

  // Завантаження ліцензійного ключа
  private async loadLicenseKey(): Promise<void> {
    try {
      const { storage } = require('../main')
      if (!storage) {
        this.log('Storage ще не ініціалізовано, повторна спроба через 500ms')
        setTimeout(() => this.loadLicenseKey(), 500)
        return
      }
      this.licenseKey = await storage?.get('licenseKey') || null
      if (this.licenseKey) {
        this.log('Ліцензійний ключ успішно завантажено з storage')
      }
    } catch (error) {
      this.log(`Помилка завантаження ліцензійного ключа: ${error}`)
    }
  }

  // Збереження ліцензійного ключа
  async setLicenseKey(key: string): Promise<UpdateAccessResult> {
    try {
      const accessResult = await this.verifyLicenseKey(key)
      if (accessResult.hasAccess) {
        this.licenseKey = key
        const { storage } = require('../main')
        await storage?.set('licenseKey', key)
        this.log('Ліцензійний ключ успішно збережено')
        return accessResult
      } else {
        this.log(`Недійсний ліцензійний ключ: ${accessResult.reason}`)
        return accessResult
      }
    } catch (error) {
      this.log(`Помилка збереження ліцензійного ключа: ${error}`)
      return {
        hasAccess: false,
        reason: `Помилка збереження ключа: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      }
    }
  }

  // Перевірка доступу до оновлень
  public async checkUpdateAccess(): Promise<UpdateAccessResult> {
    if (!this.licenseKey) {
      return {
        hasAccess: false,
        reason: 'Відсутній ліцензійний ключ. Введіть ключ для отримання оновлень.'
      }
    }

    return await this.verifyLicenseKey(this.licenseKey)
  }

  // Верифікація ліцензійного ключа на сервері
  private async verifyLicenseKey(key: string): Promise<UpdateAccessResult> {
    try {
      // Один універсальний ключ для всіх користувачів
      const MASTER_KEY = 'KONTR-NAHRYUK-2024'

      if (key === MASTER_KEY) {
        return {
          hasAccess: true,
          licenseInfo: {
            key: key,
            plan: 'Universal',
            userId: '001',
            permissions: ['updates', 'beta', 'priority-support', 'advanced-features']
          }
        }
      } else {
        return {
          hasAccess: false,
          reason: 'Недійсний ліцензійний ключ. Використовуйте: KONTR-NAHRYUK-2024'
        }
      }

      /* Реальний запит до сервера:
      const response = await fetch(`${this.updateServerUrl}/license/verify`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'User-Agent': `UkrainianDocumentProcessor/${this.currentVersion}`
        },
        body: JSON.stringify({ key, version: this.currentVersion })
      })

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      return await response.json()
      */
    } catch (error) {
      return {
        hasAccess: false,
        reason: `Помилка перевірки ліцензії: ${error instanceof Error ? error.message : 'Невідома помилка'}`
      }
    }
  }

  // Отримання інформації про поточну ліцензію
  async getLicenseInfo(): Promise<UpdateAccessResult> {
    if (!this.licenseKey) {
      return {
        hasAccess: false,
        reason: 'Ліцензійний ключ не встановлено'
      }
    }

    return await this.verifyLicenseKey(this.licenseKey)
  }

  // Видалення ліцензійного ключа
  async removeLicenseKey(): Promise<void> {
    try {
      this.licenseKey = null
      const { storage } = require('../main')
      await storage?.delete('licenseKey')
      this.log('Ліцензійний ключ видалено')
    } catch (error) {
      this.log(`Помилка видалення ліцензійного ключа: ${error}`)
    }
  }

  // Створення структури папок
  private async ensureDirectoryStructure(): Promise<void> {
    const dirs = [
      join(this.updateBasePath, 'app', 'current'),
      join(this.updateBasePath, 'app', 'backup'),
      join(this.updateBasePath, 'updates', 'staging'),
      join(this.updateBasePath, 'logs')
    ]

    for (const dir of dirs) {
      await fs.mkdir(dir, { recursive: true })
    }
  }

  // Налаштування GitHub токена для приватного репозиторію
  setGitHubToken(token: string): void {
    this.githubToken = token
    this.log('GitHub токен налаштовано для доступу до приватного репозиторію')
  }

  // Отримання заголовків для запитів до GitHub API
  private getGitHubHeaders(): Record<string, string> {
    const headers: Record<string, string> = {
      'User-Agent': 'Ukrainian-Document-Processor',
      'Accept': 'application/vnd.github+json'
    }

    if (this.githubToken) {
      headers['Authorization'] = `Bearer ${this.githubToken}`
    }

    return headers
  }

  // Перевірка доступних оновлень
  async checkForUpdates(): Promise<UpdateCheckResult> {
    try {
      this.setState(UpdateState.Checking)
      this.log('Перевірка оновлень...')

      // Перевіряємо доступ до оновлень
      const accessResult = await this.checkUpdateAccess()
      if (!accessResult.hasAccess) {
        this.setState(UpdateState.Failed)
        return {
          state: UpdateState.Failed,
          error: accessResult.reason || 'Немає доступу до оновлень',
          currentVersion: this.currentVersion
        }
      }

      this.log(`Доступ підтверджено для ключа: ${this.licenseKey}`)

      // Тимчасове рішення: симулюємо перевірку оновлень
      // В реальному проекті тут буде запит до сервера оновлень
      await new Promise(resolve => setTimeout(resolve, 1000)) // Імітація запиту

      // Поки що повертаємо, що програма оновлена
      this.setState(UpdateState.UpToDate)
      this.log(`Поточна версія: ${this.currentVersion} - актуальна`)
      
      return {
        state: UpdateState.UpToDate,
        currentVersion: this.currentVersion
      }

      /* Коментуємо реальний код поки немає сервера
      const response = await fetch(this.manifestUrl, {
        headers: {
          'User-Agent': `UkrainianDocumentProcessor/${this.currentVersion}`,
          'Cache-Control': 'no-cache'
        },
        timeout: 10000
      })

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      const updateInfo: UpdateInfo = await response.json()
      const manifest = updateInfo.latest

      // Верифікація RSA підпису маніфесту
      if (!await this.verifyManifestSignature(manifest)) {
        throw new Error('Невірний підпис маніфесту')
      }
      */

      /* Коментуємо поки немає сервера
      this.log(`Поточна версія: ${this.currentVersion}, доступна: ${manifest.version}`)

      // Порівняння версій
      if (this.compareVersions(manifest.version, this.currentVersion) <= 0) {
        this.setState(UpdateState.UpToDate)
        return {
          state: UpdateState.UpToDate,
          currentVersion: this.currentVersion
        }
      }

      // Перевірка мінімальної підтримуваної версії
      if (this.compareVersions(this.currentVersion, manifest.min_supported) < 0 || manifest.mandatory) {
        this.setState(UpdateState.MandatoryUpdate)
        return {
          state: UpdateState.MandatoryUpdate,
          manifest,
          currentVersion: this.currentVersion
        }
      }

      this.setState(UpdateState.UpdateAvailable)
      return {
        state: UpdateState.UpdateAvailable,
        manifest,
        currentVersion: this.currentVersion
      }
      */

    } catch (error) {
      this.setState(UpdateState.Failed)
      const errorMessage = error instanceof Error ? error.message : 'Невідома помилка'
      this.log(`Помилка перевірки оновлень: ${errorMessage}`)
      return {
        state: UpdateState.Failed,
        error: errorMessage,
        currentVersion: this.currentVersion
      }
    }
  }

  // Завантаження оновлення
  async downloadUpdate(manifest: UpdateManifest): Promise<boolean> {
    try {
      this.setState(UpdateState.Downloading)
      
      const packagePath = join(this.updateBasePath, 'updates', 'package.zip')
      const asset = manifest.asset

      this.log(`Завантаження оновлення з ${asset.url}`)

      // Видалення старого пакету якщо існує
      try {
        await fs.unlink(packagePath)
      } catch (e) {
        // Ігноруємо помилку якщо файл не існує
      }

      const response = await fetch(asset.url)
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      const totalBytes = asset.size || parseInt(response.headers.get('content-length') || '0')
      let bytesReceived = 0
      const startTime = Date.now()

      const fileStream = createWriteStream(packagePath)
      
      return new Promise((resolve, reject) => {
        response.body?.on('data', (chunk: Buffer) => {
          bytesReceived += chunk.length
          const elapsed = (Date.now() - startTime) / 1000
          const speedKbps = elapsed > 0 ? (bytesReceived / 1024) / elapsed : 0
          
          this.downloadProgress = {
            percent: Math.floor((bytesReceived / totalBytes) * 100),
            bytesReceived,
            totalBytes,
            speedKbps
          }
          
          this.emit('download-progress', this.downloadProgress)
        })

        response.body?.pipe(fileStream)
        
        fileStream.on('finish', async () => {
          try {
            // Верифікація SHA-256
            const actualHash = await this.calculateSHA256(packagePath)
            if (actualHash !== asset.sha256.toLowerCase()) {
              await fs.unlink(packagePath)
              reject(new Error(`SHA-256 не збігається. Очікувався: ${asset.sha256}, отримано: ${actualHash}`))
              return
            }

            this.log(`Завантаження завершено. SHA-256 підтверджено.`)
            resolve(true)
          } catch (error) {
            reject(error)
          }
        })

        fileStream.on('error', (error) => {
          reject(error)
        })

        response.body?.on('error', (error: any) => {
          reject(error)
        })
      })

    } catch (error) {
      this.setState(UpdateState.Failed)
      const errorMessage = error instanceof Error ? error.message : 'Помилка завантаження'
      this.log(`Помилка завантаження: ${errorMessage}`)
      return false
    }
  }

  // Встановлення оновлення
  async installUpdate(manifest: UpdateManifest): Promise<boolean> {
    try {
      this.setState(UpdateState.Installing)
      
      const packagePath = join(this.updateBasePath, 'updates', 'package.zip')
      const stagingPath = join(this.updateBasePath, 'updates', 'staging', manifest.version)

      this.log(`Встановлення оновлення до версії ${manifest.version}`)

      // Очищення staging директорії
      try {
        await fs.rm(stagingPath, { recursive: true, force: true })
      } catch (e) {
        // Ігноруємо помилку якщо директорія не існує
      }
      
      await fs.mkdir(stagingPath, { recursive: true })

      // Розпаковка ZIP файлу
      await this.extractZip(packagePath, stagingPath)

      // Запуск процесу оновлення
      await this.launchUpdater(manifest.version)
      
      return true

    } catch (error) {
      this.setState(UpdateState.Failed)
      const errorMessage = error instanceof Error ? error.message : 'Помилка встановлення'
      this.log(`Помилка встановлення: ${errorMessage}`)
      return false
    }
  }

  // Запуск процесу оновлення та перезапуск застосунку
  private async launchUpdater(newVersion: string): Promise<void> {
    this.setState(UpdateState.Restarting)
    
    const updaterPath = join(__dirname, '..', '..', 'updater.js')
    const appDir = join(this.updateBasePath, 'app')
    const stagingDir = join(this.updateBasePath, 'updates', 'staging', newVersion)
    const backupDir = join(this.updateBasePath, 'app', 'backup')
    const currentExePath = process.execPath

    const args = [
      updaterPath,
      '--appdir', appDir,
      '--staging', stagingDir,
      '--backup', backupDir,
      '--relaunch', currentExePath,
      '--pid', process.pid.toString()
    ]

    this.log(`Запуск updater: node ${args.join(' ')}`)

    // Запуск процесу оновлення
    const updater = spawn('node', args, {
      detached: true,
      stdio: 'ignore'
    })

    updater.unref()

    // Graceful закриття основного застосунку
    setTimeout(() => {
      app.quit()
    }, 1000)
  }

  // Допоміжні методи
  private async verifyManifestSignature(manifest: UpdateManifest): Promise<boolean> {
    try {
      const canonicalJson = JSON.stringify(manifest, Object.keys(manifest).sort())
      const verify = createVerify('SHA256')
      verify.update(canonicalJson)
      verify.end()
      
      return verify.verify(this.publicKey, manifest.asset.sig, 'base64')
    } catch (error) {
      this.log(`Помилка верифікації підпису: ${error}`)
      return false
    }
  }

  private async calculateSHA256(filePath: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const hash = createHash('sha256')
      const stream = createReadStream(filePath)
      
      stream.on('data', data => hash.update(data))
      stream.on('end', () => resolve(hash.digest('hex')))
      stream.on('error', reject)
    })
  }

  private async extractZip(zipPath: string, extractPath: string): Promise<void> {
    return new Promise((resolve, reject) => {
      createReadStream(zipPath)
        .pipe(Extract({ path: extractPath }))
        .on('close', resolve)
        .on('error', reject)
    })
  }

  private compareVersions(a: string, b: string): number {
    const partsA = a.split('.').map(x => parseInt(x))
    const partsB = b.split('.').map(x => parseInt(x))
    
    for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
      const partA = partsA[i] || 0
      const partB = partsB[i] || 0
      
      if (partA > partB) return 1
      if (partA < partB) return -1
    }
    
    return 0
  }

  private setState(newState: UpdateState): void {
    this.state = newState
    this.emit('state-changed', newState)
  }

  private async log(message: string): Promise<void> {
    const timestamp = new Date().toISOString()
    const logMessage = `${timestamp} - ${message}\n`
    const logPath = join(this.updateBasePath, 'logs', `update-${new Date().toISOString().slice(0, 10)}.txt`)
    
    try {
      await fs.appendFile(logPath, logMessage)
    } catch (error) {
      console.error('Помилка запису логу:', error)
    }
    
    console.log(`[UpdateService] ${message}`)
  }

  // Публічні властивості
  getCurrentVersion(): string {
    return this.currentVersion
  }

  getState(): UpdateState {
    return this.state
  }

  getDownloadProgress(): UpdateProgress | null {
    return this.downloadProgress
  }

  // Новий метод для перевірки оновлень через GitHub API
  async checkForUpdatesViaGitHub(): Promise<any> {
    try {
      this.log('Перевірка оновлень через GitHub API...')
      
      const response = await fetch('https://api.github.com/repos/sashashostak/KontrNahryuk/releases/latest', {
        method: 'GET',
        headers: {
          'User-Agent': `KontrNahryuk/${this.currentVersion}`,
          'Accept': 'application/vnd.github.v3+json'
        },
        timeout: 10000 // 10 секунд тайм-аут
      })

      if (!response.ok) {
        throw new Error(`GitHub API відповів з помилкою: ${response.status} ${response.statusText}`)
      }

      const release = await response.json()
      const latestVersion = release.tag_name || release.name || 'unknown'
      const currentVersion = this.currentVersion

      this.log(`Поточна версія: ${currentVersion}, Остання версія: ${latestVersion}`)

      // Перевіряємо чи є оновлення
      const hasUpdate = latestVersion !== currentVersion && 
                       latestVersion !== `v${currentVersion}` &&
                       latestVersion !== currentVersion.replace(/^v/, '')

      return {
        hasUpdate,
        latestVersion,
        currentVersion,
        releaseInfo: hasUpdate ? release : null,
        error: null
      }

    } catch (error) {
      this.log(`Помилка перевірки оновлень: ${error instanceof Error ? error.message : String(error)}`)
      return {
        hasUpdate: false,
        latestVersion: null,
        currentVersion: this.currentVersion,
        releaseInfo: null,
        error: error instanceof Error ? error.message : 'Невідома помилка мережі'
      }
    }
  }
}

export { UpdateService }