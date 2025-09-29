import { ipcRenderer } from 'electron'

export interface UpdateInfo {
  version: string
  releaseNotes?: string
  available: boolean
  downloading: boolean
  downloaded: boolean
  progress?: number
}

class UpdateService {
  private listeners: ((info: UpdateInfo) => void)[] = []
  private currentInfo: UpdateInfo = {
    version: '',
    available: false,
    downloading: false,
    downloaded: false
  }

  constructor() {
    this.setupListeners()
  }

  private setupListeners() {
    // Слухаємо події від головного процесу
    ipcRenderer.on('update-available', (_, info) => {
      this.currentInfo = {
        ...this.currentInfo,
        available: true,
        version: info.version,
        releaseNotes: info.releaseNotes
      }
      this.notifyListeners()
    })

    ipcRenderer.on('download-progress', (_, progress) => {
      this.currentInfo = {
        ...this.currentInfo,
        downloading: true,
        progress: progress.percent
      }
      this.notifyListeners()
    })

    ipcRenderer.on('update-downloaded', () => {
      this.currentInfo = {
        ...this.currentInfo,
        downloading: false,
        downloaded: true,
        progress: 100
      }
      this.notifyListeners()
    })

    ipcRenderer.on('update-error', (_, error) => {
      console.error('Помилка оновлення:', error)
      this.currentInfo = {
        ...this.currentInfo,
        downloading: false,
        available: false
      }
      this.notifyListeners()
    })
  }

  // Підписка на зміни статусу оновлення
  subscribe(callback: (info: UpdateInfo) => void) {
    this.listeners.push(callback)
    return () => {
      this.listeners = this.listeners.filter(l => l !== callback)
    }
  }

  // Перевірити оновлення вручну
  checkForUpdates() {
    ipcRenderer.invoke('check-for-updates')
  }

  // Встановити оновлення
  installUpdate() {
    if (this.currentInfo.downloaded) {
      ipcRenderer.invoke('install-update')
    }
  }

  // Отримати поточний статус
  getStatus(): UpdateInfo {
    return { ...this.currentInfo }
  }

  private notifyListeners() {
    this.listeners.forEach(listener => listener({ ...this.currentInfo }))
  }
}

export const updateService = new UpdateService()