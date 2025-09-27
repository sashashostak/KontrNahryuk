import { updateService, UpdateInfo } from '../services/updateService'

export class UpdateNotification {
  private container: HTMLElement
  private currentInfo: UpdateInfo = {
    version: '',
    available: false,
    downloading: false,
    downloaded: false
  }

  constructor(containerId: string) {
    const element = document.getElementById(containerId)
    if (!element) {
      throw new Error(`Element with id "${containerId}" not found`)
    }
    this.container = element
    this.setupEventListeners()
    this.render()
  }

  private setupEventListeners() {
    updateService.subscribe((info) => {
      this.currentInfo = info
      this.render()
    })
  }

  private render() {
    if (!this.currentInfo.available && !this.currentInfo.downloading && !this.currentInfo.downloaded) {
      this.container.innerHTML = `
        <div class="update-notification">
          <button id="check-updates-btn" class="update-check-btn" title="Перевірити оновлення">
            🔄 Перевірити оновлення
          </button>
        </div>
      `
      
      const checkBtn = document.getElementById('check-updates-btn')
      checkBtn?.addEventListener('click', () => updateService.checkForUpdates())
      return
    }

    if (this.currentInfo.downloading) {
      const progress = Math.round(this.currentInfo.progress || 0)
      this.container.innerHTML = `
        <div class="update-notification downloading">
          <div class="update-info">
            <span>⬇️ Завантажується оновлення v${this.currentInfo.version}</span>
            <div class="progress-bar">
              <div class="progress-fill" style="width: ${progress}%"></div>
            </div>
            <span class="progress-text">${progress}%</span>
          </div>
        </div>
      `
      return
    }

    if (this.currentInfo.downloaded) {
      this.container.innerHTML = `
        <div class="update-notification ready">
          <div class="update-info">
            <span>✅ Оновлення v${this.currentInfo.version} готове до встановлення</span>
            <button id="install-update-btn" class="install-btn">
              Перезапустити і встановити
            </button>
          </div>
        </div>
      `
      
      const installBtn = document.getElementById('install-update-btn')
      installBtn?.addEventListener('click', () => updateService.installUpdate())
      return
    }

    if (this.currentInfo.available) {
      const releaseNotesHtml = this.currentInfo.releaseNotes 
        ? `<details class="release-notes">
             <summary>Що нового?</summary>
             <div>${this.currentInfo.releaseNotes}</div>
           </details>`
        : ''
      
      this.container.innerHTML = `
        <div class="update-notification available">
          <div class="update-info">
            <span>🆕 Доступна нова версія v${this.currentInfo.version}</span>
            ${releaseNotesHtml}
          </div>
        </div>
      `
      return
    }

    this.container.innerHTML = ''
  }

  public destroy() {
    this.container.innerHTML = ''
  }
}