import { app, nativeTheme, Notification, powerSaveBlocker, shell, protocol } from 'electron'
import path from 'node:path'

let powerBlockerId: number | null = null

export function setupOSIntegration() {
  // У macOS/Win можна налаштувати about-інформацію
  if (process.platform === 'darwin') {
    app.setAboutPanelOptions({
      applicationName: app.getName(),
      applicationVersion: app.getVersion(),
      credits: 'Створено з Electron ❤️',
    })
  }

  // Реєстрація кастомного протоколу: myapp://open?x=1
  if (app.isDefaultProtocolClient('myapp') === false) {
    try { 
      app.setAsDefaultProtocolClient('myapp') 
    } catch (error) {
      console.warn('Failed to set as default protocol client:', error)
    }
  }

  // Нічна тема за системою (приклад API)
  nativeTheme.on('updated', () => {
    console.log('[nativeTheme] shouldUseDarkColors=', nativeTheme.shouldUseDarkColors)
  })
}

export function notify(title: string, body: string) {
  if (Notification.isSupported()) {
    new Notification({ title, body, icon: path.join(process.resourcesPath, 'icon.png') }).show()
  }
}

export function togglePowerSaveBlocker(enable: boolean) {
  if (enable && powerBlockerId == null) {
    powerBlockerId = powerSaveBlocker.start('prevent-app-suspension')
  } else if (!enable && powerBlockerId != null) {
    powerSaveBlocker.stop(powerBlockerId)
    powerBlockerId = null
  }
}

export async function openExternal(url: string) {
  await shell.openExternal(url)
}
