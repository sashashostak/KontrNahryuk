import { app, dialog, BrowserWindow, ipcMain } from 'electron'
import { autoUpdater } from 'electron-updater'

let mainWindow: BrowserWindow | null = null

export function setupAutoUpdate(window: BrowserWindow) {
  mainWindow = window
  
  // Налаштування автооновлення
  autoUpdater.autoDownload = true
  autoUpdater.allowPrerelease = false
  
  // Налаштування IPC обробників
  ipcMain.handle('check-for-updates', async () => {
    if (app.isPackaged) {
      try {
        await autoUpdater.checkForUpdates()
      } catch (error) {
        console.error('[autoUpdater] Manual check error:', error)
      }
    }
  })
  
  ipcMain.handle('install-update', () => {
    autoUpdater.quitAndInstall()
  })

  // Обробка подій автооновлення
  autoUpdater.on('update-available', (info) => {
    console.log('[autoUpdater] update available:', info.version)
    mainWindow?.webContents.send('update-available', {
      version: info.version,
      releaseNotes: info.releaseNotes
    })
  })

  autoUpdater.on('download-progress', (progress) => {
    console.log(`[autoUpdater] download ${progress.percent.toFixed(1)}%`)
    mainWindow?.webContents.send('download-progress', {
      percent: progress.percent,
      transferred: progress.transferred,
      total: progress.total
    })
  })

  autoUpdater.on('update-downloaded', async (info) => {
    console.log('[autoUpdater] update downloaded:', info.version)
    mainWindow?.webContents.send('update-downloaded', {
      version: info.version
    })
    
    // Показати діалог користувачу
    const res = await dialog.showMessageBox(mainWindow!, {
      type: 'question',
      buttons: ['Перезапустити зараз', 'Потім'],
      defaultId: 0,
      title: 'Оновлення завантажено',
      message: `Завантажено нову версію ${info.version}. Перезапустити застосунок зараз?`,
    })
    
    if (res.response === 0) {
      autoUpdater.quitAndInstall()
    }
  })

  autoUpdater.on('update-not-available', () => {
    console.log('[autoUpdater] no updates available')
  })

  autoUpdater.on('error', (error) => {
    console.error('[autoUpdater] error:', error)
    mainWindow?.webContents.send('update-error', error.message)
  })

  // Запуск перевірки оновлень
  if (app.isPackaged) {
    setTimeout(() => {
      autoUpdater.checkForUpdates().catch(err => console.error('[autoUpdater] check error', err))
    }, 3000)
    
    // Перевіряти оновлення кожні 4 години
    setInterval(() => {
      autoUpdater.checkForUpdates().catch(err => console.error('[autoUpdater] periodic check error', err))
    }, 4 * 60 * 60 * 1000)
  }
}
