import { contextBridge, ipcRenderer } from 'electron'

console.log('[preload] loaded')

// API interfaces for better type safety
interface OrderProcessPayload {
  wordBuf: ArrayBuffer
  outputPath: string
  excelPath?: string
  flags: {
    saveDBPath: boolean
    is2BSP: boolean
    isOrder: boolean
    tokens: boolean
    autoOpen: boolean
  }
  mode: string
}

interface ElectronAPI {
  notify(title: string, body: string): Promise<void>
  openExternal(url: string): Promise<void>
  togglePowerBlocker(enable: boolean): Promise<void>
  
  getSetting(key: string, fallback?: any): Promise<any>
  setSetting(key: string, value: any): Promise<void>
  addNote(text: string): Promise<void>
  listNotes(): Promise<any[]>
  
  // Updates API
  checkForUpdates(): Promise<any>
  downloadUpdate(manifest: any): Promise<boolean>
  installUpdate(manifest: any): Promise<boolean>
  downloadAndInstallUpdate(updateInfo: any): Promise<boolean>
  cancelUpdate(): Promise<boolean>
  restartApp(): Promise<void>
  getUpdateVersion(): Promise<string>
  getUpdateState(): Promise<string>
  getUpdateProgress(): Promise<any>
  setLicenseKey(key: string): Promise<any>
  getLicenseInfo(): Promise<any>
  checkUpdateAccess(): Promise<any>
  checkExistingLicense(): Promise<any>
  onUpdateStateChanged(callback: (state: string) => void): void
  onUpdateProgress(callback: (progress: any) => void): void
  onUpdateError(callback: (error: string) => void): void
  onUpdateComplete(callback: () => void): void
  
  // Batch Processing API
  selectBatchDirectory(): Promise<string | undefined>
  selectBatchOutputFile(): Promise<string | undefined>
  selectExcelFile(): Promise<string | undefined>
  scanExcelFiles(folderPath: string): Promise<string[]>
  startBatchProcessing(options: any): Promise<any>
  cancelBatchProcessing(): Promise<boolean>
  isBatchRunning(): Promise<boolean>
  onBatchProgress(callback: (progress: any) => void): void
  onBatchLog(callback: (logEntry: { level: string, message: string }) => void): void
  onBatchComplete(callback: (result: any) => void): void
  
  chooseSavePath(suggestName?: string): Promise<string | undefined>
  selectFolder(): Promise<{ filePath: string } | undefined>
  
  processOrder(payload: OrderProcessPayload): Promise<any>
}

contextBridge.exposeInMainWorld('api', {
  notify: (title: string, body: string): Promise<void> => ipcRenderer.invoke('os:notify', { title, body }),
  openExternal: (url: string): Promise<void> => ipcRenderer.invoke('os:openExternal', { url }),
  togglePowerBlocker: (enable: boolean): Promise<void> => ipcRenderer.invoke('os:powerBlocker', { enable }),

  getSetting: (key: string, fallback?: any): Promise<any> => ipcRenderer.invoke('storage:getSetting', { key, fallback }),
  setSetting: (key: string, value: any): Promise<void> => ipcRenderer.invoke('storage:setSetting', { key, value }),
  addNote: (text: string): Promise<void> => ipcRenderer.invoke('storage:addNote', { text }),
  listNotes: (): Promise<any[]> => ipcRenderer.invoke('storage:listNotes'),

  // Updates API
  checkForUpdates: (): Promise<any> => ipcRenderer.invoke('updates:check-github'),
  downloadUpdate: (manifest: any): Promise<boolean> => ipcRenderer.invoke('updates:download', manifest),
  installUpdate: (manifest: any): Promise<boolean> => ipcRenderer.invoke('updates:install', manifest),
  downloadAndInstallUpdate: (updateInfo: any): Promise<boolean> => ipcRenderer.invoke('updates:download-and-install', updateInfo),
  cancelUpdate: (): Promise<boolean> => ipcRenderer.invoke('updates:cancel'),
  restartApp: (): Promise<void> => ipcRenderer.invoke('updates:restart-app'),
  getUpdateVersion: (): Promise<string> => ipcRenderer.invoke('updates:get-version'),
  getUpdateState: (): Promise<string> => ipcRenderer.invoke('updates:get-state'),
  getUpdateProgress: (): Promise<any> => ipcRenderer.invoke('updates:get-progress'),
  setLicenseKey: (key: string): Promise<any> => ipcRenderer.invoke('updates:set-license', key),
  getLicenseInfo: (): Promise<any> => ipcRenderer.invoke('updates:get-license-info'),
  checkUpdateAccess: (): Promise<any> => ipcRenderer.invoke('updates:check-access'),
  checkExistingLicense: (): Promise<any> => ipcRenderer.invoke('updates:check-existing-license'),
  saveUpdateLog: (content: string): Promise<boolean> => ipcRenderer.invoke('updates:save-log', content),
  onUpdateStateChanged: (callback: (state: string) => void): void => {
    ipcRenderer.on('updates:state-changed', (_, state) => callback(state))
  },
  onUpdateProgress: (callback: (progress: any) => void): void => {
    ipcRenderer.on('updates:progress', (_, progress) => callback(progress))
  },
  onUpdateError: (callback: (error: string) => void): void => {
    ipcRenderer.on('updates:error', (_, error) => callback(error))
  },
  onUpdateDownloadStarted: (callback: (info: any) => void): void => {
    ipcRenderer.on('updates:download-started', (_, info) => callback(info))
  },
  onUpdateDownloadCompleted: (callback: (info: any) => void): void => {
    ipcRenderer.on('updates:download-completed', (_, info) => callback(info))
  },
  onUpdateComplete: (callback: () => void): void => {
    ipcRenderer.on('updates:complete', () => callback())
  },

  // Batch Processing API
  selectBatchDirectory: (): Promise<string | undefined> => ipcRenderer.invoke('batch:select-directory'),
  selectBatchOutputFile: (): Promise<string | undefined> => ipcRenderer.invoke('batch:select-output-file'),
  selectExcelFile: (): Promise<string | undefined> => ipcRenderer.invoke('batch:select-excel-file'),
  scanExcelFiles: (folderPath: string): Promise<string[]> => ipcRenderer.invoke('batch:scan-excel-files', folderPath),
  startBatchProcessing: (options: any): Promise<any> => ipcRenderer.invoke('batch:process', options),
  cancelBatchProcessing: (): Promise<boolean> => ipcRenderer.invoke('batch:cancel'),
  isBatchRunning: (): Promise<boolean> => ipcRenderer.invoke('batch:is-running'),
  onBatchProgress: (callback: (progress: any) => void): void => {
    ipcRenderer.on('batch:progress', (_, progress) => callback(progress))
  },
  onBatchLog: (callback: (logEntry: { level: string, message: string }) => void): void => {
    ipcRenderer.on('batch:log', (_, logEntry) => callback(logEntry))
  },
  onBatchComplete: (callback: (result: any) => void): void => {
    ipcRenderer.on('batch:complete', (_, result) => callback(result))
  },

  chooseSavePath: (suggestName?: string): Promise<string | undefined> => ipcRenderer.invoke('dialog:save', { suggestName }),
  selectFolder: (): Promise<{ filePath: string } | undefined> => ipcRenderer.invoke('dialog:select-folder'),
  
  processOrder: (payload: OrderProcessPayload): Promise<any> => ipcRenderer.invoke('order:process', payload),
} as ElectronAPI)
