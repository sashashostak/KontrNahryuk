import { contextBridge, ipcRenderer } from 'electron'

console.log('[preload] loaded')

// API interfaces for better type safety
interface OrderProcessPayload {
  hasWordBuf: boolean
  hasExcelBuf: boolean
  outputPath: string
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
  getUpdateVersion(): Promise<string>
  getUpdateState(): Promise<string>
  getUpdateProgress(): Promise<any>
  setLicenseKey(key: string): Promise<any>
  getLicenseInfo(): Promise<any>
  checkUpdateAccess(): Promise<any>
  onUpdateStateChanged(callback: (state: string) => void): void
  onUpdateProgress(callback: (progress: any) => void): void
  
  // Batch Processing API
  selectBatchDirectory(): Promise<string | undefined>
  selectBatchOutputFile(): Promise<string | undefined>
  startBatchProcessing(options: any): Promise<any>
  cancelBatchProcessing(): Promise<boolean>
  isBatchRunning(): Promise<boolean>
  onBatchProgress(callback: (progress: any) => void): void
  onBatchLog(callback: (logEntry: { level: string, message: string }) => void): void
  onBatchComplete(callback: (result: any) => void): void
  
  chooseSavePath(suggestName?: string): Promise<string | undefined>
  
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
  getUpdateVersion: (): Promise<string> => ipcRenderer.invoke('updates:get-version'),
  getUpdateState: (): Promise<string> => ipcRenderer.invoke('updates:get-state'),
  getUpdateProgress: (): Promise<any> => ipcRenderer.invoke('updates:get-progress'),
  setLicenseKey: (key: string): Promise<any> => ipcRenderer.invoke('updates:set-license', key),
  getLicenseInfo: (): Promise<any> => ipcRenderer.invoke('updates:get-license-info'),
  checkUpdateAccess: (): Promise<any> => ipcRenderer.invoke('updates:check-access'),
  onUpdateStateChanged: (callback: (state: string) => void): void => {
    ipcRenderer.on('updates:state-changed', (_, state) => callback(state))
  },
  onUpdateProgress: (callback: (progress: any) => void): void => {
    ipcRenderer.on('updates:download-progress', (_, progress) => callback(progress))
  },

  // Batch Processing API
  selectBatchDirectory: (): Promise<string | undefined> => ipcRenderer.invoke('batch:select-directory'),
  selectBatchOutputFile: (): Promise<string | undefined> => ipcRenderer.invoke('batch:select-output-file'),
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
  
  processOrder: (payload: OrderProcessPayload): Promise<any> => ipcRenderer.invoke('order:process', payload),
} as ElectronAPI)
