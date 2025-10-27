import { contextBridge, ipcRenderer } from 'electron'

// API interfaces for better type safety
interface OrderProcessPayload {
  wordBuf: ArrayBuffer
  outputPath: string
  excelPath?: string
  excelSheetsCount?: number
  searchText?: string
  flags: {
    saveDBPath: boolean
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
  
  // ============================================================================
  // Updates API (Simplified - based on TZ Order Processor)
  // ============================================================================
  getVersion(): Promise<string>
  checkForUpdates(): Promise<any>
  downloadUpdate(updateInfo: any): Promise<any>
  restartApp(): Promise<void>
  invoke(channel: string, ...args: any[]): Promise<any>
  
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
  
  // Logger API
  onLog(callback: (level: string, message: string) => void): void
  
  chooseSavePath(suggestName?: string): Promise<string | undefined>
  selectFolder(): Promise<{ filePath: string } | undefined>
  selectFile(options?: { title?: string, filters?: Array<{ name: string, extensions: string[] }> }): Promise<string | undefined>
  readDirectory(folderPath: string): Promise<Array<{ name: string, path: string }>>
  readExcelFile(filePath: string): Promise<ArrayBuffer>
  writeExcelFile(filePath: string, buffer: ArrayBuffer): Promise<void>

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

  // ============================================================================
  // Updates API (Advanced - with Progress Bar)
  // ============================================================================
  getVersion: (): Promise<string> => ipcRenderer.invoke('updates:get-version'),
  checkForUpdates: (): Promise<any> => ipcRenderer.invoke('updates:check'),
  downloadUpdate: (updateInfo: any): Promise<any> => ipcRenderer.invoke('updates:download', updateInfo),
  restartApp: (): Promise<void> => ipcRenderer.invoke('updates:restart-app'),
  invoke: (channel: string, ...args: any[]): Promise<any> => ipcRenderer.invoke(channel, ...args),

  // Event listeners для прогрес-бару та статусів
  onDownloadProgress: (callback: (progress: any) => void): void => {
    ipcRenderer.on('update:download-progress', (_, progress) => callback(progress))
  },
  onUpdateStatus: (callback: (status: any) => void): void => {
    ipcRenderer.on('update:status', (_, status) => callback(status))
  },
  onUpdateError: (callback: (error: any) => void): void => {
    ipcRenderer.on('update:error', (_, error) => callback(error))
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

  // Logger API
  onLog: (callback: (level: string, message: string) => void): void => {
    ipcRenderer.on('main:log', (_, level, message) => callback(level, message))
  },

  chooseSavePath: (suggestName?: string): Promise<string | undefined> => ipcRenderer.invoke('dialog:save', { suggestName }),
  selectFolder: (): Promise<{ filePath: string } | undefined> => ipcRenderer.invoke('dialog:select-folder'),
  selectFile: (options?: { title?: string, filters?: Array<{ name: string, extensions: string[] }> }): Promise<string | undefined> => ipcRenderer.invoke('dialog:select-file', options),
  readDirectory: (folderPath: string): Promise<Array<{ name: string, path: string }>> => ipcRenderer.invoke('fs:read-directory', folderPath),
  readExcelFile: (filePath: string): Promise<ArrayBuffer> => ipcRenderer.invoke('fs:read-excel-file', filePath),
  writeExcelFile: (filePath: string, buffer: ArrayBuffer): Promise<void> => ipcRenderer.invoke('fs:write-excel-file', filePath, buffer),

  processOrder: (payload: OrderProcessPayload): Promise<any> => ipcRenderer.invoke('order:process', payload),
} as ElectronAPI)
