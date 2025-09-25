/// <reference types="vite/client" />

interface OrderProcessPayload {
  wordBuf: ArrayBuffer
  outputPath: string
  mode?: string
  flags: {
    is2BSP: boolean
    isOrder: boolean
    tokens: boolean
    autoOpen: boolean
  }
  fieldMapping?: {
    pib?: string
    posada?: string
    data?: string
    nomer?: string
    sheet?: string
  }
}

interface Window {
  api?: {
    getVersion(): Promise<string>
    notify(title: string, body: string): Promise<void>
    openExternal(url: string): Promise<void>
    togglePowerBlocker(enable: boolean): Promise<void>
    getSetting(key: string, fallback?: any): Promise<any>
    setSetting(key: string, value: any): Promise<void>
    addNote(text: string): Promise<{id:string, text:string, createdAt:number}>
    listNotes(): Promise<{id:string, text:string, createdAt:number}[]>
    
    // Updates API
    checkForUpdates(): Promise<{
      state: 'idle' | 'checking' | 'uptodate' | 'available' | 'mandatory' | 'downloading' | 'verifying' | 'installing' | 'restarting' | 'failed'
      manifest?: any
      error?: string
      currentVersion: string
    }>
    downloadUpdate(manifest: any): Promise<boolean>
    installUpdate(manifest: any): Promise<boolean>
    getUpdateVersion(): Promise<string>
    getUpdateState(): Promise<string>
    getUpdateProgress(): Promise<any>
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
    
    chooseSavePath(suggestName?: string): Promise<string | null>
    processOrder(payload: OrderProcessPayload): Promise<{
      ok: boolean
      out?: string
      stats?: {
        tokens: number
        paragraphs: number
        matched: number
      }
      error?: string
    }>
  }
}
