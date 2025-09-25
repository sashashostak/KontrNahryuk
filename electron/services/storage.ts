import { app } from 'electron'
import path from 'node:path'
import fs from 'node:fs/promises'
import { existsSync, mkdirSync } from 'node:fs'

type DBData = {
  settings: { [k: string]: any }
  notes: { id: string, text: string, createdAt: number }[]
}

export interface Storage {
  getSetting<T=any>(key: string, fallback?: T): Promise<T>
  setSetting<T=any>(key: string, value: T): Promise<void>
  addNote(text: string): Promise<{ id: string, text: string, createdAt: number }>
  listNotes(): Promise<{ id: string, text: string, createdAt: number }[]>
}

class JSONStorage implements Storage {
  private filePath: string
  private data: DBData = { settings: {}, notes: [] }
  
  constructor(file: string) {
    this.filePath = file
    this.init().catch(console.error)
  }
  
  private async init() {
    try {
      const content = await fs.readFile(this.filePath, 'utf-8')
      this.data = JSON.parse(content)
    } catch (error) {
      // File doesn't exist or is invalid, use default data
      this.data = { settings: {}, notes: [] }
      await this.save()
    }
  }
  
  private async save() {
    await fs.writeFile(this.filePath, JSON.stringify(this.data, null, 2), 'utf-8')
  }
  
  async getSetting<T>(key: string, fallback?: T): Promise<T> {
    return (this.data.settings[key] as T) ?? (fallback as T)
  }
  
  async setSetting<T>(key: string, value: T): Promise<void> {
    this.data.settings[key] = value
    await this.save()
  }
  
  async addNote(text: string): Promise<{ id: string, text: string, createdAt: number }> {
    const note = { id: Math.random().toString(36).slice(2, 9), text, createdAt: Date.now() }
    this.data.notes.push(note)
    await this.save()
    return note
  }
  
  async listNotes(): Promise<{ id: string, text: string, createdAt: number }[]> {
    return this.data.notes
  }
}

export function createStorage(): Storage {
  const dataDir = app.getPath('userData')
  if (!existsSync(dataDir)) mkdirSync(dataDir, { recursive: true })

  const jsonPath = path.join(dataDir, 'db.json')
  return new JSONStorage(jsonPath)
}
