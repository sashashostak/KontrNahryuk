import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron'
import path from 'node:path'
import fs from 'node:fs/promises'
import * as mammoth from 'mammoth'
import { Document, Paragraph, Packer, TextRun } from 'docx'

import { setupOSIntegration, notify, openExternal, togglePowerSaveBlocker } from './services/osIntegration'
import { createStorage } from './services/storage'
import { UpdateService, UpdateState } from './services/updateService'

const isDev = process.env.NODE_ENV !== 'production' && (process.env.VITE_DEV_SERVER_URL !== undefined || process.argv.includes('--dev'))

let storage: any
let updateService: UpdateService

function norm(str: string): string {
  return str.toLowerCase()
    .replace(/[«»"""''`]/g, '"')
    .replace(/[—–−]/g, '-')
    .replace(/…/g, '...')
    .replace(/\s+/g, ' ')
    .trim()
}

function createWindow(): BrowserWindow {
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })

  if (isDev) {
    const devURL = process.env.VITE_DEV_SERVER_URL || 'http://localhost:5177'
    mainWindow.loadURL(devURL)
    mainWindow.webContents.openDevTools()
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'))
  }

  // Додаємо можливість відкрити DevTools в продакшн режимі (F12)
  mainWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'F12') {
      mainWindow.webContents.toggleDevTools()
    }
  })
  
  return mainWindow
}

function setupUpdateHandlers() {
  // Перевірка оновлень
  ipcMain.handle('updates:check', async () => {
    try {
      const result = await updateService.checkForUpdates()
      return result
    } catch (error) {
      console.error('Помилка перевірки оновлень:', error)
      return { 
        state: UpdateState.Failed, 
        error: error instanceof Error ? error.message : 'Невідома помилка',
        currentVersion: updateService.getCurrentVersion()
      }
    }
  })

  // Завантаження оновлення
  ipcMain.handle('updates:download', async (_, manifest) => {
    try {
      return await updateService.downloadUpdate(manifest)
    } catch (error) {
      console.error('Помилка завантаження оновлення:', error)
      return false
    }
  })

  // Встановлення оновлення
  ipcMain.handle('updates:install', async (_, manifest) => {
    try {
      return await updateService.installUpdate(manifest)
    } catch (error) {
      console.error('Помилка встановлення оновлення:', error)
      return false
    }
  })

  // Отримання поточної версії
  ipcMain.handle('updates:get-version', () => {
    return updateService.getCurrentVersion()
  })

  // Отримання стану оновлення
  ipcMain.handle('updates:get-state', () => {
    return updateService.getState()
  })

  // Отримання прогресу завантаження
  ipcMain.handle('updates:get-progress', () => {
    return updateService.getDownloadProgress()
  })

  // Встановлення ліцензійного ключа
  ipcMain.handle('updates:set-license', async (_, key: string) => {
    return await updateService.setLicenseKey(key)
  })

  // Перевірка існуючого ліцензійного ключа при запуску
  ipcMain.handle('updates:check-existing-license', async () => {
    return await updateService.checkUpdateAccess()
  })

  // Отримання інформації про ліцензію
  ipcMain.handle('updates:get-license-info', async () => {
    return await updateService.getLicenseInfo()
  })

  ipcMain.handle('updates:check-access', async () => {
    return await updateService.checkUpdateAccess()
  })

  // Перевірка оновлень через GitHub API
  ipcMain.handle('updates:check-github', async () => {
    return await updateService.checkForUpdatesViaGitHub()
  })

  // Автооновлення - завантаження і встановлення через GitHub Releases
  ipcMain.handle('updates:download-and-install', async (_, updateInfo) => {
    try {
      console.log('Спроба завантажити оновлення:', updateInfo)
      
      const releaseInfo = updateInfo.releaseInfo
      if (!releaseInfo) {
        throw new Error('Відсутня інформація про реліз')
      }

      // Шукаємо portable ZIP файл в assets
      const portableAsset = releaseInfo.assets?.find((asset: any) => 
        asset.name.toLowerCase().includes('portable') && asset.name.endsWith('.zip')
      )
      
      if (!portableAsset) {
        // Якщо немає portable файлу, відкриваємо сторінку релізу для ручного завантаження
        shell.openExternal(releaseInfo.html_url)
        throw new Error('Автоматичне оновлення недоступне. Відкрито сторінку для ручного завантаження.')
      }

      // Повідомляємо користувача що почали завантаження
      BrowserWindow.getAllWindows().forEach(window => {
        window.webContents.send('updates:download-started', {
          fileName: portableAsset.name,
          size: portableAsset.size
        })
      })

      // Завантажуємо portable версію через GitHub API
      const success = await updateService.downloadFromGitHub(portableAsset)
      
      if (success) {
        // Повідомляємо про успішне завантаження
        BrowserWindow.getAllWindows().forEach(window => {
          window.webContents.send('updates:download-completed', {
            filePath: success
          })
        })
        return true
      } else {
        throw new Error('Не вдалося завантажити файл оновлення')
      }

    } catch (error) {
      console.error('Помилка автооновлення:', error)
      const errorMessage = error instanceof Error ? error.message : String(error)
      BrowserWindow.getAllWindows().forEach(window => {
        window.webContents.send('updates:error', errorMessage)
      })
      return false
    }
  })

  // Скасування оновлення
  ipcMain.handle('updates:cancel', async () => {
    try {
      // Тут можна додати логіку для скасування завантаження
      // Поки що просто повертаємо true
      return true
    } catch (error) {
      console.error('Помилка скасування оновлення:', error)
      return false
    }
  })

  // Збереження лог файлу оновлень
  ipcMain.handle('updates:save-log', async (_, content: string) => {
    try {
      const { dialog } = require('electron')
      const fs = require('fs')
      const path = require('path')
      const os = require('os')
      
      // Пропонуємо зберегти у Downloads
      const defaultPath = path.join(os.homedir(), 'Downloads', `KontrNahryuk-Update-Log-${Date.now()}.txt`)
      
      const result = await dialog.showSaveDialog(BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0], {
        title: 'Зберегти лог оновлення',
        defaultPath,
        filters: [
          { name: 'Text Files', extensions: ['txt'] },
          { name: 'All Files', extensions: ['*'] }
        ]
      })
      
      if (!result.canceled && result.filePath) {
        fs.writeFileSync(result.filePath, content, 'utf8')
        return true
      }
      
      return false
    } catch (error) {
      console.error('Помилка збереження лог файлу:', error)
      return false
    }
  })

  // Перезапуск додатка
  ipcMain.handle('updates:restart-app', async () => {
    try {
      app.relaunch()
      app.exit(0)
    } catch (error) {
      console.error('Помилка перезапуску:', error)
    }
  })

  // Пересилання подій оновлення до рендера
  updateService.on('state-changed', (state) => {
    BrowserWindow.getAllWindows().forEach(window => {
      window.webContents.send('updates:state-changed', state)
    })
  })

  updateService.on('download-progress', (progress) => {
    BrowserWindow.getAllWindows().forEach(window => {
      window.webContents.send('updates:download-progress', progress)
    })
  })

  // Додаткові події для завершення оновлення
  updateService.on('update-complete', () => {
    BrowserWindow.getAllWindows().forEach(window => {
      window.webContents.send('updates:complete')
    })
  })

  updateService.on('update-error', (error) => {
    BrowserWindow.getAllWindows().forEach(window => {
      window.webContents.send('updates:error', error)
    })
  })
}

// Налаштування обробників пакетної обробки
function setupBatchProcessing() {
  const { BatchProcessor } = require('./services/batch/BatchProcessor')
  let batchProcessor: any = null

  // Запуск пакетної обробки
  ipcMain.handle('batch:process', async (_, options) => {
    try {
      if (batchProcessor && batchProcessor.isRunning()) {
        throw new Error('Пакетна обробка вже виконується')
      }

      batchProcessor = new BatchProcessor()
      
      // Пересилання подій прогресу
      batchProcessor.on('progress', (progress: any) => {
        BrowserWindow.getAllWindows().forEach(window => {
          window.webContents.send('batch:progress', progress)
        })
      })

      // Пересилання подій логування
      batchProcessor.on('log', (level: string, message: string) => {
        BrowserWindow.getAllWindows().forEach(window => {
          window.webContents.send('batch:log', { level, message })
        })
      })

      // Пересилання події завершення
      batchProcessor.on('complete', (result: any) => {
        BrowserWindow.getAllWindows().forEach(window => {
          window.webContents.send('batch:complete', result)
        })
      })

      return await batchProcessor.processDirectory(options)
    } catch (error) {
      console.error('Помилка пакетної обробки:', error)
      throw error
    }
  })

  // Перевірка стану пакетної обробки
  ipcMain.handle('batch:is-running', () => {
    return batchProcessor ? batchProcessor.isRunning() : false
  })

  // Скасування пакетної обробки
  ipcMain.handle('batch:cancel', () => {
    if (batchProcessor) {
      batchProcessor.cancel()
      return true
    }
    return false
  })

  // Вибір директорії
  ipcMain.handle('batch:select-directory', async () => {
    const { dialog } = require('electron')
    const result = await dialog.showOpenDialog({
      properties: ['openDirectory'],
      title: 'Оберіть директорію з Excel файлами'
    })
    return result.canceled ? null : result.filePaths[0]
  })

  // Вибір файлу для збереження
  ipcMain.handle('batch:select-output-file', async () => {
    const { dialog } = require('electron')
    const result = await dialog.showSaveDialog({
      title: 'Оберіть місце збереження результату',
      defaultPath: 'Індекс_бійців.xlsx',
      filters: [
        { name: 'Excel Files', extensions: ['xlsx'] }
      ]
    })
    return result.canceled ? null : result.filePath
  })

  // Вибір Excel файлу з іменами
  ipcMain.handle('batch:select-excel-file', async () => {
    const { dialog } = require('electron')
    const result = await dialog.showOpenDialog({
      title: 'Оберіть Excel файл з іменами',
      filters: [
        { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
        { name: 'All Files', extensions: ['*'] }
      ],
      properties: ['openFile']
    })
    return result.canceled ? null : result.filePaths[0]
  })
}

app.whenReady().then(async () => {
  storage = createStorage()
  updateService = new UpdateService(storage)
  
  // Даємо час storage ініціалізуватись та ініціалізуємо ліцензію
  setTimeout(async () => {
    await updateService.initializeLicense()
  }, 200)
  
  setupUpdateHandlers()
  setupBatchProcessing()
  setupOSIntegration(); 
  const window = createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

// IPC handlers
ipcMain.handle('os:notify', (_e, { title, body }) => notify(title, body))
ipcMain.handle('os:openExternal', (_e, { url }) => openExternal(url))
ipcMain.handle('os:powerBlocker', (_e, { enable }) => togglePowerSaveBlocker(!!enable))

ipcMain.handle('storage:getSetting', (_e, { key, fallback }) => storage.getSetting(key, fallback))
ipcMain.handle('storage:setSetting', (_e, { key, value }) => storage.setSetting(key, value))
ipcMain.handle('storage:addNote', (_e, { text }) => storage.addNote(text))
ipcMain.handle('storage:listNotes', () => storage.listNotes())

ipcMain.handle('dialog:save', async (e, { suggestName }) => {
  const result = await dialog.showSaveDialog(BrowserWindow.fromWebContents(e.sender)!, {
    defaultPath: suggestName || 'result.docx',
    filters: [
      { name: 'Word Documents', extensions: ['docx'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  })
  
  return result.canceled ? null : result.filePath
})

// Простий пошук за ключовими словами
function findParagraphsByKeyword(paragraphs: string[], keyword: string): string[] {
  const normalizedKeyword = keyword.toLowerCase()
  const matched: string[] = []
  
  console.log(`[findParagraphsByKeyword] Шукаю "${keyword}" (нормалізований: "${normalizedKeyword}") в ${paragraphs.length} абзацах`)
  
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i]
    const normalizedParagraph = norm(paragraph)
    
    if (normalizedParagraph.includes(normalizedKeyword)) {
      matched.push(paragraph)
      console.log(`[findParagraphsByKeyword] Знайдено збіг #${matched.length} в абзаці ${i + 1}: "${paragraph.substring(0, 150)}..."`)
    }
  }
  
  console.log(`[findParagraphsByKeyword] Загалом знайдено ${matched.length} збігів для "${keyword}"`)
  return matched
}

async function extractParagraphsFromWord(wordBuf: ArrayBuffer): Promise<string[]> {
  try {
    const result = await mammoth.convertToHtml({ 
      buffer: Buffer.from(wordBuf)
    })
    
    // Розбити HTML на абзаци
    const paragraphs = result.value
      .split(/<\/?p[^>]*>/i)
      .map(p => p.replace(/<[^>]+>/g, '').trim()) // Видалити HTML теги
      .filter(p => p.length > 0) // Залишити тільки непорожні
    
    return paragraphs
  } catch (err) {
    throw new Error(`Помилка читання Word: ${err instanceof Error ? err.message : String(err)}`)
  }
}

// Нова функція для витягування форматованих абзаців
async function extractFormattedParagraphsFromWord(wordBuf: ArrayBuffer): Promise<{
  paragraphs: Array<{ text: string, html: string }>,
  firstLine: string
}> {
  try {
    const result = await mammoth.convertToHtml({ 
      buffer: Buffer.from(wordBuf)
    })
    
    // Розбити HTML на абзаци, зберігаючи HTML форматування
    const htmlParagraphs = result.value.split(/<\/?p[^>]*>/i).filter(p => p.trim().length > 0)
    
    const paragraphs = htmlParagraphs.map(htmlPara => ({
      text: htmlPara.replace(/<[^>]+>/g, '').trim(),
      html: htmlPara.trim()
    })).filter(p => p.text.length > 0)
    
    // Перша строка (перший абзац)
    const firstLine = paragraphs.length > 0 ? paragraphs[0].text : ''
    
    return { paragraphs, firstLine }
  } catch (err) {
    throw new Error(`Помилка читання Word: ${err instanceof Error ? err.message : String(err)}`)
  }
}

// Типи для структури наказу
interface OrderItem {
  type: 'point' | 'subpoint' | 'subsubpoint' | 'paragraph'
  number?: string  // "1", "7.1", "8.3" тощо
  text: string
  html: string
  index: number    // оригінальний індекс в документі
  children: OrderItem[]
  parent?: OrderItem
}

// Функція для розбору структури наказу з пунктами та підпунктами
function parseOrderStructure(paragraphs: Array<{ text: string, html: string }>): OrderItem[] {
  const structure: OrderItem[] = []
  let currentPoint: OrderItem | null = null
  let currentSubpoint: OrderItem | null = null
  
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i]
    const text = para.text.trim()
    const html = para.html.trim()
    
    // Розпізнавання основних пунктів (1, 2, 3, ... або 1., 2., 3., ...)
    // Покращений regex для точнішого розпізнавання
    const pointMatch = text.match(/^(\d+)\.?\s+(.+)/) && !text.match(/^\d+\.\d+/)
    if (pointMatch) {
      const matches = text.match(/^(\d+)\.?\s+(.+)/)
      const pointNumber = matches![1]
      const pointText = matches![2]
      
      console.log(`[parseOrderStructure] Знайдено пункт ${pointNumber}: "${pointText.substring(0, 50)}..."`)
      
      currentPoint = {
        type: 'point',
        number: pointNumber,
        text: text,
        html: html,
        index: i,
        children: []
      }
      structure.push(currentPoint)
      currentSubpoint = null
      continue
    }
    
    // Розпізнавання підпунктів (7.1, 8.3, ... або 7.1., 8.3., ...)
    const subpointMatch = text.match(/^(\d+\.\d+)\.?\s+(.+)/)
    if (subpointMatch) {
      const subpointNumber = subpointMatch[1]
      const subpointText = subpointMatch[2]
      
      console.log(`[parseOrderStructure] Знайдено підпункт ${subpointNumber}: "${subpointText.substring(0, 50)}..."`)
      
      currentSubpoint = {
        type: 'subpoint',
        number: subpointNumber,
        text: text,
        html: html,
        index: i,
        children: [],
        parent: currentPoint || undefined
      }
      
      if (currentPoint) {
        currentPoint.children.push(currentSubpoint)
      } else {
        // Якщо підпункт без основного пункту, створюємо окремий елемент
        structure.push(currentSubpoint)
      }
      continue
    }
    
    // Розпізнавання підпунктів підпунктів (жирний підкреслений текст)
    const hasStrongUnderline = html.includes('<strong>') && html.includes('<u>')
    if (hasStrongUnderline) {
      const subsubpoint: OrderItem = {
        type: 'subsubpoint',
        text: text,
        html: html,
        index: i,
        children: [],
        parent: currentSubpoint || currentPoint || undefined
      }
      
      if (currentSubpoint) {
        currentSubpoint.children.push(subsubpoint)
      } else if (currentPoint) {
        currentPoint.children.push(subsubpoint)
      } else {
        structure.push(subsubpoint)
      }
      continue
    }
    
    // Звичайні абзаци
    const paragraph: OrderItem = {
      type: 'paragraph',
      text: text,
      html: html,
      index: i,
      children: [],
      parent: currentSubpoint || currentPoint || undefined
    }
    
    if (currentSubpoint) {
      currentSubpoint.children.push(paragraph)
    } else if (currentPoint) {
      currentPoint.children.push(paragraph)
    } else {
      structure.push(paragraph)
    }
  }
  
  return structure
}

// Функція для пошуку в структурі з контекстом
function findInOrderStructure(structure: OrderItem[], keyword: string): OrderItem[] {
  const results: OrderItem[] = []
  const normalizedKeyword = keyword.toLowerCase()
  
  function searchRecursive(items: OrderItem[], parentContext: OrderItem[] = []): void {
    for (const item of items) {
      const normalizedText = norm(item.text)
      
      if (normalizedText.includes(normalizedKeyword)) {
        // Знайшли збіг - зберігаємо весь контекст
        const fullContext = [...parentContext]
        
        // Додаємо батьківський пункт якщо є
        if (item.parent && !fullContext.includes(item.parent)) {
          fullContext.push(item.parent)
        }
        
        // Додаємо сам елемент
        if (!results.includes(item)) {
          results.push(item)
        }
        
        // Додаємо всі елементи контексту
        for (const contextItem of fullContext) {
          if (!results.includes(contextItem)) {
            results.push(contextItem)
          }
        }
      }
      
      // Рекурсивний пошук в дочірніх елементах
      if (item.children.length > 0) {
        const newContext = [...parentContext]
        if (item.type === 'point' || item.type === 'subpoint') {
          newContext.push(item)
        }
        searchRecursive(item.children, newContext)
      }
    }
  }
  
  searchRecursive(structure)
  
  // Сортуємо результати за оригінальним порядком в документі
  return results.sort((a, b) => a.index - b.index)
}

// Функція для розширеної обробки форматованих параграфів (з HTML розміткою)
async function createFormattedResultDocument(
  matchedParagraphs: Array<{ text: string, html: string }>, 
  outputPath: string, 
  firstLine?: string
): Promise<void> {
  const children = []
  
  // Додати першу строку наказу, якщо є (завжди шрифт Calibri)
  if (firstLine) {
    children.push(new Paragraph({ 
      children: [new TextRun({
        text: firstLine,
        font: {
          name: 'Calibri'
        },
        size: 28, // 12pt
        bold: false
      })],
      spacing: { after: 200 } // Відступ після першої строки
    }))
  }
  
  // Додати знайдені абзаци з збереженням оригінального форматування + зміна на Calibri
  if (matchedParagraphs.length > 0) {
    for (const para of matchedParagraphs) {
      // Створюємо простий абзац з текстом (Calibri шрифт)
      // TODO: В майбутньому можна додати парсинг HTML для збереження жирного/курсивного тексту
      children.push(new Paragraph({ 
        children: [new TextRun({
          text: para.text,
          font: {
            name: 'Calibri'
          },
          size: 28 // 14pt
        })]
      }))
    }
  } else {
    children.push(new Paragraph({ 
      children: [new TextRun({
        text: '— Нічого не знайдено за вказаними ключовими словами —',
        font: {
          name: 'Calibri'
        },
        size: 28,
        italics: true
      })]
    }))
  }
  
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: {
              name: 'Calibri'
            },
            size: 28 // 14pt
          }
        }
      }
    },
    sections: [{
      properties: {},
      children: children
    }]
  })
  
  const buffer = await Packer.toBuffer(doc)
  
  // Створити директорію якщо потрібно
  const dir = path.dirname(outputPath)
  await fs.mkdir(dir, { recursive: true })
  
  await fs.writeFile(outputPath, buffer)
}

// Функція для створення структурованого документу з пунктами та підпунктами
async function createStructuredResultDocument(
  matchedItems: OrderItem[], 
  outputPath: string, 
  firstLine?: string
): Promise<void> {
  const children = []
  
  // Додати першу строку наказу, якщо є
  if (firstLine) {
    children.push(new Paragraph({ 
      children: [new TextRun({
        text: firstLine,
        font: { name: 'Calibri' },
        size: 28 // 14pt
      })],
      alignment: 'both', // Вирівнювання за шириною
      indent: {
        firstLine: 720 // Абзацний відступ (0.5 дюйма)
      },
      spacing: { after: 200 }
    }))
  }
  
  // Функція для конвертації OrderItem в Paragraph БЕЗ ЖОДНИХ ВІДСТУПІВ
  function createParagraphFromItem(item: OrderItem): Paragraph {
    // Формування тексту з номером пункту/підпункту
    let displayText = item.text
    
    if (item.number && (item.type === 'point' || item.type === 'subpoint')) {
      // Видалити старий номер та додати новий
      const cleanText = item.text.replace(/^\d+(\.\d+)?\.?\s*/, '')
      displayText = `${item.number}. ${cleanText}`
    }
    
    // Простий TextRun без складного HTML парсингу
    const textRun = new TextRun({
      text: displayText,
      font: "Calibri",
      size: 28, // 14pt = 28 half-points
      bold: item.type === 'point' || item.type === 'subpoint' || item.type === 'subsubpoint'
    })
    
    // БЕЗ ЖОДНИХ ВІДСТУПІВ - все по лівому краю
    return new Paragraph({
      children: [textRun],
      alignment: 'both', // Вирівнювання за шириною
      indent: {
        firstLine: 720 // Абзацний відступ (0.5 дюйма)
      }
      // Ніяких spacing - тільки текст
    })
  }
  
  // Додати знайдені елементи або повідомлення про відсутність результатів
  if (matchedItems.length > 0) {
    for (let i = 0; i < matchedItems.length; i++) {
      const item = matchedItems[i]
      const nextItem = i < matchedItems.length - 1 ? matchedItems[i + 1] : null
      
      // Додати пустий рядок перед пунктами та підпунктами
      if (item.type === 'point' || item.type === 'subpoint') {
        children.push(new Paragraph({
          children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
          alignment: 'both', // Вирівнювання за шириною
          spacing: { after: 0 }
        }))
      }
      
      // Додати основний абзац
      children.push(createParagraphFromItem(item))
      
      // Додати пустий рядок після пунктів та підпунктів, 
      // але НЕ додавати, якщо наступний елемент - підпункт для цього пункту
      if (item.type === 'point' || item.type === 'subpoint') {
        const shouldAddEmptyLine = !(
          item.type === 'point' && 
          nextItem && 
          nextItem.type === 'subpoint' && 
          nextItem.parent === item
        )
        
        if (shouldAddEmptyLine) {
          children.push(new Paragraph({
            children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
            alignment: 'both', // Вирівнювання за шириною
            spacing: { after: 0 }
          }))
        }
      }
    }
  } else {
    children.push(new Paragraph({ 
      children: [new TextRun({
        text: '— Нічого не знайдено за вказаними ключовими словами —',
        font: { name: 'Calibri' },
        size: 28,
        italics: true
      })],
      alignment: 'both', // Вирівнювання за шириною
      indent: {
        firstLine: 720 // Абзацний відступ
      }
    }))
  }
  
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: { name: 'Calibri' },
            size: 28 // 14pt
          }
        }
      }
    },
    sections: [{
      properties: {},
      children: children
    }]
  })
  
  const buffer = await Packer.toBuffer(doc)
  
  // Створити директорію якщо потрібно
  const dir = path.dirname(outputPath)
  await fs.mkdir(dir, { recursive: true })
  
  await fs.writeFile(outputPath, buffer)
}

ipcMain.handle('order:process', async (e, payload) => {
  try {
    console.log('[order:process] starting...', {
      hasWordBuf: !!payload.wordBuf,
      outputPath: payload.outputPath,
      flags: payload.flags,
      mode: payload.mode || 'default'
    })
    
    // 1. Валідація
    if (!payload.wordBuf) {
      return { ok: false, error: 'Word-шаблон відсутній' }
    }
    
    if (!payload.outputPath) {
      return { ok: false, error: 'Шлях збереження відсутній' }
    }
    
    // 2. Обробка з підтримкою різних типів документів
    if (payload.mode === 'tokens') {
      const results: Array<{type: string, path: string, stats: any}> = []
      
      // Витягнути форматовані абзаци з Word (для всіх режимів)
      const { paragraphs: formattedParagraphs, firstLine } = await extractFormattedParagraphsFromWord(payload.wordBuf)
      
      // Розбір структури наказу на пункти та підпункти
      const orderStructure = parseOrderStructure(formattedParagraphs)
      console.log(`[order:process] Розібрано структуру наказу: ${orderStructure.length} основних елементів`)
      
      const paragraphs = formattedParagraphs.map(p => p.text) // Отримати тільки текст для зворотної сумісності
      
      console.log(`[order:process] Знайдено абзаців у Word: ${paragraphs.length}`)
      console.log(`[order:process] Перша строка наказу: "${firstLine}"`)
      console.log(`[order:process] Перші 3 абзаци:`, paragraphs.slice(0, 3))
      
      // 2БСП режим - пошук за ключовим словом з контекстом структури
      if (payload.flags.is2BSP) {
        try {
          console.log('[order:process] Режим 2БСП: пошук за ключовим словом "2-го батальйону" з збереженням структури...')
        
          // Знайти збіги в структурі наказу з контекстом
          const matchedItems = findInOrderStructure(orderStructure, "2-го батальйону")
          console.log(`[order:process] Збігів знайдено за ключовим словом "2-го батальйону": ${matchedItems.length}`)
          
          // Показати перші кілька знайдених збігів з їх типами
          if (matchedItems.length > 0) {
            console.log('[order:process] Перші 3 знайдені елементи з контекстом:')
            for (let i = 0; i < Math.min(3, matchedItems.length); i++) {
              const item = matchedItems[i]
              console.log(`[order:process] Збіг #${i+1} [${item.type}${item.number ? ` ${item.number}` : ''}]: "${item.text.substring(0, 100)}..."`)
            }
          }
          
          // Додати 2БСП результат до списку
          results.push({
            type: '2БСП',
            path: payload.outputPath.replace('.docx', '_2БСП.docx'),
            stats: {
              tokens: 1,
              paragraphs: formattedParagraphs.length,
              matched: matchedItems.length,
              structureElements: orderStructure.length
            }
          })
          
          // Створити 2БСП документ з першою строкою та структурою
          const bspPath = payload.outputPath.replace('.docx', '_2БСП.docx')
          await createStructuredResultDocument(matchedItems, bspPath, firstLine)
          
        } catch (err: any) {
          console.error('[order:process] 2БСП processing error:', err)
          return { ok: false, error: `2БСП обробка: ${err instanceof Error ? err.message : String(err)}` }
        }
      }
      
      // Розпорядження режим - пошук за ключовим словом "розпорядженні"
      if (payload.flags.isOrder) {
        try {
          console.log('[order:process] Режим Розпорядження: пошук за ключовим словом "розпорядженні" з збереженням структури...')
          
          // Знайти збіги в структурі наказу з контекстом
          const orderMatchedItems = findInOrderStructure(orderStructure, "розпорядженні")
          console.log(`[order:process] Збігів знайдено за ключовим словом "розпорядженні": ${orderMatchedItems.length}`)
          
          // Показати перші кілька знайдених збігів з їх типами
          if (orderMatchedItems.length > 0) {
            console.log('[order:process] Перші 3 знайдені елементи з контекстом (розпорядження):')
            for (let i = 0; i < Math.min(3, orderMatchedItems.length); i++) {
              const item = orderMatchedItems[i]
              console.log(`[order:process] Збіг #${i+1} [${item.type}${item.number ? ` ${item.number}` : ''}]: "${item.text.substring(0, 100)}..."`)
            }
          }
          
          // Додати результат розпорядження до списку
          results.push({
            type: 'Розпорядження',
            path: payload.outputPath.replace('.docx', '_Розпорядження.docx'),
            stats: {
              tokens: 1,
              paragraphs: formattedParagraphs.length,
              matched: orderMatchedItems.length,
              structureElements: orderStructure.length
            }
          })
          
          // Створити документ розпорядження з першою строкою та структурою
          const orderPath = payload.outputPath.replace('.docx', '_Розпорядження.docx')
          await createStructuredResultDocument(orderMatchedItems, orderPath, firstLine)
          
        } catch (err: any) {
          console.error('[order:process] Розпорядження processing error:', err)
          return { ok: false, error: `Розпорядження обробка: ${err instanceof Error ? err.message : String(err)}` }
        }
      }
      
      // Перевірка що хоча б один режим включений
      if (!payload.flags.is2BSP && !payload.flags.isOrder) {
        return { ok: false, error: 'Оберіть хоча б один режим: 2БСП або Розпорядження' }
      }
      
      // Підсумок створених документів
      console.log(`[order:process] === ПІДСУМОК ===`)
      console.log(`[order:process] Створено документів: ${results.length}`)
      results.forEach((result, index) => {
        console.log(`[order:process] ${index + 1}. ${result.type}: ${result.path} (${result.stats.matched} збігів)`)
      })
      
      // Автовідкриття (якщо потрібно) - відкрити ВСІ створені документи
      if (payload.flags.autoOpen && results.length > 0) {
        try {
          for (let i = 0; i < results.length; i++) {
            // Додаємо затримку між відкриттями, щоб не перевантажити систему
            setTimeout(() => {
              shell.openPath(results[i].path)
              console.log(`[order:process] Auto-opening document ${i + 1}/${results.length}:`, results[i].path)
            }, 500 + i * 200) // 500ms для першого, 700ms для другого і т.д.
          }
        } catch (err) {
          console.error('[order:process] Auto-open error:', err)
        }
      }
      
      return {
        ok: true,
        out: results.length > 0 ? results.map(r => r.path).join(', ') : payload.outputPath,
        stats: results.length > 0 ? {
          totalDocuments: results.length,
          documents: results.map(r => ({ type: r.type, matched: r.stats.matched }))
        } : { tokens: 0, paragraphs: 0, matched: 0 },
        results: results
      }
    }
    
    // Якщо не вибрано жодного режиму
    return { ok: false, error: 'Оберіть хоча б один режим: 2БСП або Розпорядження' }
    
  } catch (err) {
    console.error('[order:process] Unexpected error:', err)
    return { ok: false, error: `Несподівана помилка: ${err instanceof Error ? err.message : String(err)}` }
  }
})

console.log('[main] dialog:save handler ready')
