import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron'
import path from 'node:path'
import fs from 'node:fs/promises'
import * as mammoth from 'mammoth'
import { Document, Paragraph, Packer, TextRun } from 'docx'
import * as xlsx from 'xlsx'

import { setupOSIntegration, notify, openExternal, togglePowerSaveBlocker } from './services/osIntegration'
import { createStorage } from './services/storage'
import { UpdateService, UpdateState } from './services/updateService'
import { UkrainianNameDeclension } from './services/UkrainianNameDeclension'

console.log('\n\n🌟🌟🌟🌟🌟 MAIN.TS ФАЙЛ ЗАВАНТАЖЕНО - ВЕРСІЯ 17.10.2025-15:00 🌟🌟🌟🌟🌟\n')
console.log('📁 Поточний файл:', __filename)
console.log('📂 Директорія:', __dirname)

const isDev = process.env.NODE_ENV !== 'production' && (process.env.VITE_DEV_SERVER_URL !== undefined || process.argv.includes('--dev'))

let storage: any
let updateService: UpdateService
let mainWindow: BrowserWindow | null = null

// Функція для відправки логів в renderer process
function sendLog(level: 'info' | 'warn' | 'error', message: string) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('main:log', level, message)
  }
}

// Перехоплюємо console.log/warn/error для відправки в UI
const originalLog = console.log
const originalWarn = console.warn
const originalError = console.error

console.log = (...args: any[]) => {
  originalLog.apply(console, args)
  const message = args.map(arg => {
    if (typeof arg === 'object') {
      try { return JSON.stringify(arg, null, 2) }
      catch { return String(arg) }
    }
    return String(arg)
  }).join(' ')
  sendLog('info', message)
}

console.warn = (...args: any[]) => {
  originalWarn.apply(console, args)
  const message = args.map(arg => String(arg)).join(' ')
  sendLog('warn', message)
}

console.error = (...args: any[]) => {
  originalError.apply(console, args)
  const message = args.map(arg => String(arg)).join(' ')
  sendLog('error', message)
}

function norm(str: string): string {
  return str.toLowerCase()
    .replace(/[«»"""''`]/g, '"')
    .replace(/[—–−]/g, '-')
    .replace(/…/g, '...')
    .replace(/\s+/g, ' ')
    .trim()
}

function createWindow(): BrowserWindow {
  mainWindow = new BrowserWindow({
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
    
    // Відкриваємо DevTools з налаштуваннями
    mainWindow.webContents.openDevTools({ mode: 'detach' })
    
    // FIXED: Вимикаємо Autofill в DevTools для уникнення помилок
    mainWindow.webContents.on('devtools-opened', () => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.devToolsWebContents?.executeJavaScript(`
          // Приховуємо помилки Autofill у консолі
          const originalError = console.error;
          console.error = function(...args) {
            const msg = args[0]?.toString() || '';
            if (msg.includes('Autofill')) {
              return; // Ігноруємо помилки Autofill
            }
            originalError.apply(console, args);
          };
        `).catch(() => {
          // Ігноруємо помилки виконання скрипта
        });
      }
    });
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'))
  }

  // Додаємо можливість відкрити DevTools в продакшн режимі (F12)
  mainWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'F12' && mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.toggleDevTools()
    }
  })
  
  return mainWindow
}

function setupUpdateHandlers() {
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

  // Сканування Excel файлів у папці
  ipcMain.handle('batch:scan-excel-files', async (event, folderPath: string) => {
    const fs = require('fs')
    const path = require('path')
    
    try {
      if (!fs.existsSync(folderPath)) {
        return []
      }
      
      const files = fs.readdirSync(folderPath)
      const excelFiles = files.filter((file: string) => {
        const ext = path.extname(file).toLowerCase()
        return ext === '.xlsx' || ext === '.xls'
      })
      
      return excelFiles
    } catch (error) {
      console.error('Помилка сканування файлів:', error)
      return []
    }
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

ipcMain.handle('dialog:select-folder', async (e) => {
  const result = await dialog.showOpenDialog(BrowserWindow.fromWebContents(e.sender)!, {
    properties: ['openDirectory']
  })
  
  return result.canceled ? null : { filePath: result.filePaths[0] }
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
    
    const htmlParagraphs = result.value.split(/<\/?p[^>]*>/i).filter(p => p.trim().length > 0)
    
    const paragraphs = htmlParagraphs.map((htmlPara) => {
      const text = htmlPara.replace(/<[^>]+>/g, '').trim()
      const html = htmlPara.trim()
      
      return { text, html }
    }).filter(p => p.text.length > 0)
    
    const firstLine = paragraphs.length > 0 ? paragraphs[0].text : ''
    
    console.log(`[extractFormatted] Загалом абзаців: ${paragraphs.length}`)
    
    return { paragraphs, firstLine }
  } catch (err) {
    throw new Error(`Помилка читання Word: ${err instanceof Error ? err.message : String(err)}`)
  }
}

// Типи для структури наказу
interface OrderItem {
  type: 'point' | 'subpoint' | 'dash-point' | 'paragraph'
  number?: string  // "1", "7.1", "8.3" тощо
  text: string
  html: string
  index: number    // оригінальний індекс в документі
  children: OrderItem[]
  parent?: OrderItem
  matchedNames?: string[] // ПІБ знайдені в цьому елементі
}

// ============================================================================
// РОЗПІЗНАВАННЯ ШТРИХПУНКТУ ЗА ВІЙСЬКОВИМИ ЗВАННЯМИ
// ============================================================================

function isDashPointByPattern(text: string): boolean {
  // Очищуємо текст від зайвих пробілів
  const cleanText = text.trim().toLowerCase()
  
  // === СПИСОК ДОЗВОЛЕНИХ ВІЙСЬКОВИХ ЗВАНЬ ===
  const allowedRanks = [
    'солдат',
    'старший солдат',
    'молодший сержант',
    'сержант',
    'старший сержант',
    'головний сержант',
    'штаб-сержант',
    'капітан',
    'майор',
    'молодший лейтенант',
    'лейтенант',
    'старший лейтенант'
  ]
  
  // Перевіряємо точну відповідність (можливо з тире в кінці)
  for (const rank of allowedRanks) {
    // Варіант 1: Точно як у списку
    if (cleanText === rank) {
      console.log(`[isDashPoint] ✅ Знайдено звання: "${text}"`)
      return true
    }
    
    // Варіант 2: Із тире або пробілом і тире в кінці
    if (cleanText === `${rank} -` || cleanText === `${rank}-`) {
      console.log(`[isDashPoint] ✅ Знайдено звання з тире: "${text}"`)
      return true
    }
    
    // Варіант 3: Із двокрапкою в кінці
    if (cleanText === `${rank}:` || cleanText === `${rank} :`) {
      console.log(`[isDashPoint] ✅ Знайдено звання з двокрапкою: "${text}"`)
      return true
    }
  }
  
  return false
}

// Функція для розбору структури наказу з пунктами та підпунктами
function parseOrderStructure(paragraphs: Array<{ text: string, html: string }>): OrderItem[] {
  const structure: OrderItem[] = []
  let currentPoint: OrderItem | null = null
  let currentSubpoint: OrderItem | null = null
  let currentDashPoint: OrderItem | null = null
  
  console.log('[parseOrderStructure] Початок розбору структури...\n')
  
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i]
    const text = para.text.trim()
    const html = para.html.trim()
    
    // Розпізнавання основних пунктів (1, 2, 3, ... або 1., 2., 3., ...)
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
      currentDashPoint = null
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
        structure.push(currentSubpoint)
      }
      currentDashPoint = null
      continue
    }
    
    // === 3. ПЕРЕВІРКА НА ШТРИХПУНКТ ЗА ВІЙСЬКОВИМ ЗВАННЯМ ===
    if (isDashPointByPattern(text)) {
      console.log(`[parseOrderStructure] ⭐⭐⭐ ШтрихПункт (звання) на позиції ${i}: "${text}"`)
      
      const dashpoint: OrderItem = {
        type: 'dash-point',
        text: text,
        html: html,
        index: i,
        children: [],
        parent: currentSubpoint || currentPoint || undefined
      }
      
      if (currentSubpoint) {
        currentSubpoint.children.push(dashpoint)
        console.log(`[parseOrderStructure]    → Додано до підпункту ${currentSubpoint.number}`)
      } else if (currentPoint) {
        currentPoint.children.push(dashpoint)
        console.log(`[parseOrderStructure]    → Додано до пункту ${currentPoint.number}`)
      } else {
        structure.push(dashpoint)
        console.log(`[parseOrderStructure]    → УВАГА: Додано до кореня`)
      }
      
      currentDashPoint = dashpoint
      continue
    }
    
    // Звичайні абзаци
    const paragraph: OrderItem = {
      type: 'paragraph',
      text: text,
      html: html,
      index: i,
      children: [],
      parent: currentDashPoint || currentSubpoint || currentPoint || undefined
    }
    
    if (currentDashPoint) {
      currentDashPoint.children.push(paragraph)
    } else if (currentSubpoint) {
      currentSubpoint.children.push(paragraph)
    } else if (currentPoint) {
      currentPoint.children.push(paragraph)
    } else {
      structure.push(paragraph)
    }
  }
  
  // === ДІАГНОСТИЧНИЙ ВИВІД СТРУКТУРИ ===
  console.log('\n[parseOrderStructure] === СТРУКТУРА ДОКУМЕНТА ===')
  function printStructure(items: OrderItem[], depth = 0) {
    for (const item of items) {
      const indent = '  '.repeat(depth)
      const icon = item.type === 'dash-point' ? '⭐⭐⭐' : 
                   item.type === 'point' ? '📌' :
                   item.type === 'subpoint' ? '📍' : '📝'
      console.log(`${indent}${icon} ${item.type}${item.number ? ` ${item.number}` : ''}: "${item.text.substring(0, 50)}..." (idx: ${item.index})`)
      if (item.children.length > 0) {
        printStructure(item.children, depth + 1)
      }
    }
  }
  printStructure(structure)
  console.log('[parseOrderStructure] === КІНЕЦЬ СТРУКТУРИ ===\n')
  
  // Підрахунок ШтрихПунктів
  function countDashPoints(items: OrderItem[]): number {
    let count = 0
    for (const item of items) {
      if (item.type === 'dash-point') count++
      count += countDashPoints(item.children)
    }
    return count
  }
  
  const dashPointCount = countDashPoints(structure)
  console.log(`[parseOrderStructure] ЗАГАЛОМ ШТРИХПУНКТІВ У СТРУКТУРІ: ${dashPointCount}\n`)
  
  return structure
}

// Функція для читання Excel файлу та отримання ПІБ з колонки D
async function readExcelColumnD(filePath: string): Promise<string[]> {
  try {
    const data = await fs.readFile(filePath)
    const workbook = xlsx.read(data, { type: 'buffer' })
    const sheetName = workbook.SheetNames[0] // Перший аркуш
    const sheet = workbook.Sheets[sheetName]
    
    const names: string[] = []
    let row = 1 // Починаємо з першого рядка
    
    while (true) {
      const cellAddress = `D${row}` // Колонка D
      const cell = sheet[cellAddress]
      
      if (!cell || !cell.v) break // Якщо комірка пуста, зупиняємося
      
      const value = String(cell.v).trim()
      if (value) {
        names.push(value)
      }
      
      row++
    }
    
    console.log(`[Excel] Зчитано ${names.length} ПІБ з колонки D:`, names.slice(0, 3)) // Показуємо перші 3
    return names
  } catch (error) {
    console.error('[Excel] Помилка читання файлу:', error)
    return []
  }
}

// Функція для пошуку в структурі з контекстом
function findInOrderStructure(structure: OrderItem[], keyword: string): OrderItem[] {
  const results: OrderItem[] = []
  const addedIndices = new Set<number>()  // Для уникнення дублів
  const normalizedKeyword = keyword.toLowerCase()
  
  const norm = (text: string) => text.toLowerCase()
  
  // === СТАТИСТИКА ДЛЯ ДІАГНОСТИКИ ===
  let foundParagraphs = 0
  let foundPoints = 0
  let foundSubpoints = 0
  let foundDashPoints = 0
  
  function addWithHierarchy(item: OrderItem): void {
    // Функція для додавання елемента разом з усією його ієрархією батьків
    const hierarchyChain: OrderItem[] = []
    
    // Збираємо всю ієрархію від елемента до кореня
    let current: OrderItem | undefined = item
    while (current) {
      hierarchyChain.unshift(current)  // Додаємо на початок
      current = current.parent
    }
    
    // === ЛОГУВАННЯ ІЄРАРХІЇ ===
    console.log(`[findInOrderStructure] Додавання ієрархії для "${item.text.substring(0, 40)}...":`)
    for (const h of hierarchyChain) {
      console.log(`[findInOrderStructure]   ${h.type}${h.number ? ` ${h.number}` : ''}: "${h.text.substring(0, 40)}..."`)
    }
    
    // Додаємо всі елементи ієрархії, уникаючи дублів
    for (const hierarchyItem of hierarchyChain) {
      if (!addedIndices.has(hierarchyItem.index)) {
        results.push(hierarchyItem)
        addedIndices.add(hierarchyItem.index)
        
        // Підрахунок статистики
        if (hierarchyItem.type === 'dash-point') {
          foundDashPoints++
          console.log(`[findInOrderStructure]   ✅ Додано ШтрихПункт: "${hierarchyItem.text.substring(0, 40)}..."`)
        } else if (hierarchyItem.type === 'point') {
          foundPoints++
        } else if (hierarchyItem.type === 'subpoint') {
          foundSubpoints++
        } else if (hierarchyItem.type === 'paragraph') {
          foundParagraphs++
        }
      }
    }
  }
  
  function searchRecursive(items: OrderItem[]): void {
    for (const item of items) {
      const normalizedText = norm(item.text)
      
      if (normalizedText.includes(normalizedKeyword)) {
        console.log(`[findInOrderStructure] 🎯 Знайдено збіг в ${item.type}${item.number ? ` ${item.number}` : ''}: "${item.text.substring(0, 60)}..."`)
        
        if (item.parent) {
          console.log(`[findInOrderStructure]    Батько: ${item.parent.type}${item.parent.number ? ` ${item.parent.number}` : ''}: "${item.parent.text.substring(0, 40)}..."`)
        }
        
        addWithHierarchy(item)
      }
      
      // Рекурсивний пошук в дочірніх елементах
      if (item.children.length > 0) {
        searchRecursive(item.children)
      }
    }
  }
  
  searchRecursive(structure)
  
  // === ФІНАЛЬНА СТАТИСТИКА ===
  console.log(`[findInOrderStructure] === СТАТИСТИКА ПОШУКУ ===`)
  console.log(`[findInOrderStructure] Всього елементів в результаті: ${results.length}`)
  console.log(`[findInOrderStructure]   - Пунктів: ${foundPoints}`)
  console.log(`[findInOrderStructure]   - Підпунктів: ${foundSubpoints}`)
  console.log(`[findInOrderStructure]   - ШтрихПунктів: ${foundDashPoints} ⭐`)
  console.log(`[findInOrderStructure]   - Абзаців: ${foundParagraphs}`)
  
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
  
  // Функція для конвертації OrderItem в Paragraph
  function createParagraphFromItem(item: OrderItem): Paragraph {
    // Формування тексту з номером пункту/підпункту
    let displayText = item.text
    
    if (item.number && (item.type === 'point' || item.type === 'subpoint')) {
      // Видалити старий номер та додати новий
      const cleanText = item.text.replace(/^\d+(\.\d+)?\.?\s*/, '')
      displayText = `${item.number}. ${cleanText}`
    }
    
    // Визначити форматування залежно від типу
    // ШтрихПункт = тільки підкреслений (БЕЗ жирного)
    const isBold = item.type === 'point' || item.type === 'subpoint'
    const isUnderline = item.type === 'dash-point'
    
    // TextRun з правильним форматуванням
    const textRun = new TextRun({
      text: displayText,
      font: "Calibri",
      size: 28, // 14pt = 28 half-points
      bold: isBold,
      underline: isUnderline ? { type: 'single' } : undefined
    })
    
    return new Paragraph({
      children: [textRun],
      alignment: 'both', // Вирівнювання за шириною
      indent: {
        firstLine: 720 // Абзацний відступ (0.5 дюйма)
      }
    })
  }
  
  // Додати знайдені елементи або повідомлення про відсутність результатів
  if (matchedItems.length > 0) {
    for (let i = 0; i < matchedItems.length; i++) {
      const item = matchedItems[i]
      const prevItem = i > 0 ? matchedItems[i - 1] : null
      const nextItem = i < matchedItems.length - 1 ? matchedItems[i + 1] : null
      
      // Додати пустий рядок перед пунктами та підпунктами
      if (item.type === 'point' || item.type === 'subpoint') {
        children.push(new Paragraph({
          children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
          alignment: 'both',
          spacing: { after: 0 }
        }))
      }
      
      // Додати пустий рядок перед ШтрихПунктом, якщо попередній елемент був абзацем
      if (item.type === 'dash-point' && prevItem?.type === 'paragraph') {
        children.push(new Paragraph({
          children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
          alignment: 'both',
          spacing: { after: 0 }
        }))
      }
      
      // Додати основний абзац
      children.push(createParagraphFromItem(item))
      
      // Додати пустий рядок після пунктів та підпунктів
      // НЕ додавати, якщо наступний елемент - підпункт або ШтрихПункт цього пункту
      if (item.type === 'point' || item.type === 'subpoint') {
        const shouldAddEmptyLine = !(
          item.type === 'point' && 
          nextItem && 
          (nextItem.type === 'subpoint' || nextItem.type === 'dash-point') && 
          nextItem.parent === item
        )
        
        if (shouldAddEmptyLine) {
          children.push(new Paragraph({
            children: [new TextRun({ text: "", font: "Calibri", size: 28 })],
            alignment: 'both',
            spacing: { after: 0 }
          }))
        }
      }
      
      // Після ШтрихПункту НІКОЛИ не додавати пустий рядок
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
  // КРИТИЧНА ДІАГНОСТИКА #1: ТОЧКА ВХОДУ
  console.log('\n\n🚨🚨🚨 [order:process] ТОЧКА ВХОДУ - HANDLER СТАРТУВАВ 🚨🚨🚨\n')
  console.log('═'.repeat(80))
  console.log('🎯🎯🎯 [order:process] HANDLER ВИКЛИКАНО - ВЕРСІЯ 17.10.2025-15:55 🎯🎯🎯')
  console.log('═'.repeat(80))
  console.log('\n')
  
  // ДІАГНОСТИКА #2: PAYLOAD СТРУКТУРА
  console.log('📦 [order:process] Payload keys:', Object.keys(payload))
  console.log('📦 [order:process] Payload.mode:', payload.mode)
  console.log('📦 [order:process] Payload.wordBuf exists:', !!payload.wordBuf)
  console.log('📦 [order:process] Payload.outputPath:', payload.outputPath)
  
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
    
    // 2. Обробка наказу (всі режими)
    const results: Array<{type: string, path: string, stats: any}> = []
    
    console.log('\n\n🚀🚀🚀 [order:process] ПЕРЕД викликом extractFormattedParagraphsFromWord 🚀🚀🚀')
    console.log(`[order:process] wordBuf type: ${typeof payload.wordBuf}`)
    console.log(`[order:process] wordBuf constructor: ${payload.wordBuf?.constructor?.name}`)
    console.log(`[order:process] wordBuf keys:`, Object.keys(payload.wordBuf || {}).slice(0, 10))
    console.log(`[order:process] wordBuf byteLength: ${payload.wordBuf?.byteLength}`)
    console.log(`[order:process] wordBuf buffer: ${payload.wordBuf?.buffer?.byteLength}`)
    console.log(`[order:process] Is ArrayBuffer: ${payload.wordBuf instanceof ArrayBuffer}`)
    console.log(`[order:process] Is Buffer: ${Buffer.isBuffer(payload.wordBuf)}`)
    console.log(`[order:process] Is Uint8Array: ${payload.wordBuf instanceof Uint8Array}`)
    
    // Витягнути форматовані абзаци з Word (для всіх режимів)
    const { paragraphs: formattedParagraphs, firstLine } = await extractFormattedParagraphsFromWord(payload.wordBuf)
    
    console.log('\n\n✅✅✅ [order:process] ПІСЛЯ extractFormattedParagraphsFromWord ✅✅✅')
    console.log(`[order:process] Отримано paragraphs: ${formattedParagraphs?.length || 0}, firstLine: "${firstLine?.substring(0, 50) || 'undefined'}"`)
    
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
      
      // Розпорядження режим - пошук з Excel файлом і правильною логікою
      if (payload.flags.isOrder) {
        try {
          console.log('[order:process] Режим Розпорядження: читання Excel файлу та пошук з AND логікою...')
          
          // Читання Excel файлу для отримання ПІБ
          let excelNames: string[] = []
          if (payload.excelPath) {
            console.log(`[order:process] Читання Excel файлу: ${payload.excelPath}`)
            excelNames = await readExcelColumnD(payload.excelPath)
            console.log(`[order:process] Знайдено ${excelNames.length} ПІБ в Excel колонці D`)
          } else {
            console.log('[order:process] ⚠️ Excel файл не вибрано - використовуємо тільки пошук "розпорядженні"')
          }
          
          // Використання правильної логіки з UkrainianNameDeclension
          const wordText = paragraphs.join('\n\n')
          const orderResults = UkrainianNameDeclension.findOrderParagraphs(wordText, excelNames)
          
          console.log(`[order:process] Знайдено абзаців з правильною логікою: ${orderResults.length}`)
          
          // Показати перші кілька знайдених збігів
          if (orderResults.length > 0) {
            console.log('[order:process] Перші 3 знайдені абзаци з AND логікою:')
            for (let i = 0; i < Math.min(3, orderResults.length); i++) {
              const result = orderResults[i]
              console.log(`[order:process] Збіг #${i+1}: ПІБ: [${result.matchedNames.join(', ')}] - "${result.paragraph.substring(0, 100)}..."`)
            }
          }
          
          // Конвертуємо результати в формат OrderItem для сумісності
          const orderMatchedItems: OrderItem[] = orderResults.map((result, index) => ({
            type: 'paragraph' as const,
            text: result.paragraph,
            html: result.paragraph,
            index: result.startPosition,
            children: [],
            matchedNames: result.matchedNames
          }))
          
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
    
  } catch (err) {
    console.error('[order:process] Unexpected error:', err)
    return { ok: false, error: `Несподівана помилка: ${err instanceof Error ? err.message : String(err)}` }
  }
})

console.log('[main] dialog:save handler ready')

// ============================================================================
// ТЕСТУВАННЯ РОЗПІЗНАВАННЯ ВІЙСЬКОВИХ ЗВАНЬ
// ============================================================================

function testDashPointPatterns() {
  console.log('\n=== ТЕСТ РОЗПІЗНАВАННЯ ВІЙСЬКОВИХ ЗВАНЬ ===\n')
  
  const testCases = [
    // ✅ Мають розпізнатися
    { text: 'солдат', expected: true },
    { text: 'старший солдат', expected: true },
    { text: 'молодший сержант', expected: true },
    { text: 'сержант', expected: true },
    { text: 'старший сержант', expected: true },
    { text: 'головний сержант', expected: true },
    { text: 'штаб-сержант', expected: true },
    { text: 'капітан', expected: true },
    { text: 'майор', expected: true },
    { text: 'молодший лейтенант', expected: true },
    { text: 'лейтенант', expected: true },
    { text: 'старший лейтенант', expected: true },
    { text: 'молодший сержант -', expected: true },
    { text: 'сержант:', expected: true },
    { text: 'Молодший Сержант', expected: true },
    { text: 'головний сержант -', expected: true },
    
    // ❌ НЕ мають розпізнатися
    { text: '13. ОГОЛОСИТИ про присвоєння', expected: false },
    { text: '15.1. Зі складу сил', expected: false },
    { text: 'старшого лейтенанта ПЕТРЕНКА', expected: false },
    { text: 'прапорщик', expected: false },
    { text: 'молодший сержант дуже довгий текст', expected: false }
  ]
  
  let passed = 0
  let failed = 0
  
  for (const testCase of testCases) {
    const result = isDashPointByPattern(testCase.text)
    const status = result === testCase.expected ? '✅' : '❌'
    
    if (result === testCase.expected) {
      passed++
    } else {
      failed++
      console.log(`${status} FAIL: "${testCase.text}"`)
      console.log(`   Очікувалось: ${testCase.expected}, отримано: ${result}`)
    }
  }
  
  console.log(`\n=== РЕЗУЛЬТАТИ ===`)
  console.log(`✅ Пройдено: ${passed}/${testCases.length}`)
  console.log(`❌ Провалено: ${failed}/${testCases.length}`)
  console.log(`==================\n`)
}

// Розкоментуйте для тестування перед обробкою наказів:
// testDashPointPatterns()
