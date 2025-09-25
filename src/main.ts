type Route = '/functions' | '/batch' | '/notes' | '/updates' | '/settings';

function navigate(hash: string) {
  const route = (hash.replace('#', '') || '/functions') as Route;
  document.querySelectorAll<HTMLElement>('.route').forEach(s => {
    s.hidden = s.dataset.route !== route;
  });
  document.querySelectorAll('.nav a').forEach(a => {
    const href = a.getAttribute('href') || '';
    a.classList.toggle('active', href === `#${route}`);
  });
}
window.addEventListener('hashchange', () => navigate(location.hash));
navigate(location.hash);

// Load saved settings on startup
async function loadSettings() {
  try {
    // Load checkbox states
    const is2BSP = await window.api?.getSetting?.('is2BSP', true);
    const isOrder = await window.api?.getSetting?.('isOrder', false);
    const autoOpen = await window.api?.getSetting?.('autoOpen', true);
    
    // Apply saved states
    const is2BSPCheckbox = byId<HTMLInputElement>('t-2bsp');
    const isOrderCheckbox = byId<HTMLInputElement>('t-order');
    const autoOpenCheckbox = byId<HTMLInputElement>('t-autopen');
    
    if (is2BSPCheckbox) is2BSPCheckbox.checked = is2BSP;
    if (isOrderCheckbox) isOrderCheckbox.checked = isOrder;
    if (autoOpenCheckbox) autoOpenCheckbox.checked = autoOpen;
    
  } catch (err) {
    console.warn('Failed to load settings:', err);
  }
}

// Save checkbox states when they change
function setupSettingsAutoSave() {
  const checkboxes = ['t-2bsp', 't-order', 't-autopen'];
  const settingsMap = {
    't-2bsp': 'is2BSP', 
    't-order': 'isOrder',
    't-autopen': 'autoOpen'
  };
  
  checkboxes.forEach(id => {
    const checkbox = byId<HTMLInputElement>(id);
    if (checkbox) {
      checkbox.addEventListener('change', async () => {
        const setting = settingsMap[id as keyof typeof settingsMap];
        await window.api?.setSetting?.(setting, checkbox.checked);
      });
    }
  });
}

// Initialize settings on page load
(async () => {
  await loadSettings();
  setupSettingsAutoSave();
})();

const byId = <T extends HTMLElement>(id: string) => document.getElementById(id) as T | null;
const log = (msg: string) => {
  const el = byId<HTMLPreElement>('log-body');
  if (!el) return;
  const time = new Date().toLocaleTimeString();
  el.textContent += `[${time}] ${msg}\n`;
  el.scrollTop = el.scrollHeight;
};

byId<HTMLButtonElement>('btn-notify')?.addEventListener('click', () => {
  window.api?.notify?.('Привіт!', 'Це нотифікація з main-процесу');
});
byId<HTMLButtonElement>('btn-docs')?.addEventListener('click', () => {
  window.api?.openExternal?.('https://electronjs.org');
});

byId<HTMLInputElement>('word-file')?.addEventListener('change', (e) => {
  const f = (e.target as HTMLInputElement).files?.[0];
  if (f) log(`Word: ${f.name}`);
});

byId<HTMLButtonElement>('choose-result')?.addEventListener('click', async () => {
  if (!window.api?.chooseSavePath) {
    log('❌ API chooseSavePath недоступне (перевір preload.ts та перезапуск Electron).');
    return;
  }
  log('Відкриваю провідник для вибору місця збереження…');
  const timeout = new Promise<null>(res => setTimeout(()=>res(null), 12000)); // 12s
  const pick = window.api.chooseSavePath('наказ.docx');
  const result = await Promise.race([pick as Promise<any>, timeout]);
  if (result === null) {
    log('⏱️ Немає відповіді від main-процесу. Перевір DevTools → Console.');
    return;
  }
  if (result?.error) {
    log(`❌ Помилка діалогу: ${result.error}`);
    return;
  }
  if (!result) {
    log('Вибір скасовано.');
    return;
  }
  const input = byId<HTMLInputElement>('result-path');
  if (input) input.value = String(result);
  log(`✅ Обрано: ${result}`);
});

byId<HTMLButtonElement>('btn-process-order')?.addEventListener('click', async () => {
  try {
    // 1. Валідація полів
    const wordInput = byId<HTMLInputElement>('word-file');
    const outputPath = byId<HTMLInputElement>('result-path')?.value?.trim();
    
    if (!wordInput?.files?.[0]) {
      log('❌ Не вибрано Word-файл');
      return;
    }
    
    if (!outputPath) {
      log('❌ Не вказано місце збереження');
      return;
    }
    
    // 2. Читання флажків
    const flags = {
      is2BSP: byId<HTMLInputElement>('t-2bsp')?.checked || false,
      isOrder: byId<HTMLInputElement>('t-order')?.checked || false,
      tokens: true, // Завжди обробляємо токени
      autoOpen: byId<HTMLInputElement>('t-autopen')?.checked || false,
    };
    
    const flagsStr = [
      flags.is2BSP ? '2БСП' : null,
      flags.isOrder ? 'Розпорядження' : null,
      flags.autoOpen ? 'Автовідкриття' : null,
    ].filter(Boolean).join(', ');
    
    log(`🚀 Старт обробки. Вихід: ${outputPath}. Опції: ${flagsStr || '—'}`);
    log(`🔍 Режим пошуку: ключові слова в абзацах`);
    
    // 3. Читання Word файлу
    const wordFile = wordInput.files[0];
    const wordBuf = await wordFile.arrayBuffer();
    
    // 4. Виклик IPC для простої обробки (без Excel)
    const payload = {
      wordBuf,
      outputPath,
      flags,
      mode: 'tokens'
    };
    
    const result = await window.api?.processOrder?.(payload);
    
    if (!result) {
      log('❌ Помилка виклику API');
      return;
    }
    
    if (!result.ok) {
      log(`❌ ${result.error || 'Невідома помилка'}`);
      return;
    }
    
    // 6. Обробка результату з покращеною статистикою
    if ((result as any).stats) {
      const stats = (result as any).stats;
      
      if (stats.totalDocuments) {
        log(`📊 Створено документів: ${stats.totalDocuments}`);
        stats.documents.forEach((doc: any) => {
          log(`  📄 ${doc.type}: ${doc.matched} збігів`);
        });
      } else {
        log(`📊 Статистика пошуку: ${stats.paragraphs} абзаців, ${stats.matched} збігів`);
      }
    }

    log(`✅ Збережено: ${result.out}`);
    
    if (flags.autoOpen) {
      const docCount = ((result as any).stats?.totalDocuments) || 1;
      log(`🔓 Відкрито документів: ${docCount}`);
    }  } catch (err) {
    log(`❌ Помилка: ${err instanceof Error ? err.message : String(err)}`);
  }
});

async function refreshNotes() {
  const ul = byId<HTMLUListElement>('notes-list');
  if (!ul) return;
  const notes = await window.api?.listNotes?.();
  if (!notes) return;
  ul.innerHTML = notes.slice().reverse().map(n => 
    `<li><span style="opacity:.7;font-size:12px">${new Date(n.createdAt).toLocaleString()}</span> — ${n.text}</li>`
  ).join('');
}
byId<HTMLButtonElement>('note-add')?.addEventListener('click', async () => {
  const input = byId<HTMLInputElement>('note-input');
  if (!input || !input.value.trim()) return;
  await window.api?.addNote?.(input.value.trim());
  input.value = '';
  refreshNotes();
});
refreshNotes();

(async () => {
  const sel = byId<HTMLSelectElement>('theme-select');
  if (!sel) return;
  const val = await window.api?.getSetting?.('theme', 'system');
  if (val) sel.value = val;
  sel.addEventListener('change', () => window.api?.setSetting?.('theme', sel.value));
})();

byId<HTMLButtonElement>('btn-check-update')?.addEventListener('click', () => {
  window.api?.checkForUpdates?.();
});

// FILE PICKER LOGIC
function bindPrettyFile(id: string) {
  const input = document.getElementById(id) as HTMLInputElement | null;
  if (!input) return;
  const label = input.closest('.file-picker');
  const nameSpan = label?.querySelector('.file-name') as HTMLElement | null;
  const fileBtn = label?.querySelector('.file-btn') as HTMLElement | null;
  const placeholder = nameSpan?.getAttribute('data-placeholder') || 'Файл не вибрано';
  
  const refresh = () => {
    const f = input.files?.[0];
    if (nameSpan) {
      nameSpan.textContent = f ? f.name : placeholder;
      nameSpan.classList.toggle('empty', !f);
    }
  };
  
  // Handle file button click
  fileBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    input.click();
  });
  
  input.addEventListener('change', refresh);
  refresh();
}

// Initialize file pickers
bindPrettyFile('word-file');
bindPrettyFile('excel-db');

// Updates functionality
class UpdateManager {
  private currentManifest: any = null;
  private isProcessing = false;

  constructor() {
    this.bindEvents();
    this.setupEventListeners();
    this.loadCurrentVersion();
  }

  private bindEvents() {
    byId('btn-check-updates')?.addEventListener('click', () => this.checkForUpdates());
    byId('btn-download-update')?.addEventListener('click', () => this.downloadUpdate());
    byId('btn-mandatory-update')?.addEventListener('click', () => this.downloadUpdate());
    byId('btn-exit-app')?.addEventListener('click', () => window.close());
    byId('btn-retry-update')?.addEventListener('click', () => this.retryUpdate());
    byId('btn-cancel-update')?.addEventListener('click', () => this.cancelUpdate());
    byId('btn-save-log')?.addEventListener('click', () => this.saveErrorLog());
  }

  private setupEventListeners() {
    // Слухач змін стану оновлення
    window.api?.onUpdateStateChanged?.((state: string) => {
      this.handleStateChange(state);
    });

    // Слухач прогресу завантаження
    window.api?.onUpdateProgress?.((progress: any) => {
      this.updateProgress(progress);
    });
  }

  private async loadCurrentVersion() {
    try {
      const version = await window.api?.getUpdateVersion?.();
      if (version) {
        const versionEl = byId('current-version');
        if (versionEl) versionEl.textContent = version;
      }
    } catch (error) {
      console.error('Помилка отримання версії:', error);
    }
  }

  private async checkForUpdates() {
    if (this.isProcessing) return;
    
    this.isProcessing = true;
    this.setStatus('Перевірка оновлень...');
    this.hideAllPanels();

    try {
      const result = await window.api?.checkForUpdates?.();
      
      if (!result) {
        this.setStatus('Помилка перевірки оновлень');
        return;
      }

      switch (result.state) {
        case 'uptodate':
          this.setStatus('Встановлена остання версія');
          break;
        case 'available':
          this.showUpdateAvailable(result.manifest);
          break;
        case 'mandatory':
          this.showMandatoryUpdate(result.manifest);
          break;
        case 'failed':
          this.showError(result.error || 'Невідома помилка');
          break;
      }
    } catch (error) {
      this.showError(error instanceof Error ? error.message : 'Помилка перевірки оновлень');
    } finally {
      this.isProcessing = false;
    }
  }

  private async downloadUpdate() {
    if (!this.currentManifest || this.isProcessing) return;
    
    this.isProcessing = true;
    this.hideAllPanels();
    this.showProgress('Завантаження оновлення...');

    try {
      const success = await window.api?.downloadUpdate?.(this.currentManifest);
      
      if (success) {
        // Автоматично переходимо до встановлення
        await this.installUpdate();
      } else {
        this.showError('Помилка завантаження оновлення');
      }
    } catch (error) {
      this.showError(error instanceof Error ? error.message : 'Помилка завантаження');
    } finally {
      this.isProcessing = false;
    }
  }

  private async installUpdate() {
    if (!this.currentManifest) return;
    
    this.showProgress('Встановлення оновлення...');

    try {
      const success = await window.api?.installUpdate?.(this.currentManifest);
      
      if (success) {
        this.showProgress('Перезапуск застосунку...');
        // Застосунок перезапуститься автоматично
      } else {
        this.showError('Помилка встановлення оновлення');
      }
    } catch (error) {
      this.showError(error instanceof Error ? error.message : 'Помилка встановлення');
    }
  }

  private handleStateChange(state: string) {
    console.log('Стан оновлення змінився:', state);
    
    switch (state) {
      case 'checking':
        this.setStatus('Перевірка оновлень...');
        break;
      case 'downloading':
        this.showProgress('Завантаження...');
        break;
      case 'verifying':
        this.showProgress('Перевірка цілісності...');
        break;
      case 'installing':
        this.showProgress('Встановлення...');
        break;
      case 'restarting':
        this.showProgress('Перезапуск застосунку...');
        break;
      case 'failed':
        this.showError('Помилка оновлення');
        break;
    }
  }

  private updateProgress(progress: any) {
    const progressEl = byId('progress-fill');
    const percentEl = byId('progress-percent');
    const speedEl = byId('progress-speed');
    const sizeEl = byId('progress-size');

    if (progressEl) {
      progressEl.style.width = `${progress.percent}%`;
    }
    
    if (percentEl) {
      percentEl.textContent = `${progress.percent}%`;
    }
    
    if (speedEl && progress.speedKbps) {
      speedEl.textContent = `${Math.round(progress.speedKbps)} кБ/с`;
    }
    
    if (sizeEl) {
      const receivedMB = (progress.bytesReceived / 1024 / 1024).toFixed(1);
      const totalMB = (progress.totalBytes / 1024 / 1024).toFixed(1);
      sizeEl.textContent = `${receivedMB} МБ з ${totalMB} МБ`;
    }
  }

  private showUpdateAvailable(manifest: any) {
    this.currentManifest = manifest;
    this.setStatus('Доступне оновлення');
    
    const panel = byId('update-available');
    const versionEl = byId('new-version');
    const sizeEl = byId('update-size');
    const dateEl = byId('update-date');
    const notesLink = byId('release-notes') as HTMLAnchorElement;

    if (panel) panel.hidden = false;
    if (versionEl) versionEl.textContent = manifest.version;
    if (sizeEl) sizeEl.textContent = this.formatSize(manifest.asset.size);
    if (dateEl) dateEl.textContent = new Date(manifest.published_utc).toLocaleDateString('uk-UA');
    if (notesLink) notesLink.href = manifest.notes_url;
  }

  private showMandatoryUpdate(manifest: any) {
    this.currentManifest = manifest;
    const panel = byId('mandatory-update');
    if (panel) panel.hidden = false;
  }

  private showProgress(text: string) {
    const panel = byId('update-progress');
    const textEl = byId('progress-text');
    
    if (panel) panel.hidden = false;
    if (textEl) textEl.textContent = text;
  }

  private showError(message: string) {
    const panel = byId('update-error');
    const messageEl = byId('error-message');
    
    if (panel) panel.hidden = false;
    if (messageEl) messageEl.textContent = message;
    
    this.setStatus('Помилка оновлення');
  }

  private hideAllPanels() {
    const panels = ['update-available', 'mandatory-update', 'update-progress', 'update-error'];
    panels.forEach(id => {
      const panel = byId(id);
      if (panel) panel.hidden = true;
    });
  }

  private setStatus(status: string) {
    const statusEl = byId('update-status');
    if (statusEl) statusEl.textContent = status;
  }

  private retryUpdate() {
    this.hideAllPanels();
    this.checkForUpdates();
  }

  private cancelUpdate() {
    this.isProcessing = false;
    this.hideAllPanels();
    this.setStatus('Оновлення скасовано');
  }

  private saveErrorLog() {
    // Тут можна реалізувати збереження логу помилок
    console.log('Збереження логу помилок...');
  }

  private formatSize(bytes: number): string {
    const mb = bytes / 1024 / 1024;
    return `${mb.toFixed(1)} МБ`;
  }
}

// Ініціалізація менеджера оновлень
const updateManager = new UpdateManager();

// Клас для управління пакетною обробкою
class BatchManager {
  private isProcessing = false;
  private logContainer?: HTMLElement;
  
  constructor() {
    this.setupElements();
    this.setupEventListeners();
    this.loadSavedSettings();
  }

  private setupElements() {
    this.logContainer = byId('batch-log-body') || undefined;
  }

  private setupEventListeners() {
    // Вибір папки
    byId('btn-choose-batch-folder')?.addEventListener('click', () => {
      this.selectInputFolder();
    });

    // Вибір файлу результату
    byId('btn-choose-batch-output')?.addEventListener('click', () => {
      this.selectOutputFile();
    });

    // Запуск обробки
    byId('btn-start-batch')?.addEventListener('click', () => {
      this.startProcessing();
    });

    // Скасування обробки
    byId('btn-cancel-batch')?.addEventListener('click', () => {
      this.cancelProcessing();
    });

    // Очищення логу
    byId('btn-clear-batch-log')?.addEventListener('click', () => {
      this.clearLog();
    });

    // Збереження логу
    byId('btn-save-batch-log')?.addEventListener('click', () => {
      this.saveLog();
    });

    // Слухачі подій з backend
    window.api?.onBatchProgress?.((progress: any) => {
      this.updateProgress(progress);
    });

    window.api?.onBatchLog?.((logEntry: { level: string, message: string }) => {
      this.addLogEntry(logEntry.level, logEntry.message);
    });

    window.api?.onBatchComplete?.((result: any) => {
      this.onProcessingComplete(result);
    });
  }

  private async selectInputFolder() {
    try {
      const folderPath = await window.api?.selectBatchDirectory?.();
      if (folderPath) {
        const input = byId<HTMLInputElement>('batch-input-folder');
        if (input) {
          input.value = folderPath;
          this.updateButtonStates();
        }
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка вибору папки: ${error}`);
    }
  }

  private async selectOutputFile() {
    try {
      const filePath = await window.api?.selectBatchOutputFile?.();
      if (filePath) {
        const input = byId<HTMLInputElement>('batch-output-file');
        if (input) {
          input.value = filePath;
        }
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка вибору файлу: ${error}`);
    }
  }

  private async startProcessing() {
    if (this.isProcessing) return;

    const inputFolder = byId<HTMLInputElement>('batch-input-folder')?.value;
    if (!inputFolder) {
      this.addLogEntry('error', 'Оберіть вхідну папку');
      return;
    }

    let outputFile = byId<HTMLInputElement>('batch-output-file')?.value;
    if (!outputFile) {
      // Генеруємо автоматичне ім'я файлу
      const now = new Date();
      const dateStr = now.toISOString().split('T')[0]; // YYYY-MM-DD
      const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-'); // HH-MM-SS
      outputFile = `${inputFolder}\\Індекс_бійців_${dateStr}_${timeStr}.xlsx`;
    }

    const options = {
      inputDirectory: inputFolder,
      outputFilePath: outputFile,
      includeStats: true,
      includeConflicts: true,
      includeOccurrences: false, // За замовчуванням відключено через розмір
      resolveConflicts: true
    };

    try {
      this.isProcessing = true;
      this.updateButtonStates();
      this.showProgress();
      this.clearLog();
      this.addLogEntry('info', 'Початок пакетної обробки...');

      await window.api?.startBatchProcessing?.(options);
    } catch (error) {
      this.addLogEntry('error', `Помилка запуску: ${error}`);
      this.isProcessing = false;
      this.updateButtonStates();
      this.hideProgress();
    }
  }

  private async cancelProcessing() {
    if (!this.isProcessing) return;

    try {
      const cancelled = await window.api?.cancelBatchProcessing?.();
      if (cancelled) {
        this.addLogEntry('warning', 'Обробку скасовано');
      }
    } catch (error) {
      this.addLogEntry('error', `Помилка скасування: ${error}`);
    }
  }

  private updateProgress(progress: any) {
    // Оновлення прогрес-бару
    const progressFill = byId('batch-progress-fill');
    const progressPercent = byId('batch-progress-percent');
    const progressStatus = byId('batch-progress-status');
    const progressDetail = byId('batch-progress-detail');

    if (progressFill) {
      progressFill.style.width = `${progress.percentage}%`;
    }

    if (progressPercent) {
      progressPercent.textContent = `${progress.percentage}%`;
    }

    if (progressStatus) {
      progressStatus.textContent = progress.message;
    }

    if (progressDetail) {
      const timeElapsedStr = Math.round(progress.timeElapsed / 1000);
      let detailText = `Файлів оброблено: ${progress.filesProcessed}/${progress.totalFiles} (${timeElapsedStr}с)`;
      
      if (progress.estimatedTimeRemaining) {
        const etaStr = Math.round(progress.estimatedTimeRemaining / 1000);
        detailText += `, залишилось ~${etaStr}с`;
      }
      
      progressDetail.textContent = detailText;
    }
  }

  private addLogEntry(level: string, message: string) {
    if (!this.logContainer) return;

    const timestamp = new Date().toLocaleTimeString('uk-UA');
    const levelIcon = {
      'info': 'ℹ️',
      'warning': '⚠️',
      'error': '❌'
    }[level] || 'ℹ️';

    const logLine = `[${timestamp}] ${levelIcon} ${message}\n`;
    this.logContainer.textContent += logLine;
    this.logContainer.scrollTop = this.logContainer.scrollHeight;
  }

  private onProcessingComplete(result: any) {
    this.isProcessing = false;
    this.updateButtonStates();
    this.hideProgress();

    if (result.success) {
      this.addLogEntry('info', `✅ Обробка завершена успішно!`);
      this.addLogEntry('info', `📊 Знайдено ${result.stats.fightersFound} бійців`);
      this.addLogEntry('info', `📁 Результат: ${result.outputFilePath}`);
      
      if (result.stats.conflicts > 0) {
        this.addLogEntry('warning', `⚠️ Знайдено ${result.stats.conflicts} конфліктів`);
      }
    } else {
      this.addLogEntry('error', '❌ Обробка завершилася з помилками');
      result.errors.forEach((error: string) => {
        this.addLogEntry('error', error);
      });
    }
  }

  private showProgress() {
    const progressContainer = byId('batch-progress');
    if (progressContainer) {
      progressContainer.hidden = false;
    }
  }

  private hideProgress() {
    const progressContainer = byId('batch-progress');
    if (progressContainer) {
      progressContainer.hidden = true;
    }
  }

  private updateButtonStates() {
    const inputFolder = byId<HTMLInputElement>('batch-input-folder')?.value;
    const startBtn = byId<HTMLButtonElement>('btn-start-batch');
    const cancelBtn = byId<HTMLButtonElement>('btn-cancel-batch');

    if (startBtn) {
      startBtn.disabled = this.isProcessing || !inputFolder;
      startBtn.textContent = this.isProcessing ? 'Обробляється...' : 'Обробити';
    }

    if (cancelBtn) {
      cancelBtn.disabled = !this.isProcessing;
    }
  }

  private clearLog() {
    if (this.logContainer) {
      this.logContainer.textContent = '';
    }
  }

  private saveLog() {
    if (this.logContainer) {
      const logText = this.logContainer.textContent;
      const blob = new Blob([logText], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = url;
      a.download = `batch_log_${new Date().toISOString().split('T')[0]}.txt`;
      a.click();
      
      URL.revokeObjectURL(url);
    }
  }

  private loadSavedSettings() {
    // Завантажуємо збережені налаштування
    this.updateButtonStates();
  }
}

// Ініціалізація менеджера пакетної обробки
const batchManager = new BatchManager();
