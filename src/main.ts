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

// Check existing license key on startup
async function checkExistingLicense() {
  try {
    if (window.api?.checkExistingLicense) {
      const licenseResult = await window.api.checkExistingLicense();
      if (licenseResult?.hasAccess) {
        return true;
      } else {
        return false;
      }
    }
  } catch (error) {
    console.error('Помилка перевірки ліцензійного ключа:', error);
    return false;
  }
  return false;
}

// Initialize settings on page load
(async () => {
  await loadSettings();
  setupSettingsAutoSave();
  
  // Перевіряємо ліцензійний ключ при запуску
  setTimeout(async () => {
    await checkExistingLicense();
  }, 1000);
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
bindPrettyFile('word-files');
bindPrettyFile('excel-db');

// Source Selection Manager
class SourceSelectionManager {
  private sourceRadios: NodeListOf<HTMLInputElement>;
  private singleFileInput: HTMLElement | null;
  private multipleFilesInput: HTMLElement | null;
  private folderInput: HTMLElement | null;

  constructor() {
    this.sourceRadios = document.querySelectorAll('input[name="source-type"]');
    this.singleFileInput = byId('single-file-input');
    this.multipleFilesInput = byId('multiple-files-input');
    this.folderInput = byId('folder-input');
    
    this.bindEvents();
  }

  private bindEvents() {
    this.sourceRadios.forEach(radio => {
      radio.addEventListener('change', () => {
        this.handleSourceChange(radio.value);
      });
    });

    // Обробник для кнопки вибору папки
    byId('choose-folder')?.addEventListener('click', () => {
      this.selectFolder();
    });
  }

  private handleSourceChange(sourceType: string) {
    // Ховаємо всі inputs
    if (this.singleFileInput) this.singleFileInput.style.display = 'none';
    if (this.multipleFilesInput) this.multipleFilesInput.style.display = 'none';
    if (this.folderInput) this.folderInput.style.display = 'none';

    // Показуємо потрібний input
    switch (sourceType) {
      case 'single-file':
        if (this.singleFileInput) this.singleFileInput.style.display = 'block';
        break;
      case 'multiple-files':
        if (this.multipleFilesInput) this.multipleFilesInput.style.display = 'block';
        break;
      case 'folder':
        if (this.folderInput) this.folderInput.style.display = 'block';
        break;
    }
  }

  private async selectFolder() {
    try {
      const folderPath = await window.api?.selectBatchDirectory?.();
      if (folderPath) {
        const folderInput = byId('folder-path') as HTMLInputElement;
        if (folderInput) {
          folderInput.value = folderPath;
        }
        log(`📁 Обрано папку: ${folderPath}`);
      }
    } catch (error) {
      console.error('Помилка вибору папки:', error);
      log(`❌ Помилка вибору папки: ${error}`);
    }
  }

  public getSelectedSource() {
    const checkedRadio = document.querySelector('input[name="source-type"]:checked') as HTMLInputElement;
    return checkedRadio ? checkedRadio.value : 'single-file';
  }

  public getSelectedFiles() {
    const sourceType = this.getSelectedSource();
    
    switch (sourceType) {
      case 'single-file':
        const singleFile = byId('word-file') as HTMLInputElement;
        return singleFile?.files ? Array.from(singleFile.files) : [];
        
      case 'multiple-files':
        const multipleFiles = byId('word-files') as HTMLInputElement;
        return multipleFiles?.files ? Array.from(multipleFiles.files) : [];
        
      case 'folder':
        const folderPath = (byId('folder-path') as HTMLInputElement)?.value;
        return folderPath ? [folderPath] : [];
        
      default:
        return [];
    }
  }
}

// Initialize source selection
const sourceManager = new SourceSelectionManager();

// Section Manager for conditional display
class SectionManager {
  private orderCheckbox: HTMLInputElement | null;
  private excelSection: HTMLElement | null;

  constructor() {
    this.orderCheckbox = byId('t-order') as HTMLInputElement;
    this.excelSection = byId('excel-section');
    this.bindEvents();
  }

  private bindEvents() {
    this.orderCheckbox?.addEventListener('change', () => {
      this.toggleExcelSection();
    });

    // Обробник для кнопки вибору Excel файлу
    byId('choose-excel')?.addEventListener('click', () => {
      this.selectExcelFile();
    });
  }

  private toggleExcelSection() {
    if (this.excelSection) {
      this.excelSection.style.display = this.orderCheckbox?.checked ? 'block' : 'none';
    }
  }

  private async selectExcelFile() {
    try {
      const filePath = await window.api?.selectExcelFile?.();
      if (filePath) {
        const excelInput = byId('excel-path') as HTMLInputElement;
        if (excelInput) {
          excelInput.value = filePath;
        }
        log(`📊 Обрано Excel файл: ${filePath}`);
      }
    } catch (error) {
      console.error('Помилка вибору Excel файлу:', error);
      log(`❌ Помилка вибору Excel файлу: ${error}`);
    }
  }
}

// Initialize section manager
const sectionManager = new SectionManager();

// Updates functionality
class UpdateManager {
  private isProcessing = false;

  constructor() {
    this.bindEvents();
    this.loadCurrentVersion();
    this.checkLicenseOnStartup();
    this.setupUpdateEventListeners();
  }

  private bindEvents() {
    byId('btn-check-updates')?.addEventListener('click', () => this.checkForUpdates());

    // Автооновлення
    byId('btn-auto-update')?.addEventListener('click', () => this.downloadAndInstallUpdate());
    byId('btn-manual-download')?.addEventListener('click', () => this.openDownloadPage());
    byId('btn-cancel-update')?.addEventListener('click', () => this.cancelUpdate());
    byId('btn-restart-after-update')?.addEventListener('click', () => this.restartApp());
    
    // Кнопки діалогу помилок оновлення
    byId('btn-retry-update')?.addEventListener('click', () => this.retryUpdate());
    byId('btn-save-log')?.addEventListener('click', () => this.saveUpdateLog());

    // Обробка ліцензійного ключа
    byId('btn-set-license')?.addEventListener('click', () => this.setLicenseKey());
    
    // Обробка ліцензійного gate
    byId('gate-license-btn')?.addEventListener('click', () => this.activateLicense());

    // Підписка на події оновлення
    this.setupUpdateListeners();
  }

  private async loadCurrentVersion() {
    const versionEl = byId('current-version');
    if (versionEl) versionEl.textContent = '1.2.4';
  }

  private async checkForUpdates() {
    if (this.isProcessing) return;
    
    this.isProcessing = true;
    const statusDiv = byId('update-status');
    const updateAvailableDiv = byId('update-available');
    
    if (statusDiv) statusDiv.textContent = 'Перевіряємо наявність оновлень...';
    if (updateAvailableDiv) updateAvailableDiv.hidden = true;

    try {
      // Використовуємо API через electron main процес для обходу CORS
      const updateInfo = await (window as any).api?.checkForUpdates?.();
      
      if (!updateInfo) {
        throw new Error('Не вдалося отримати інформацію про оновлення');
      }

      if (updateInfo.error) {
        throw new Error(updateInfo.error);
      }

      const { hasUpdate, latestVersion, releaseInfo } = updateInfo;
      
      if (hasUpdate && releaseInfo) {
        // Є нове оновлення
        if (statusDiv) statusDiv.textContent = `Доступна нова версія: ${latestVersion}`;
        
        // Показуємо блок з кнопкою оновлення
        if (updateAvailableDiv) {
          // Заповнюємо дані про оновлення
          const newVersionSpan = byId('new-version');
          const updateDateSpan = byId('update-date');
          const releasNotesLink = byId<HTMLAnchorElement>('release-notes');
          
          if (newVersionSpan) newVersionSpan.textContent = latestVersion;
          if (updateDateSpan) updateDateSpan.textContent = new Date(releaseInfo.published_at).toLocaleDateString();
          if (releasNotesLink) releasNotesLink.href = releaseInfo.html_url;
          
          // Показуємо блок оновлення
          updateAvailableDiv.hidden = false;
          
          // Зберігаємо інформацію про оновлення для використання пізніше
          this.currentUpdateInfo = { hasUpdate, latestVersion, releaseInfo };
        }
      } else {
        // Актуальна версія - ховаємо блок оновлення і показуємо статус
        if (statusDiv) statusDiv.textContent = 'Актуальна версія';
        if (updateAvailableDiv) updateAvailableDiv.hidden = true;
      }
    } catch (error) {
      console.error('Помилка перевірки оновлень:', error);
      if (statusDiv) statusDiv.textContent = 'Помилка перевірки оновлень. Перевірте інтернет-з\'єднання.';
    } finally {
      this.isProcessing = false;
    }
  }

  private async setLicenseKey(): Promise<void> {
    const input = byId<HTMLInputElement>('license-key-input');
    const statusDiv = byId('license-status');
    
    if (!input || !statusDiv) return;
    
    const key = input.value.trim();
    if (!key) {
      this.updateLicenseStatus('Введіть ліцензійний ключ', 'invalid');
      return;
    }

    this.updateLicenseStatus('Перевірка ключа...', 'pending');
    
    try {
      const result = await (window as any).api.setLicenseKey(key);
      if (result.hasAccess) {
        this.updateLicenseStatus(`Ліцензія активна (${result.licenseInfo?.plan || 'Basic'})`, 'valid');
        input.value = '';
      } else {
        this.updateLicenseStatus(result.reason || 'Невірний ліцензійний ключ', 'invalid');
      }
    } catch (error) {
      console.error('Помилка при встановленні ліцензії:', error);
      this.updateLicenseStatus('Помилка з\'єднання', 'invalid');
    }
  }

  private updateLicenseStatus(message: string, state: 'valid' | 'invalid' | 'pending'): void {
    const statusDiv = byId('license-status');
    if (!statusDiv) return;
    
    statusDiv.textContent = message;
    statusDiv.className = `license-status ${state}`;
  }

  private async loadLicenseInfo(): Promise<void> {
    try {
      const info = await (window as any).api.getLicenseInfo();
      if (info?.hasAccess) {
        this.updateLicenseStatus(`Ліцензія активна (${info.licenseInfo?.plan || 'Universal'})`, 'valid');
        // Приховуємо поле введення ліцензійного ключа якщо ліцензія активна
        this.hideLicenseInput();
      } else {
        this.updateLicenseStatus('Ліцензія не активована', 'invalid');
        this.showLicenseInput();
      }
    } catch (error) {
      console.error('Помилка завантаження інформації про ліцензію:', error);
      this.updateLicenseStatus('Ліцензія не активована', 'invalid');
      this.showLicenseInput();
    }
  }

  private hideLicenseInput(): void {
    const licenseInputSection = byId('license-input-section');
    if (licenseInputSection) {
      licenseInputSection.style.display = 'none';
    }
  }

  private showLicenseInput(): void {
    const licenseInputSection = byId('license-input-section');
    if (licenseInputSection) {
      licenseInputSection.style.display = 'block';
    }
  }

  // Нові методи для обов'язкового ліцензування
  private async checkLicenseOnStartup(): Promise<void> {
    try {
      const info = await (window as any).api.getLicenseInfo();
      if (info?.hasAccess) {
        this.showMainApp();
      } else {
        this.showLicenseGate();
      }
    } catch (error) {
      console.error('Помилка перевірки ліцензії:', error);
      this.showLicenseGate();
    }
  }

  private showLicenseGate(): void {
    const gate = byId('license-gate');
    if (gate) {
      gate.style.display = 'flex';
    }
    // Приховуємо основний контент
    document.querySelectorAll<HTMLElement>('.route').forEach(route => {
      route.style.display = 'none';
    });
  }

  private showMainApp(): void {
    const gate = byId('license-gate');
    if (gate) {
      gate.style.display = 'none';
    }
    // Показуємо основний контент
    document.querySelectorAll<HTMLElement>('.route').forEach(route => {
      route.style.display = '';
    });
    // Завантажуємо інформацію про ліцензію для updates секції
    this.loadLicenseInfo();
  }

  private async activateLicense(): Promise<void> {
    const input = byId<HTMLInputElement>('gate-license-input');
    const statusDiv = byId('gate-license-status');
    
    if (!input || !statusDiv) return;

    const key = input.value.trim();
    if (!key) {
      this.updateGateStatus('Введіть ліцензійний ключ', 'invalid');
      return;
    }

    this.updateGateStatus('Перевірка ключа...', 'pending');

    try {
      const result = await (window as any).api.setLicenseKey(key);
      if (result.hasAccess) {
        this.updateGateStatus(`Ліцензія активована успішно!`, 'valid');
        // Затримка для показу успішного повідомлення
        setTimeout(() => {
          this.showMainApp();
        }, 1000);
      } else {
        this.updateGateStatus(result.reason || 'Невірний ліцензійний ключ', 'invalid');
      }
    } catch (error) {
      console.error('Помилка активації ліцензії:', error);
      this.updateGateStatus('Помилка з\'єднання', 'invalid');
    }
  }

  private updateGateStatus(message: string, state: 'valid' | 'invalid' | 'pending'): void {
    const statusDiv = byId('gate-license-status');
    if (!statusDiv) return;
    
    statusDiv.textContent = message;
    statusDiv.className = `license-status ${state}`;
  }

  // Нові методи для автооновлення
  private currentUpdateInfo: any = null;

  private setupUpdateListeners(): void {
    // Підписка на прогрес оновлення
    (window as any).api?.onUpdateProgress?.((progress: any) => {
      this.handleUpdateProgress(progress);
    });

    // Підписка на зміну статусу оновлення
    (window as any).api?.onUpdateStateChanged?.((state: string) => {
      this.handleUpdateStateChange(state);
    });

    // Підписка на помилки оновлення
    (window as any).api?.onUpdateError?.((error: string) => {
      this.handleUpdateError(error);
    });

    // Підписка на завершення оновлення
    (window as any).api?.onUpdateComplete?.(() => {
      this.handleUpdateComplete();
    });
  }

  private async downloadAndInstallUpdate(): Promise<void> {
    if (!this.currentUpdateInfo) {
      this.showUpdateError('Немає інформації про оновлення');
      return;
    }

    try {
      this.showUpdateProgress('Підготовка до оновлення...');
      
      const success = await (window as any).api?.downloadAndInstallUpdate?.(this.currentUpdateInfo);
      
      if (!success) {
        this.showUpdateError('Не вдалося розпочати оновлення');
      }
    } catch (error) {
      console.error('Помилка оновлення:', error);
      this.showUpdateError('Помилка під час оновлення: ' + error);
    }
  }

  private openDownloadPage(): void {
    if (this.currentUpdateInfo?.releaseInfo?.html_url) {
      (window as any).api?.openExternal?.(this.currentUpdateInfo.releaseInfo.html_url);
    }
  }

  private async cancelUpdate(): Promise<void> {
    try {
      await (window as any).api?.cancelUpdate?.();
      this.hideUpdateProgress();
      this.showUpdateAvailable(); // Повертаємо до стану "доступне оновлення"
    } catch (error) {
      console.error('Помилка скасування оновлення:', error);
    }
  }

  private async restartApp(): Promise<void> {
    try {
      await (window as any).api?.restartApp?.();
    } catch (error) {
      console.error('Помилка перезапуску:', error);
    }
  }

  private showUpdateProgress(text: string): void {
    const progressDiv = byId('update-progress');
    const progressText = byId('progress-text');
    const availableDiv = byId('update-available');
    
    if (progressDiv) progressDiv.hidden = false;
    if (progressText) progressText.textContent = text;
    if (availableDiv) availableDiv.hidden = true;
  }

  private hideUpdateProgress(): void {
    const progressDiv = byId('update-progress');
    if (progressDiv) progressDiv.hidden = true;
  }

  private showUpdateAvailable(): void {
    const availableDiv = byId('update-available');
    const progressDiv = byId('update-progress');
    
    if (availableDiv) availableDiv.hidden = false;
    if (progressDiv) progressDiv.hidden = true;
  }

  private handleUpdateProgress(progress: any): void {
    const progressFill = byId('progress-fill');
    const progressPercent = byId('progress-percent');
    const progressSpeed = byId('progress-speed');
    const progressSize = byId('progress-size');

    if (progressFill) {
      progressFill.style.width = `${progress.percent || 0}%`;
    }
    
    if (progressPercent) {
      progressPercent.textContent = `${Math.round(progress.percent || 0)}%`;
    }
    
    if (progressSpeed && progress.speedKbps) {
      progressSpeed.textContent = `${Math.round(progress.speedKbps)} KB/s`;
    }
    
    if (progressSize && progress.bytesReceived && progress.totalBytes) {
      const receivedMB = (progress.bytesReceived / 1024 / 1024).toFixed(1);
      const totalMB = (progress.totalBytes / 1024 / 1024).toFixed(1);
      progressSize.textContent = `${receivedMB} / ${totalMB} MB`;
    }
  }

  private handleUpdateStateChange(state: string): void {
    const progressText = byId('progress-text');
    const restartBtn = byId('btn-restart-after-update');
    
    if (!progressText) return;

    switch (state) {
      case 'downloading':
        progressText.textContent = '🐷 Завантажуємо оновлення...';
        break;
      case 'verifying':
        progressText.textContent = '🔍 Перевіряємо цілісність файлів...';
        break;
      case 'installing':
        progressText.textContent = '⚙️ Встановлюємо оновлення...';
        break;
      case 'restarting':
        progressText.textContent = '🔄 Готуємо перезапуск...';
        if (restartBtn) restartBtn.style.display = 'block';
        break;
      case 'uptodate':
        this.hideUpdateProgress();
        break;
      case 'failed':
        this.showUpdateError('Помилка під час оновлення');
        break;
    }
  }

  private handleUpdateError(error: string): void {
    this.showUpdateError(error);
  }

  private handleUpdateComplete(): void {
    const progressText = byId('progress-text');
    const restartBtn = byId('btn-restart-after-update');
    
    if (progressText) progressText.textContent = '✅ Оновлення встановлено! Натисніть "Перезапустити"';
    if (restartBtn) restartBtn.style.display = 'block';
  }

  private showUpdateError(message: string): void {
    const errorDiv = byId('update-error');
    const errorMessage = byId('error-message');
    const progressDiv = byId('update-progress');
    
    if (errorDiv) errorDiv.hidden = false;
    if (errorMessage) errorMessage.textContent = message;
    if (progressDiv) progressDiv.hidden = true;
  }

  private retryUpdate(): void {
    // Ховаємо діалог помилки і повторюємо спробу оновлення
    const errorDiv = byId('update-error');
    if (errorDiv) errorDiv.hidden = true;
    
    // Повторюємо завантаження оновлення
    this.downloadAndInstallUpdate();
  }

  private async saveUpdateLog(): Promise<void> {
    try {
      // Отримуємо текст помилки
      const errorMessage = byId('error-message')?.textContent || 'Невідома помилка оновлення';
      
      // Створюємо лог з деталями
      const logContent = [
        `=== Лог помилки оновлення KontrNahryuk ===`,
        `Час: ${new Date().toLocaleString()}`,
        `Поточна версія: 1.2.2`,
        `Спроба оновлення до: ${this.currentUpdateInfo?.latestVersion || 'невідомо'}`,
        `Помилка: ${errorMessage}`,
        ``,
        `Деталі оновлення:`,
        JSON.stringify(this.currentUpdateInfo, null, 2),
        ``,
        `=== Кінець логу ===`
      ].join('\n');

      // Використовуємо API для збереження файлу
      const success = await (window as any).api?.saveUpdateLog?.(logContent);
      
      if (success) {
        // Показуємо повідомлення про успішне збереження
        const errorMessage = byId('error-message');
        if (errorMessage) {
          const originalText = errorMessage.textContent;
          errorMessage.textContent = 'Лог збережено! Перевірте папку Downloads.';
          
          // Повертаємо оригінальний текст через 3 секунди
          setTimeout(() => {
            if (errorMessage) errorMessage.textContent = originalText;
          }, 3000);
        }
      } else {
        console.error('Не вдалося зберегти лог оновлення');
      }
    } catch (error) {
      console.error('Помилка збереження логу:', error);
    }
  }

  private setupUpdateEventListeners(): void {
    // Обробники повідомлень від electron process
    (window as any).api?.onUpdateProgress?.((progress: any) => {
      this.updateProgressDisplay(progress);
    });

    (window as any).api?.onUpdateError?.((error: string) => {
      this.showUpdateError(error);
    });

    (window as any).api?.onUpdateDownloadStarted?.((info: any) => {
      this.showUpdateProgress(`Завантаження ${info.fileName}...`);
    });

    (window as any).api?.onUpdateDownloadCompleted?.(() => {
      this.showUpdateProgress('✅ Завантаження завершено! Файл збережено в папці Downloads.');
      setTimeout(() => {
        this.hideUpdateProgress();
      }, 3000);
    });
  }

  private updateProgressDisplay(progress: any): void {
    const progressDiv = byId('update-progress');
    const progressText = byId('progress-text');
    const progressBar = byId('progress-bar');

    if (progressDiv) progressDiv.hidden = false;
    
    if (progressText) {
      if (progress.percentage !== undefined) {
        progressText.textContent = `Завантаження: ${Math.round(progress.percentage)}%`;
      } else if (progress.message) {
        progressText.textContent = progress.message;
      }
    }

    if (progressBar && progress.percentage !== undefined) {
      progressBar.style.width = `${progress.percentage}%`;
    }
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
