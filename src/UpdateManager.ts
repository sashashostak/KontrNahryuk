/**
 * UpdateManager - Управління автооновленням програми
 * FIXED: Винесено з main.ts (рядки 486-914)
 * v3.0: Інтеграція з Advanced Auto-Update System (EventEmitter API)
 * 
 * Відповідальність:
 * - Перевірка наявності оновлень
 * - Завантаження та встановлення оновлень
 * - Відображення прогресу оновлення (реальний час: %, швидкість, розмір)
 * - Обробка статусних повідомлень (розпакування, backup, заміна файлів)
 * - Обробка помилок оновлення
 * - Перезапуск після оновлення
 * 
 * Інтеграція з updateService v3.0:
 * - onDownloadProgress: прогрес завантаження з fetch streams
 * - onUpdateStatus: статуси автоматичної установки
 * - onUpdateError: обробка помилок в реальному часі
 * 
 * @class UpdateManager
 */

import type { UpdateInfo } from './types';
import { byId, log } from './helpers';

export class UpdateManager {
  private isProcessing: boolean = false;
  private currentUpdateInfo: UpdateInfo | null = null;

  /**
   * Конструктор - ініціалізує UpdateManager
   * FIXED: Автоматично налаштовує слухачів подій
   */
  constructor() {
    this.setupUpdateListeners();
    this.setupUpdateEventListeners();
    this.bindUpdateButtons();
  }

  /**
   * Прив'язка обробників до кнопок оновлення
   * FIXED: Використовує правильні ID з HTML
   * @private
   */
  private bindUpdateButtons(): void {
    // Кнопка автоматичного оновлення
    byId('btn-auto-update')?.addEventListener('click', () => {
      this.downloadAndInstallUpdate();
    });

    // Кнопка обов'язкового оновлення
    byId('btn-mandatory-update')?.addEventListener('click', () => {
      this.downloadAndInstallUpdate();
    });

    // Кнопка ручного завантаження
    byId('btn-manual-download')?.addEventListener('click', () => {
      this.openDownloadPage();
    });

    // Кнопка скасування оновлення
    byId('btn-cancel-update')?.addEventListener('click', () => {
      this.cancelUpdate();
    });

    // Кнопка перезапуску після оновлення
    byId('btn-restart-after-update')?.addEventListener('click', () => {
      this.restartApp();
    });

    // Кнопка повтору оновлення при помилці
    byId('btn-retry-update')?.addEventListener('click', () => {
      this.retryUpdate();
    });

    // Кнопка збереження логу помилки
    byId('btn-save-log')?.addEventListener('click', () => {
      this.saveUpdateLog();
    });

    // Кнопка виходу з додатку
    byId('btn-exit-app')?.addEventListener('click', () => {
      window.close();
    });
  }

  /**
   * Перевірка наявності оновлень
   * FIXED: Асинхронна перевірка з обробкою помилок
   * @public
   */
  public async checkForUpdates(): Promise<void> {
    if (this.isProcessing) {
      log('⚠️ Перевірка оновлень вже виконується');
      return;
    }

    this.isProcessing = true;
    
    const statusDiv = byId('update-status');
    const updateAvailableDiv = byId('update-available');
    
    if (statusDiv) statusDiv.textContent = 'Перевіряємо наявність оновлень...';
    if (updateAvailableDiv) updateAvailableDiv.hidden = true;

    try {
      // FIXED: Використовуємо API через electron main процес для обходу CORS
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

  /**
   * Налаштування слухачів подій оновлення від main process
   * FIXED: Оновлено для Advanced API з EventEmitter (v3.0)
   * @private
   */
  private setupUpdateListeners(): void {
    // Прогрес завантаження з updateService v3.0 (EventEmitter)
    (window as any).api?.onDownloadProgress?.((progress: any) => {
      this.handleUpdateProgress({
        percent: progress.percent || 0,
        speedKbps: progress.bytesPerSecond ? (progress.bytesPerSecond / 1024) : 0,
        bytesReceived: progress.downloadedBytes || 0,
        totalBytes: progress.totalBytes || 0
      });
    });

    // Статусні повідомлення (розпакування, створення backup, заміна файлів)
    (window as any).api?.onUpdateStatus?.((status: any) => {
      const progressText = document.getElementById('progress-text');
      if (progressText && status.message) {
        progressText.textContent = `🐷 ${status.message}`;
      }
    });

    // Підписка на зміну статусу оновлення (залишаємо для сумісності)
    (window as any).api?.onUpdateStateChanged?.((state: string) => {
      this.handleUpdateStateChange(state);
    });

    // Обробка помилок оновлення
    (window as any).api?.onUpdateError?.((error: any) => {
      this.handleUpdateError(error.message || error);
    });

    // Підписка на завершення оновлення
    (window as any).api?.onUpdateComplete?.(() => {
      this.handleUpdateComplete();
    });
  }

  /**
   * Додаткове налаштування слухачів подій (Advanced API з прогрес-баром)
   * FIXED: Оновлено для підтримки EventEmitter API з updateService v3.0
   * @private
   */
  private setupUpdateEventListeners(): void {
    // Прогрес завантаження (нова EventEmitter API з v3.0)
    (window as any).api?.onDownloadProgress?.((progress: any) => {
      this.handleUpdateProgress({
        percent: progress.percent || 0,
        speedKbps: progress.bytesPerSecond ? (progress.bytesPerSecond / 1024) : 0,
        bytesReceived: progress.downloadedBytes || 0,
        totalBytes: progress.totalBytes || 0
      });
    });

    // Статусні повідомлення (розпакування, backup, заміна файлів)
    (window as any).api?.onUpdateStatus?.((status: any) => {
      const progressText = document.getElementById('progress-text');
      if (progressText && status.message) {
        progressText.textContent = `🐷 ${status.message}`;
      }
    });

    // Обробка помилок
    (window as any).api?.onUpdateError?.((error: any) => {
      this.showUpdateError(error.message || error);
    });
  }

  /**
   * Завантаження та встановлення оновлення
   * FIXED: Асинхронне завантаження з відображенням прогресу
   * @private
   */
  private async downloadAndInstallUpdate(): Promise<void> {
    if (!this.currentUpdateInfo) {
      this.showUpdateError('Немає інформації про оновлення');
      return;
    }

    try {
      this.showUpdateProgress('Завантаження оновлення...');
      
      const result = await (window as any).api?.downloadUpdate?.(this.currentUpdateInfo);
      
      if (!result?.success) {
        this.showUpdateError('Не вдалося завантажити оновлення');
      }
    } catch (error) {
      console.error('Помилка завантаження:', error);
      this.showUpdateError('Помилка під час завантаження: ' + error);
    }
  }

  /**
   * Відкриття сторінки завантаження у браузері
   * FIXED: Використовує external browser
   * @private
   */
  private openDownloadPage(): void {
    if (this.currentUpdateInfo?.releaseInfo?.html_url) {
      (window as any).api?.openExternal?.(this.currentUpdateInfo.releaseInfo.html_url);
    }
  }

  /**
   * Скасування поточного оновлення
   * NOTE: У спрощеній версії просто приховуємо прогрес та повертаємось до початкового стану
   * @private
   */
  private async cancelUpdate(): Promise<void> {
    try {
      this.hideUpdateProgress();
      this.showUpdateAvailable(); // Повертаємо до стану "доступне оновлення"
    } catch (error) {
      console.error('Помилка скасування оновлення:', error);
    }
  }

  /**
   * Перезапуск застосунку після встановлення оновлення
   * FIXED: Викликає API перезапуску
   * @private
   */
  private async restartApp(): Promise<void> {
    try {
      await (window as any).api?.restartApp?.();
    } catch (error) {
      console.error('Помилка перезапуску:', error);
    }
  }

  /**
   * Показати прогрес оновлення
   * FIXED: Керує видимістю елементів UI
   * @private
   */
  private showUpdateProgress(text: string): void {
    const progressDiv = byId('update-progress');
    const progressText = byId('progress-text');
    const availableDiv = byId('update-available');
    
    if (progressDiv) progressDiv.hidden = false;
    if (progressText) progressText.textContent = text;
    if (availableDiv) availableDiv.hidden = true;
  }

  /**
   * Сховати прогрес оновлення
   * FIXED: Приховує UI прогресу
   * @private
   */
  private hideUpdateProgress(): void {
    const progressDiv = byId('update-progress');
    if (progressDiv) progressDiv.hidden = true;
  }

  /**
   * Показати блок "Доступне оновлення"
   * FIXED: Перемикання між UI станами
   * @private
   */
  private showUpdateAvailable(): void {
    const availableDiv = byId('update-available');
    const progressDiv = byId('update-progress');
    
    if (availableDiv) availableDiv.hidden = false;
    if (progressDiv) progressDiv.hidden = true;
  }

  /**
   * Обробка прогресу оновлення від main process
   * FIXED: Оновлює UI елементи прогресу
   * @private
   */
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

  /**
   * Обробка зміни стану оновлення
   * FIXED: Показує емодзі статуси
   * @private
   */
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

  /**
   * Обробка помилки оновлення
   * FIXED: Показує повідомлення про помилку
   * @private
   */
  private handleUpdateError(error: string): void {
    this.showUpdateError(error);
  }

  /**
   * Обробка завершення оновлення
   * FIXED: Показує кнопку перезапуску
   * @private
   */
  private handleUpdateComplete(): void {
    const progressText = byId('progress-text');
    const restartBtn = byId('btn-restart-after-update');
    
    if (progressText) progressText.textContent = '✅ Оновлення встановлено! Натисніть "Перезапустити"';
    if (restartBtn) restartBtn.style.display = 'block';
  }

  /**
   * Показати помилку оновлення
   * FIXED: Відображає діалог помилки
   * @private
   */
  private showUpdateError(message: string): void {
    const errorDiv = byId('update-error');
    const errorMessage = byId('error-message');
    const progressDiv = byId('update-progress');
    
    if (errorDiv) errorDiv.hidden = false;
    if (errorMessage) errorMessage.textContent = message;
    if (progressDiv) progressDiv.hidden = true;
  }

  /**
   * Повторна спроба оновлення
   * FIXED: Ховає діалог помилки та перезапускає оновлення
   * @private
   */
  private retryUpdate(): void {
    // Ховаємо діалог помилки і повторюємо спробу оновлення
    const errorDiv = byId('update-error');
    if (errorDiv) errorDiv.hidden = true;
    
    // Повторюємо завантаження оновлення
    this.downloadAndInstallUpdate();
  }

  /**
   * Збереження логу помилки оновлення
   * FIXED: Створює файл логу з деталями помилки
   * @private
   */
  private async saveUpdateLog(): Promise<void> {
    try {
      // Отримуємо текст помилки
      const errorMessage = byId('error-message')?.textContent || 'Невідома помилка оновлення';
      
      // FIXED: Створюємо лог з деталями
      const currentVersion = await (window as any).api?.invoke?.('updates:get-version') || 'невідомо';
      const logContent = [
        `=== Лог помилки оновлення KontrNahryuk ===`,
        `Час: ${new Date().toLocaleString()}`,
        `Поточна версія: ${currentVersion}`,
        `Спроба оновлення до: ${this.currentUpdateInfo?.latestVersion || 'невідомо'}`,
        `Помилка: ${errorMessage}`,
        ``,
        `Деталі оновлення:`,
        JSON.stringify(this.currentUpdateInfo, null, 2),
        ``,
        `=== Кінець логу ===`
      ].join('\n');

      // Зберігаємо лог локально через download
      const blob = new Blob([logContent], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `update-log-${new Date().getTime()}.txt`;
      a.click();
      URL.revokeObjectURL(url);
      
      // Показуємо повідомлення про успішне збереження
      const errorMessageEl = byId('error-message');
      if (errorMessageEl) {
        const originalText = errorMessageEl.textContent;
        errorMessageEl.textContent = 'Лог збережено! Перевірте папку Downloads.';
        
        // Повертаємо оригінальний текст через 3 секунди
        setTimeout(() => {
          if (errorMessageEl) errorMessageEl.textContent = originalText;
        }, 3000);
      }
    } catch (error) {
      console.error('Помилка збереження логу:', error);
    }
  }
}
