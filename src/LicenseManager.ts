/**
 * LicenseManager - Управління ліцензуванням програми
 * FIXED: Винесено з main.ts UpdateManager (рядки 548-655)
 * 
 * Відповідальність:
 * - Встановлення ліцензійного ключа
 * - Перевірка статусу ліцензії
 * - Показ License Gate при запуску
 * - Активація ліцензії
 * - Управління доступом до програми
 * 
 * @class LicenseManager
 */

import type { LicenseResult } from './types';
import { byId, log } from './helpers';

type LicenseStatus = 'valid' | 'invalid' | 'pending';

export class LicenseManager {
  /**
   * Конструктор - ініціалізує LicenseManager
   * FIXED: Автоматично налаштовує слухачів подій
   */
  constructor() {
    this.bindLicenseButtons();
  }

  /**
   * Прив'язка обробників до кнопок ліцензування
   * FIXED: Налаштовує кнопки з правильними ID з HTML
   * @private
   */
  private bindLicenseButtons(): void {
    // Кнопка встановлення ліцензії в налаштуваннях
    byId('btn-set-license')?.addEventListener('click', () => {
      this.setLicenseKey();
    });

    // Enter key в полі ліцензії (налаштування)
    byId<HTMLInputElement>('license-key-input')?.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        this.setLicenseKey();
      }
    });

    // Кнопка активації в license gate (правильний ID)
    byId('gate-license-btn')?.addEventListener('click', () => {
      this.activateLicense();
    });

    // Enter key в полі ліцензії (gate)
    byId<HTMLInputElement>('gate-license-input')?.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        this.activateLicense();
      }
    });
  }

  /**
   * Встановлення ліцензійного ключа (сторінка налаштувань)
   * FIXED: Асинхронна валідація та встановлення ліцензії
   * @private
   */
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
      const result: LicenseResult = await (window as any).api.setLicenseKey(key);
      if (result.hasAccess) {
        this.updateLicenseStatus(`Ліцензія активна (${result.licenseInfo?.plan || 'Basic'})`, 'valid');
        input.value = '';
        log('✅ Ліцензію активовано успішно');
      } else {
        this.updateLicenseStatus(result.reason || 'Невірний ліцензійний ключ', 'invalid');
        log('❌ Помилка активації ліцензії');
      }
    } catch (error) {
      console.error('Помилка при встановленні ліцензії:', error);
      this.updateLicenseStatus('Помилка з\'єднання', 'invalid');
    }
  }

  /**
   * Оновлення статусу ліцензії в UI
   * FIXED: Застосовує CSS класи для стану
   * @private
   */
  private updateLicenseStatus(message: string, state: LicenseStatus): void {
    const statusDiv = byId('license-status');
    if (!statusDiv) return;
    
    statusDiv.textContent = message;
    statusDiv.className = `license-status ${state}`;
  }

  /**
   * Завантаження інформації про ліцензію
   * FIXED: Викликається при відкритті сторінки налаштувань
   * @public
   */
  public async loadLicenseInfo(): Promise<void> {
    try {
      const info: LicenseResult = await (window as any).api.getLicenseInfo();
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

  /**
   * Приховати поле введення ліцензійного ключа
   * FIXED: Використовується коли ліцензія активна
   * @private
   */
  private hideLicenseInput(): void {
    const licenseInputSection = byId('license-input-section');
    if (licenseInputSection) {
      licenseInputSection.style.display = 'none';
    }
  }

  /**
   * Показати поле введення ліцензійного ключа
   * FIXED: Використовується коли ліцензія не активна
   * @private
   */
  private showLicenseInput(): void {
    const licenseInputSection = byId('license-input-section');
    if (licenseInputSection) {
      licenseInputSection.style.display = 'block';
    }
  }

  /**
   * Перевірка ліцензії при запуску програми
   * FIXED: Показує License Gate або основний інтерфейс
   * @public
   */
  public async checkLicenseOnStartup(): Promise<void> {
    try {
      const info: LicenseResult = await (window as any).api.getLicenseInfo();
      if (info?.hasAccess) {
        log('✅ Ліцензія дійсна - показуємо головний застосунок');
        this.showMainApp();
      } else {
        log('⚠️ Ліцензія відсутня - показуємо License Gate');
        this.showLicenseGate();
      }
    } catch (error) {
      console.error('Помилка перевірки ліцензії:', error);
      this.showLicenseGate();
    }
  }

  /**
   * Показати License Gate (блокувальний екран)
   * FIXED: Приховує основний контент до активації ліцензії
   * @private
   */
  private showLicenseGate(): void {
    const gate = byId('license-gate');
    if (gate) {
      gate.style.display = 'flex';
    }
  }

  /**
   * Показати основний застосунок
   * FIXED: Приховує License Gate, NavigationService сам керує routes
   * @private
   */
  private showMainApp(): void {
    const gate = byId('license-gate');
    if (gate) {
      gate.style.display = 'none';
    }
    // Завантажуємо інформацію про ліцензію для updates секції
    this.loadLicenseInfo();
  }

  /**
   * Активація ліцензії через License Gate
   * FIXED: Валідація та активація з переходом до основного інтерфейсу
   * @private
   */
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
      const result: LicenseResult = await (window as any).api.setLicenseKey(key);
      if (result.hasAccess) {
        this.updateGateStatus(`Ліцензія активована успішно!`, 'valid');
        log('✅ Ліцензію активовано через Gate');
        // Затримка для показу успішного повідомлення
        setTimeout(() => {
          this.showMainApp();
        }, 1000);
      } else {
        this.updateGateStatus(result.reason || 'Невірний ліцензійний ключ', 'invalid');
        log('❌ Невдала спроба активації через Gate');
      }
    } catch (error) {
      console.error('Помилка активації ліцензії:', error);
      this.updateGateStatus('Помилка з\'єднання', 'invalid');
    }
  }

  /**
   * Оновлення статусу в License Gate
   * FIXED: Застосовує CSS класи для стану в gate
   * @private
   */
  private updateGateStatus(message: string, state: LicenseStatus): void {
    const statusDiv = byId('gate-license-status');
    if (!statusDiv) return;
    
    statusDiv.textContent = message;
    statusDiv.className = `license-status ${state}`;
  }

  /**
   * Публічний метод для показу основного застосунку
   * FIXED: Використовується іншими модулями для розблокування UI
   * @public
   */
  public unlockApp(): void {
    this.showMainApp();
  }

  /**
   * Публічний метод для блокування застосунку
   * FIXED: Використовується для повернення до License Gate
   * @public
   */
  public lockApp(): void {
    this.showLicenseGate();
  }
}
