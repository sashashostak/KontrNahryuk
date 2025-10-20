/**
 * filePicker - Утиліти для стилізації file inputs
 * FIXED: Винесено з main.ts (рядки 269-298)
 * 
 * Відповідальність:
 * - Стилізація стандартних file inputs
 * - Відображення назви обраного файлу
 * - Обробка кліків на кастомну кнопку
 * - Підтримка placeholder тексту
 * 
 * @module filePicker
 */

/**
 * Прив'язка "красивого" file picker до input елемента
 * FIXED: Знаходить елементи та налаштовує обробники
 * 
 * @param id - ID input[type="file"] елемента
 * 
 * @example
 * ```html
 * <div class="file-picker">
 *   <input type="file" id="word-file" />
 *   <button class="file-btn">Обрати файл</button>
 *   <span class="file-name empty" data-placeholder="Файл не вибрано"></span>
 * </div>
 * ```
 * 
 * ```typescript
 * bindPrettyFile('word-file');
 * ```
 */
export function bindPrettyFile(id: string): void {
  const input = document.getElementById(id) as HTMLInputElement | null;
  if (!input) {
    console.warn(`File input with id "${id}" not found`);
    return;
  }

  // FIXED: Знаходимо батьківський контейнер та дочірні елементи
  const label = input.closest('.file-picker');
  const nameSpan = label?.querySelector('.file-name') as HTMLElement | null;
  const fileBtn = label?.querySelector('.file-btn') as HTMLElement | null;
  const placeholder = nameSpan?.getAttribute('data-placeholder') || 'Файл не вибрано';
  
  /**
   * Оновлення відображення назви файлу
   * FIXED: Показує назву файлу або placeholder
   * @private
   */
  const refresh = () => {
    const file = input.files?.[0];
    if (nameSpan) {
      nameSpan.textContent = file ? file.name : placeholder;
      nameSpan.classList.toggle('empty', !file);
    }
  };
  
  // FIXED: Обробка кліку на кастомну кнопку
  fileBtn?.addEventListener('click', (e) => {
    e.preventDefault();
    input.click();
  });
  
  // FIXED: Оновлення UI при зміні файлу
  input.addEventListener('change', refresh);
  
  // Початкове відображення
  refresh();
}

/**
 * Ініціалізація всіх file pickers на сторінці
 * FIXED: Викликає bindPrettyFile для існуючих ID
 * @public
 */
export function initializeFilePickers(): void {
  bindPrettyFile('word-file');
  bindPrettyFile('word-files');
  bindPrettyFile('order-word-file');
}
