# 🐍 Python Excel Processor

## Встановлення Python (для розробки)

1. **Встановіть Python 3.10+**
   - Завантажте з https://www.python.org/downloads/
   - Під час встановлення поставте галочку "Add Python to PATH"

2. **Встановіть залежності:**
   ```bash
   cd python
   pip install -r requirements.txt
   ```

3. **Перевірте встановлення:**
   ```bash
   python excel_processor.py
   ```
   Має вивести помилку про відсутність конфігурації (це нормально)

---

## Тестування скрипта

Створіть файл `test_config.json`:

```json
{
  "destination_file": "C:\\path\\to\\destination.xlsx",
  "source_files": [
    "C:\\path\\to\\source1.xlsx",
    "C:\\path\\to\\source2.xlsx"
  ],
  "sheets": [
    {
      "name": "ЗС",
      "key_column": "B",
      "data_columns": ["C", "D", "E", "F", "G", "H"],
      "blacklist": ["упр", "п"]
    },
    {
      "name": "БЗ",
      "key_column": "C",
      "data_columns": ["D", "E", "F", "G", "H"],
      "blacklist": []
    }
  ]
}
```

Запустіть:
```bash
python excel_processor.py < test_config.json
```

---

## Пакування з Electron

### Варіант 1: Користувач встановлює Python сам

**Переваги:** Простіше, менший розмір програми  
**Недоліки:** Користувач має встановити Python

### Варіант 2: Упакувати Python разом (РЕКОМЕНДОВАНО)

Використаємо `python-embed`:

1. Завантажте Python Embeddable Package:
   - https://www.python.org/downloads/windows/
   - `python-3.10.x-embed-amd64.zip`

2. Розпакуйте в `resources/python/`

3. Встановіть openpyxl в embedded Python:
   ```bash
   resources/python/python.exe -m pip install openpyxl
   ```

4. Оновіть `electron-builder.yml`:
   ```yaml
   extraResources:
     - from: "python"
       to: "python"
   ```

---

## Переваги Python рішення

✅ **Надійність:** openpyxl - найкраща бібліотека для Excel  
✅ **Без XML помилок:** Коректно обробляє складні файли  
✅ **Зберігає форматування:** Не пошкоджує стилі  
✅ **Швидкість:** Швидше ніж ExcelJS  
✅ **Тестування:** Легко тестувати окремо від Electron  

---

## Troubleshooting

### Python not found
```bash
# Перевірте чи Python в PATH
python --version

# Якщо ні - встановіть Python знову з галочкою "Add to PATH"
```

### Module not found
```bash
# Переустановіть залежності
pip uninstall openpyxl
pip install openpyxl==3.1.2
```

### Permission denied
```bash
# Запустіть CMD/PowerShell як адміністратор
```
