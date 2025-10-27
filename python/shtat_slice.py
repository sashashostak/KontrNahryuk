"""
Shtat Slice - Нарізка штатно-посадового списку по підрозділах
Розрізає Excel файл зі штаткою на окремі файли для кожного підрозділу

Логіка нарізки:
Для кожного підрозділу визначено:
- Назву файлу
- Колонку для пошуку
- Текст для пошуку в колонці
- Додаткову колонку J для деяких підрозділів

Створюються файли:
1 РСпП, 2 РСпП, 3 РСпП, РВП, МБ, РБпС, ВРЕБ, РМТЗ, МП, ВІ, ВЗ, ВРСП

Кожен файл містить:
- Рядки відповідного підрозділу з повним форматуванням
- Перший рядок (заголовок) закріплений
- Фільтри на всіх колонках
"""

import sys
import json
import os
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
import win32com.client
from datetime import datetime
import traceback


# Конфігурація підрозділів
SUBUNITS_CONFIG = [
    {
        'file_name': '1 РСпП',
        'search_column': 'C',  # колонка C (3)
        'search_text': '1 рота спеціального призначення (на бронетранспортерах)',
        'additional_column': 'J',  # колонка J (10)
        'additional_text': '1 рот'
    },
    {
        'file_name': '2 РСпП',
        'search_column': 'C',
        'search_text': '2 рота спеціального призначення (на бронетранспортерах)',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': '3 РСпП',
        'search_column': 'C',
        'search_text': '3 рота спеціального призначення (на бронетранспортерах)',
        'additional_column': 'J',
        'additional_text': '3 рот'
    },
    {
        'file_name': 'РВП',
        'search_column': 'C',
        'search_text': 'Рота вогневої підтримки спеціального призначення',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': 'МБ',
        'search_column': 'C',
        'search_text': 'мінометна батарея',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': 'РБпС',
        'search_column': 'C',
        'search_text': 'Рота безпілотних систем',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': 'ВРЕБ',
        'search_column': 'D',  # колонка D (4)
        'search_text': 'Взвод радіоелектронної боротьби',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': 'РМТЗ',
        'search_column': 'C',
        'search_text': 'Рота матеріально-технічного забезпечення',
        'additional_column': 'J',
        'additional_text': 'РМТЗ',
        'additional_conditions': [
            {
                'column': 'D',
                'text': 'S-4 (логістика)',
                'force_secondary': True
            }
        ]
    },
    {
        'file_name': 'МП',
        'search_column': 'D',
        'search_text': 'медичний пункт',
        'additional_column': 'J',
        'additional_text': 'медпункт'
    },
    {
        'file_name': 'ВІ',
        'search_column': 'E',  # колонка E (5)
        'search_text': 'відділення інструкторів',
        'additional_column': None,
        'additional_text': None
    },
    {
        'file_name': 'ВЗ',
        'search_column': 'D',
        'search_text': "взвод зв'язку",
        'additional_column': None,
        'additional_text': None,
        'additional_conditions': [
            {
                'column': 'D',
                'text': "S-6 (зв'язок)",
                'force_secondary': True
            }
        ]
    },
    {
        'file_name': 'ВРСП',
        'search_column': 'D',
        'search_text': 'взвод розвідки спеціального призначення',
        'additional_column': None,
        'additional_text': None
    }
]


def col_letter_to_number(letter: str) -> int:
    """Конвертує букву колонки в номер (A=1, B=2, ..., Z=26, AA=27)"""
    result = 0
    for char in letter.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


class ShtatSlicer:
    """Клас для нарізки штатно-посадового списку"""
    
    def __init__(self, input_file: str, output_folder: str):
        self.input_file = input_file
        self.output_folder = output_folder
        self.excel = None
        self.wb = None
        
    def log(self, message: str):
        """Виведення повідомлення в stdout"""
        print(message, flush=True)
        
    def slice(self) -> Dict:
        """
        Головна функція нарізки
        
        Returns:
            Dict з результатами обробки
        """
        try:
            # Перевірка існування файлу
            if not os.path.exists(self.input_file):
                raise FileNotFoundError(f"Файл не знайдено: {self.input_file}")
            
            # Перевірка та створення папки виводу
            os.makedirs(self.output_folder, exist_ok=True)
            
            self.log("🔌 Підключення до Excel...")
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False
            
            self.log(f"📂 Відкриття файлу: {self.input_file}")
            self.wb = self.excel.Workbooks.Open(self.input_file)
            
            # Спочатку шукаємо аркуш "ЗС"
            ws = self._find_sheet("ЗС")
            if ws is None:
                # Якщо не знайдено "ЗС", використовуємо перший аркуш
                self.log("⚠️ Аркуш 'ЗС' не знайдено, використовується перший аркуш")
                ws = self.wb.Worksheets(1)
                self.log(f"📊 Використовується аркуш: '{ws.Name}'")
            else:
                self.log("📊 Знайдено аркуш 'ЗС'")
            
            self.log("📊 Аналіз структури файлу...")
            
            # Групуємо рядки по підрозділах
            subunits_data = self._group_by_subunits(ws)
            
            if not subunits_data:
                raise ValueError("Не знайдено жодного підрозділу у файлі")
            
            self.log(f"✅ Знайдено {len(subunits_data)} підрозділів")
            
            # Створюємо окремий файл для кожного підрозділу
            created_files = []
            for subunit_name, rows_info in subunits_data.items():
                output_file = self._create_subunit_file(subunit_name, rows_info, ws)
                created_files.append(output_file)
                self.log(f"   ✓ {subunit_name}: {len(rows_info['rows'])} рядків → {Path(output_file).name}")
            
            # Закриваємо вхідний файл
            self.wb.Close(SaveChanges=False)
            self.excel.Quit()
            
            return {
                'success': True,
                'subunits_count': len(subunits_data),
                'files_created': len(created_files),
                'output_folder': self.output_folder,
                'files': created_files,
                'message': f'Створено {len(created_files)} файлів для {len(subunits_data)} підрозділів'
            }
            
        except Exception as e:
            self.log(f"❌ ПОМИЛКА: {str(e)}")
            self.log(traceback.format_exc())
            
            # Закриваємо Excel якщо відкритий
            try:
                if self.wb:
                    self.wb.Close(SaveChanges=False)
                if self.excel:
                    self.excel.Quit()
            except:
                pass
            
            return {
                'success': False,
                'error': str(e),
                'traceback': traceback.format_exc()
            }
    
    def _find_sheet(self, name: str):
        """Знайти аркуш за назвою (без урахування регістру та пробілів)"""
        name_norm = name.strip().lower()
        for ws in self.wb.Worksheets:
            if ws.Name.strip().lower() == name_norm:
                return ws
        return None
    
    def _normalize_text(self, text) -> str:
        """Нормалізація тексту для порівняння"""
        if text is None:
            return ""
        return str(text).strip().lower()
    
    def _group_by_subunits(self, ws) -> Dict[str, Dict[str, Any]]:
        """
        Групування рядків по підрозділах згідно конфігурації
        
        Args:
            ws: Worksheet об'єкт
            
        Returns:
            Dict[назва_підрозділу, {'rows': [...], 'order_map': {...}}]
        """
        subunits_data = {
            config['file_name']: {
                'primary': [],
                'secondary': [],
                'forced': []
            }
            for config in SUBUNITS_CONFIG
        }
        
        # Знаходимо останній заповнений рядок
        last_row = ws.UsedRange.Rows.Count
        
        self.log(f"📋 Обробка рядків 1-{last_row}...")
        
        # Обробляємо кожен рядок (пропускаємо заголовок)
        for row in range(2, last_row + 1):
            # Перевіряємо кожен підрозділ з конфігурації
            for config in SUBUNITS_CONFIG:
                file_name = config['file_name']
                search_col_letter = config['search_column']
                search_text = config['search_text']
                additional_col = config.get('additional_column')
                additional_text = config.get('additional_text')
                additional_conditions = config.get('additional_conditions', [])

                data_entry = subunits_data[file_name]
                primary_rows_ref = data_entry['primary']
                secondary_rows_ref = data_entry['secondary']
                forced_rows_ref = data_entry['forced']

                # Конвертуємо букву колонки в номер
                search_col_num = col_letter_to_number(search_col_letter)

                # Читаємо значення основної колонки
                cell_value = ws.Cells(row, search_col_num).Value
                cell_text = self._normalize_text(cell_value)
                search_text_norm = self._normalize_text(search_text)

                # Перевіряємо основну умову
                main_match = search_text_norm in cell_text

                # Основні збіги залишаються на початку
                if main_match:
                    if row not in primary_rows_ref and row not in forced_rows_ref:
                        primary_rows_ref.append(row)
                        self.log(f"   → {file_name}: рядок {row} (основна колонка)")

                # Додатковий пошук — накопичуємо рядки для кінця списку
                if additional_col and additional_text:
                    add_col_num = col_letter_to_number(additional_col)
                    add_cell_value = ws.Cells(row, add_col_num).Value
                    add_cell_text = self._normalize_text(add_cell_value)
                    add_text_norm = self._normalize_text(additional_text)

                    if add_text_norm in add_cell_text:
                        if row not in primary_rows_ref and row not in secondary_rows_ref and row not in forced_rows_ref:
                            secondary_rows_ref.append(row)
                            self.log(f"   → {file_name}: рядок {row} (додаткова колонка {additional_col}, додано в кінець)")

                for condition in additional_conditions:
                    cond_column = condition.get('column')
                    cond_text = condition.get('text')
                    force_secondary = condition.get('force_secondary', False)

                    if not cond_column or not cond_text:
                        continue

                    cond_col_num = col_letter_to_number(cond_column)
                    cond_cell_value = ws.Cells(row, cond_col_num).Value
                    cond_cell_text = self._normalize_text(cond_cell_value)
                    cond_text_norm = self._normalize_text(cond_text)

                    if cond_text_norm in cond_cell_text:
                        if force_secondary:
                            removed_primary = False
                            removed_secondary = False

                            if row in primary_rows_ref:
                                primary_rows_ref.remove(row)
                                removed_primary = True

                            if row in secondary_rows_ref:
                                secondary_rows_ref.remove(row)
                                removed_secondary = True

                            if row not in forced_rows_ref:
                                forced_rows_ref.append(row)

                            if removed_primary or removed_secondary:
                                self.log(f"   → {file_name}: рядок {row} (переміщено в кінець через колонку {cond_column})")
                            else:
                                self.log(f"   → {file_name}: рядок {row} (колонка {cond_column}, додано в кінець)")
                        else:
                            if row not in primary_rows_ref and row not in secondary_rows_ref and row not in forced_rows_ref:
                                secondary_rows_ref.append(row)
                                self.log(f"   → {file_name}: рядок {row} (додаткова умова колонка {cond_column}, додано в кінець)")

        final_data = {}
        for name, data in subunits_data.items():
            primary_rows = sorted(data['primary'])
            secondary_rows = sorted(row for row in data['secondary'] if row not in primary_rows)
            forced_rows = sorted(row for row in data['forced'] if row not in primary_rows and row not in secondary_rows)

            ordered_rows = primary_rows + secondary_rows + forced_rows
            if ordered_rows:
                order_map = {
                    row_idx: order
                    for order, row_idx in enumerate(ordered_rows, start=1)
                }
                final_data[name] = {
                    'rows': ordered_rows,
                    'order_map': order_map
                }

        return final_data
    
    def _create_subunit_file(self, subunit_name: str, rows_info: Dict[str, Any], source_ws) -> str:
        """
        Створення окремого файлу для підрозділу з повним форматуванням
        
        Args:
            subunit_name: Назва підрозділу
            rows_info: Дані про рядки (впорядкований список та карта порядку)
            source_ws: Вихідний worksheet
            
        Returns:
            Шлях до створеного файлу
        """
        # Очищаємо назву підрозділу для імені файлу
        safe_name = self._sanitize_filename(subunit_name)

        # Формуємо ім'я файлу без дати (тільки назва підрозділу)
        output_filename = f"{safe_name}.xlsx"
        output_path = os.path.join(self.output_folder, output_filename)

        ordered_rows = rows_info['rows']
        order_map = rows_info['order_map']

        # Створюємо новий workbook на основі повної копії аркуша, щоб зберегти ВСЕ форматування
        self.log(f"      Копіювання аркуша з повним форматуванням...")
        source_ws.Copy()
        new_wb = self.excel.ActiveWorkbook
        new_ws = new_wb.Worksheets(1)
        new_ws.Name = safe_name

        # Додаємо службовий стовпець із порядком рядків
        temp_col_index = new_ws.UsedRange.Columns.Count + 1
        new_ws.Cells(1, temp_col_index).Value = "_order"
        for row_idx, order_value in order_map.items():
            new_ws.Cells(row_idx, temp_col_index).Value = order_value

        # Видаляємо усі рядки, які не належать підрозділу (крім заголовка)
        keep_rows = {1}
        keep_rows.update(ordered_rows)
        last_row_in_copy = new_ws.UsedRange.Rows.Count
        self.log("      Видалення зайвих рядків...")
        for row_idx in range(last_row_in_copy, 0, -1):
            if row_idx not in keep_rows:
                new_ws.Rows(row_idx).Delete()

        # Сортуємо залишені рядки за службовим порядком
        last_row_after_cleanup = new_ws.UsedRange.Rows.Count
        sort_range = new_ws.Range(new_ws.Cells(1, 1), new_ws.Cells(last_row_after_cleanup, temp_col_index))
        sort_range.Sort(Key1=new_ws.Cells(1, temp_col_index), Order1=1, Header=1)

        # Видаляємо службовий стовпець
        new_ws.Columns(temp_col_index).Delete()

        # Після видалення та сортування оновлюємо дані про кількість рядків/колонок
        last_col = new_ws.UsedRange.Columns.Count
        data_row_count = len(ordered_rows)

        # Розкриваємо усі рядки на випадок, якщо вони були приховані в оригіналі
        new_ws.Rows(f"1:{data_row_count + 1}").EntireRow.Hidden = False

        # Оновлюємо порядкові номери у колонці A (починаючи з 1 для першого рядка даних)
        self.log("      Встановлення послідовної нумерації у колонці A...")
        if data_row_count > 0:
            for index in range(1, data_row_count + 1):
                new_ws.Cells(index + 1, 1).Value = index
        else:
            # Якщо даних немає – очищаємо номери, якщо вони залишились
            new_ws.Cells(2, 1).Value = None

        # Виходимо з режиму копіювання аби Excel не фіксував весь діапазон
        self.excel.CutCopyMode = False

        # Встановлюємо автофільтр на всі колонки
        self.log("      Встановлення автофільтрів...")
        if data_row_count > 0:  # Є хоча б один рядок даних
            last_data_row = new_ws.UsedRange.Rows.Count
            filter_range = new_ws.Range(new_ws.Cells(1, 1), new_ws.Cells(last_data_row, last_col))
            filter_range.AutoFilter()

        # Повертаємось на початок листа перед збереженням
        new_ws.Range("A1").Select()

        # Зберігаємо файл без закріплення
        self.log(f"      Збереження файлу без закріплення: {output_filename}")
        new_wb.SaveAs(output_path)

        if data_row_count > 0:
            # Закріплюємо перший рядок після початкового збереження
            self.log("      Закріплення першого рядка...")
            self.excel.CutCopyMode = False
            new_wb.Activate()
            new_ws.Activate()

            window = self.excel.ActiveWindow
            if window is not None:
                try:
                    window.Split = False
                except AttributeError:
                    pass

                window.FreezePanes = False
                window.SplitRow = 0
                window.SplitColumn = 0

                new_ws.Range("A2").Select()
                window.SplitRow = 1
                window.SplitColumn = 0
                window.FreezePanes = True

            # Повертаємось до комірки A1
            new_ws.Range("A1").Select()
        else:
            self.log("      Закріплення пропущено (немає рядків даних)")

        # Фінальне збереження та закриття файлу
        self.log("      Фінальне збереження файлу...")
        new_wb.Save()
        new_wb.Close(SaveChanges=False)
        
        return output_path
    
    def _sanitize_filename(self, name: str) -> str:
        """Очистка назви для використання в імені файлу"""
        # Забороненні символи для імен файлів
        invalid_chars = '<>:"/\\|?*'
        
        safe_name = name
        for char in invalid_chars:
            safe_name = safe_name.replace(char, '_')
        
        # Обмежуємо довжину
        max_length = 100
        if len(safe_name) > max_length:
            safe_name = safe_name[:max_length]
        
        return safe_name.strip()


def main():
    """Головна функція скрипту"""
    try:
        # Читаємо конфігурацію зі stdin
        config_json = sys.stdin.read()
        config = json.loads(config_json)
        
        input_file = config.get('input_file')
        output_folder = config.get('output_folder')
        
        if not input_file:
            raise ValueError("Не вказано вхідний файл")
        if not output_folder:
            raise ValueError("Не вказано папку виводу")
        
        print(f"\n{'='*60}")
        print("🔪 ШТАТ_SLICE - Нарізка штатки для підрозділів")
        print(f"{'='*60}\n")
        print(f"📥 Вхідний файл: {input_file}")
        print(f"📤 Папка виводу: {output_folder}\n")
        
        # Виконуємо нарізку
        slicer = ShtatSlicer(input_file, output_folder)
        result = slicer.slice()
        
        # Виводимо результат у форматі JSON
        print(f"\n__RESULT__{json.dumps(result, ensure_ascii=False)}__END__")
        
        if result['success']:
            sys.exit(0)
        else:
            sys.exit(1)
            
    except Exception as e:
        error_result = {
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }
        print(f"\n__RESULT__{json.dumps(error_result, ensure_ascii=False)}__END__")
        sys.exit(1)


if __name__ == "__main__":
    main()
