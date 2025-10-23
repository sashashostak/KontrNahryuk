"""
Excel Sanitizer - Очистка та нормалізація даних у Excel файлах
Адаптація VBA ExcelSanitizer v2.6 для Python + openpyxl
"""

import re
from typing import Any, Optional, Tuple
from datetime import datetime, date
from openpyxl.cell.cell import Cell


class ExcelSanitizer:
    """Клас для санітизації Excel даних"""
    
    # Налаштування
    TOUCH_FORMULAS = False  # Не чіпати формули
    CONSERVATIVE_CLEAR = True  # Консервативне видалення
    CLEAR_PREFIX_APOSTROPHE = True  # Прибирати префікс '
    
    # Символи для очистки
    NBSP = '\u00A0'  # Non-breaking space
    ZWSP = '\u200B'  # Zero-width space
    BOM = '\uFEFF'   # Byte order mark
    NARROW_NBSP = '\u202F'  # Narrow no-break space
    EM_DASH = '\u2014'
    EN_DASH = '\u2013'
    MINUS_SIGN = '\u2212'
    NON_BREAKING_HYPHEN = '\u2011'
    
    def __init__(self):
        self.changes_count = 0
        self.preview_changes = []
        self.preview_limit = 100
    
    def sanitize_cell(self, cell: Cell) -> bool:
        """
        Санітизувати одну клітинку
        
        Returns:
            True якщо були зміни
        """
        # Пропускаємо формули якщо налаштовано
        if not self.TOUCH_FORMULAS and cell.data_type == 'f':
            return False
        
        old_value = cell.value
        had_changes = False
        
        # Обробляємо значення
        new_value = self._sanitize_value(old_value, cell)
        
        if not self._equal_values(old_value, new_value):
            # Скидаємо текстовий формат якщо це дата чи число
            if isinstance(new_value, (date, datetime, int, float)):
                if cell.number_format == '@':
                    cell.number_format = 'General'
            
            cell.value = new_value
            had_changes = True
            
            # Зберігаємо превʼю зміни
            if len(self.preview_changes) < self.preview_limit:
                self.preview_changes.append({
                    'cell': f"R{cell.row}C{cell.column}",
                    'old': self._format_value(old_value),
                    'new': self._format_value(new_value)
                })
        
        return had_changes
    
    def _sanitize_value(self, value: Any, cell: Cell) -> Any:
        """Санітизувати значення"""
        
        # Порожні та помилки — не чіпаємо
        if value is None or value == '':
            return value
        
        # ТЕКСТ
        if isinstance(value, str):
            original = value
            cleaned = self._clean_string(original)
            
            # Видаляємо апостроф на початку (якщо налаштування увімкнене)
            if self.CLEAR_PREFIX_APOSTROPHE and cleaned.startswith("'"):
                cleaned = cleaned[1:]
            
            # Якщо стало порожньо — вирішуємо, чи стирати
            if len(cleaned) == 0:
                if self.CONSERVATIVE_CLEAR:
                    if self._looks_blank(original):
                        return ''
                    else:
                        return value
                return ''
            
            # 1) Спробувати розпізнати як дату DMY (strict)
            parsed_date = self._try_parse_date_dmy_strict(cleaned)
            if parsed_date:
                return parsed_date
            
            # 2) Спробувати розпізнати як число
            parsed_number = self._try_parse_number(cleaned)
            if parsed_number is not None:
                return parsed_number
            
            return cleaned
        
        # ЧИСЛО - перевіряємо чи це не "злиплена" дата
        if isinstance(value, (int, float)):
            # Якщо формат клітинки схожий на дату
            if self._is_date_like_format(cell.number_format):
                digits_str = str(int(value))
                # Тільки 6 або 8 цифр
                if len(digits_str) in (6, 8):
                    parsed_date = self._try_parse_dmy_from_digits(digits_str)
                    if parsed_date:
                        return parsed_date
        
        return value
    
    def _clean_string(self, s: str) -> str:
        """Очистити рядок від сміттєвих символів"""
        # Видаляємо BOM, ZWSP, вузькі пробіли
        s = s.replace(self.BOM, '')
        s = s.replace(self.ZWSP, '')
        s = s.replace(self.NARROW_NBSP, '')
        s = s.replace(self.NBSP, ' ')
        
        # Нормалізуємо переноси рядків та табуляцію
        s = s.replace('\r', ' ')
        s = s.replace('\n', ' ')
        s = s.replace('\t', ' ')
        
        # Нормалізуємо тире/мінуси
        s = s.replace(self.NON_BREAKING_HYPHEN, '-')
        s = s.replace(self.EN_DASH, '-')
        s = s.replace(self.EM_DASH, '-')
        s = s.replace(self.MINUS_SIGN, '-')
        
        # Прибираємо множинні пробіли
        s = ' '.join(s.split())
        
        return s.strip()
    
    def _looks_blank(self, s: str) -> bool:
        """Перевірити чи рядок виглядає порожнім"""
        cleaned = self._clean_string(s)
        return len(cleaned) == 0
    
    def _try_parse_date_dmy_strict(self, s: str) -> Optional[date]:
        """
        Спробувати розпізнати дату у форматі DMY (strict)
        Вся комірка має бути датою, без тексту
        """
        s = s.strip()
        if not s:
            return None
        
        # Якщо є літери — не дата
        if any(c.isalpha() for c in s):
            return None
        
        # Дозволені тільки цифри та роздільники .-/ та пробіл
        if not all(c.isdigit() or c in '.-/ ' for c in s):
            return None
        
        # 1) Стандартні DMY з роздільниками (25.08.2025, 25/08/25, 25-08-2025)
        pattern = r'^\s*(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})\s*$'
        match = re.match(pattern, s)
        if match:
            day = int(match.group(1))
            month = int(match.group(2))
            year = int(match.group(3))
            return self._try_build_dmy(day, month, year)
        
        # 2) «Злиплі» DMY (25082025, 25.082025, 25-08-25)
        digits = ''.join(c for c in s if c.isdigit())
        if len(digits) in (6, 8):
            return self._try_parse_dmy_from_digits(digits)
        
        return None
    
    def _try_parse_dmy_from_digits(self, digits: str) -> Optional[date]:
        """Розпізнати дату з «злиплих» цифр"""
        if len(digits) == 8:
            day = int(digits[0:2])
            month = int(digits[2:4])
            year = int(digits[4:8])
        elif len(digits) == 6:
            day = int(digits[0:2])
            month = int(digits[2:4])
            year = int(digits[4:6])
        else:
            return None
        
        return self._try_build_dmy(day, month, year)
    
    def _try_build_dmy(self, day: int, month: int, year: int) -> Optional[date]:
        """Спробувати побудувати дату з DMY компонентів"""
        # Виправляємо 2-цифрові роки
        year = self._fix_year(year)
        
        # Перевірка діапазонів
        if month < 1 or month > 12 or day < 1 or day > 31:
            return None
        
        try:
            result = date(year, month, day)
            # Перевірка що дата валідна (врахування днів у місяці)
            if result.year == year and result.month == month and result.day == day:
                return result
        except ValueError:
            pass
        
        return None
    
    def _fix_year(self, year: int) -> int:
        """Виправити 2-цифрові роки (00-29 => 2000-2029, 30-99 => 1930-1999)"""
        if year < 100:
            if year <= 29:
                return 2000 + year
            else:
                return 1900 + year
        return year
    
    def _try_parse_number(self, s: str) -> Optional[float]:
        """Спробувати розпізнати число"""
        # Видаляємо пробіли та вузькі пробіли
        s = s.replace(' ', '')
        s = s.replace(self.NBSP, '')
        s = s.replace(self.NARROW_NBSP, '')
        
        # Нормалізуємо мінуси
        s = s.replace(self.MINUS_SIGN, '-')
        s = s.replace(self.EN_DASH, '-')
        s = s.replace(self.EM_DASH, '-')
        s = s.replace(self.NON_BREAKING_HYPHEN, '-')
        
        # Рахуємо крапки та коми
        dot_count = s.count('.')
        comma_count = s.count(',')
        
        # Немає роздільників
        if dot_count == 0 and comma_count == 0:
            try:
                return float(s)
            except ValueError:
                return None
        
        # Є і крапки, і коми - визначаємо десятковий роздільник
        if dot_count > 0 and comma_count > 0:
            last_dot = s.rfind('.')
            last_comma = s.rfind(',')
            
            if last_dot > last_comma:
                # Крапка - десятковий роздільник
                s = s.replace(',', '')
            else:
                # Кома - десятковий роздільник
                s = s.replace('.', '')
                s = s.replace(',', '.')
        else:
            # Тільки один тип роздільника
            if dot_count > 0:
                sep = '.'
                last_pos = s.rfind('.')
            else:
                sep = ','
                last_pos = s.rfind(',')
            
            # Евристика: якщо після останнього роздільника 3 цифри, 
            # і є щось перед ним - це тисячний роздільник
            if last_pos > 0 and last_pos < len(s) - 1:
                left_len = last_pos
                right_len = len(s) - last_pos - 1
                
                if right_len == 3 and left_len >= 1:
                    # Тисячний роздільник
                    s = s.replace(sep, '')
                else:
                    # Десятковий роздільник
                    if sep == ',':
                        s = s.replace(',', '.')
        
        try:
            return float(s)
        except ValueError:
            return None
    
    def _is_date_like_format(self, fmt: str) -> bool:
        """Перевірити чи формат схожий на дату"""
        if not fmt:
            return False
        fmt_lower = fmt.lower()
        return 'd' in fmt_lower and 'm' in fmt_lower
    
    def _equal_values(self, a: Any, b: Any) -> bool:
        """Порівняти значення"""
        if a is None and b is None:
            return True
        if type(a) != type(b):
            return False
        if isinstance(a, (date, datetime)) and isinstance(b, (date, datetime)):
            return a == b
        return a == b
    
    def _format_value(self, v: Any) -> str:
        """Форматувати значення для виведення"""
        if v is None:
            return '[empty]'
        if isinstance(v, bool):
            return f'[{v} | bool]'
        if isinstance(v, (date, datetime)):
            return f'[{v.strftime("%d.%m.%Y")} | date]'
        if isinstance(v, (int, float)):
            return f'[{v} | num]'
        if isinstance(v, str):
            return f'[{v} | txt]'
        return f'[{str(v)}]'


def sanitize_cells(cells, show_preview: bool = True) -> Tuple[int, int, list]:
    """
    Санітизувати набір клітинок
    
    Args:
        cells: Ітератор клітинок openpyxl
        show_preview: Показувати превʼю змін
    
    Returns:
        (total_cells, changed_cells, preview_changes)
    """
    sanitizer = ExcelSanitizer()
    total = 0
    changed = 0
    
    for cell in cells:
        total += 1
        if sanitizer.sanitize_cell(cell):
            changed += 1
    
    preview = sanitizer.preview_changes if show_preview else []
    return total, changed, preview
