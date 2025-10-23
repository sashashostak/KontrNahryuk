"""
Тест модуля excel_sanitizer.py
Перевіряє основні функції санітизації даних
"""

from openpyxl import Workbook
from excel_sanitizer import ExcelSanitizer

def test_sanitizer():
    """Тестування основних функцій санітизації"""
    
    # Створюємо тестову книгу
    wb = Workbook()
    ws = wb.active
    ws.title = "Тест"
    
    # Тестові дані
    test_data = [
        ("A1", "25.08.2025", "DMY дата з крапками"),
        ("A2", "25/08/2025", "DMY дата з слешами"),
        ("A3", "25082025", "Stuck together дата"),
        ("A4", "  test  ", "Текст з пробілами"),
        ("A5", "test\u00A0word", "Текст з NBSP"),
        ("A6", "\uFEFFHello", "Текст з BOM"),
        ("A7", "1 234,56", "Число з пробілами"),
        ("A8", "1,234.56", "Число US формат"),
        ("A9", "'123", "Число з апострофом"),
        ("A10", "", "Порожня клітинка"),
    ]
    
    print("🧪 === ТЕСТУВАННЯ САНІТИЗАЦІЇ ===\n")
    
    sanitizer = ExcelSanitizer()
    
    for cell_ref, value, description in test_data:
        cell = ws[cell_ref]
        cell.value = value
        
        print(f"📌 {description}")
        print(f"   Клітинка: {cell_ref}")
        print(f"   До: '{value}' (тип: {type(value).__name__})")
        
        changed = sanitizer.sanitize_cell(cell)
        
        print(f"   Після: '{cell.value}' (тип: {type(cell.value).__name__})")
        print(f"   Змінено: {'✅ Так' if changed else '❌ Ні'}")
        print()
    
    print("✅ Тест завершено!")

if __name__ == '__main__':
    test_sanitizer()
