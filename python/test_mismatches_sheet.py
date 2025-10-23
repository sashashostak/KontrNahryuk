"""
Тест створення листа з невідповідностями
"""

from openpyxl import Workbook
from excel_processor import ExcelProcessor
from excel_mismatches import MismatchEntry

def test_mismatches_sheet():
    """Тестування створення листа НЕВІДПОВІДНОСТІ"""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "ЗС"
    
    # Створюємо простий файл
    test_file = "d:/Додатки/1SForS1/python/test_mismatches.xlsx"
    wb.save(test_file)
    
    # Створюємо процесор та завантажуємо файл
    processor = ExcelProcessor()
    processor.destination_path = test_file
    processor.destination_wb = wb
    
    # Створюємо тестові невідповідності
    mismatches = [
        MismatchEntry("ЗС", "D10", "Іванов Іван Іванович", "відсутнє у БЗ"),
        MismatchEntry("ЗС", "D25", "Петров Петро Петрович", "відсутнє у БЗ"),
        MismatchEntry("ЗС", "D42", "Сидоров Сидір Сидорович", "відсутнє у БЗ"),
        MismatchEntry("БЗ", "E15", "Коваль Василь Васильович", "відсутнє у ЗС"),
        MismatchEntry("БЗ", "E33", "Мельник Олександр Олександрович", "відсутнє у ЗС"),
    ]
    
    stats = {
        'total': 5,
        's1_missing_in_s2': 3,
        's2_missing_in_s1': 2
    }
    
    print("🧪 === ТЕСТ СТВОРЕННЯ ЛИСТА НЕВІДПОВІДНОСТІ ===\n")
    print(f"📊 Тестові дані:")
    print(f"   Всього невідповідностей: {stats['total']}")
    print(f"   ЗС → відсутні в БЗ: {stats['s1_missing_in_s2']}")
    print(f"   БЗ → відсутні в ЗС: {stats['s2_missing_in_s1']}")
    
    print(f"\n📋 Список невідповідностей:")
    for i, m in enumerate(mismatches, 1):
        print(f"   {i}. {m}")
    
    # Створюємо лист
    processor._create_mismatches_sheet(mismatches, stats)
    
    # Зберігаємо
    wb.save(test_file)
    
    print(f"\n✅ Тестовий файл збережено: {test_file}")
    print(f"📂 Відкрийте файл для перегляду листа НЕВІДПОВІДНОСТІ")

if __name__ == '__main__':
    test_mismatches_sheet()
