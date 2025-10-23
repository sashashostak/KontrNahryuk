"""
Тест інтеграції невідповідностей в LOG
"""

from excel_processor import ProcessingLog
from excel_mismatches import MismatchEntry

print("🧪 === ТЕСТ ІНТЕГРАЦІЇ НЕВІДПОВІДНОСТЕЙ В LOG ===\n")

# Створюємо лог
log = ProcessingLog()

# Додаємо звичайні операції
log.add_separator('КОПІЮВАННЯ ДАНИХ')
log.add_entry(
    operation='Копіювання даних',
    source_file='1БрТА_Січень.xlsx',
    sheet='1241',
    rows=45,
    details='Скопійовано з B4:F48'
)

log.add_separator('САНІТИЗАЦІЯ ДАНИХ')
log.add_entry(
    operation='Санітизація даних (A:H)',
    source_file='',
    sheet='ЗС',
    rows=23,
    details='Очищено 23 клітинок з 1850 (колонки A:H)'
)

# Тестові дані невідповідностей
mismatches = [
    MismatchEntry(sheet='ЗС', cell_addr='D10', value='Іванов Іван Іванович', reason='відсутнє у БЗ'),
    MismatchEntry(sheet='ЗС', cell_addr='D25', value='Петров Петро Петрович', reason='відсутнє у БЗ'),
    MismatchEntry(sheet='БЗ', cell_addr='E15', value='Коваль Василь Васильович', reason='відсутнє у ЗС'),
]

stats = {
    'total': 3,
    's1_missing_in_s2': 2,
    's2_missing_in_s1': 1
}

# Додаємо невідповідності в LOG (як у методі check_mismatches)
log.add_separator('ПЕРЕВІРКА НЕВІДПОВІДНОСТЕЙ')

log.add_entry(
    operation='Знайдено невідповідностей',
    source_file='',
    sheet='ЗС / БЗ',
    rows=stats['total'],
    details=f"ЗС→БЗ: {stats['s1_missing_in_s2']}, БЗ→ЗС: {stats['s2_missing_in_s1']}"
)

# Додаємо детальну інформацію про кожну невідповідність
for mismatch in mismatches:
    log.add_entry(
        operation='Невідповідність',
        source_file='',
        sheet=mismatch.sheet,
        rows='',
        details=f"{mismatch.cell_addr}: «{mismatch.value}» — {mismatch.reason}"
    )

# Показуємо структуру LOG
print("📋 Структура LOG:\n")
print("=" * 100)
print(f"{'№':<5} {'Операція':<30} {'Файл':<25} {'Лист':<10} {'Рядків':<8} {'Деталі':<30}")
print("=" * 100)

entry_number = 0
for entry in log.entries:
    if entry.get('is_separator', False):
        print(f"\n{'':5} {entry['operation']}")
        print("-" * 100)
    else:
        entry_number += 1
        print(f"{entry_number:<5} {entry['operation']:<30} {entry['source_file']:<25} {entry['sheet']:<10} {str(entry['rows']):<8} {entry['details'][:50]:<30}")

print("\n" + "=" * 100)

# Підсумок
summary = log.get_summary()
print(f"\n📊 Підсумок:")
print(f"   Всього операцій: {summary['operations']}")
print(f"   Всього файлів: {summary['files_processed']}")
print(f"   Всього рядків: {summary['total_rows']}")
print(f"   Час обробки: {summary['duration']:.2f} сек")

print("\n✅ Невідповідності тепер інтегровані в LOG під розділювачем '═══ ПЕРЕВІРКА НЕВІДПОВІДНОСТЕЙ ═══'")
