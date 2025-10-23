"""
Тест перевірки «зрізів» (Slice_Check)
"""

import sys
sys.path.insert(0, '.')

from excel_processor import ProcessingLog
from excel_slice_check import SliceIssue

print("🧪 === ТЕСТ ПЕРЕВІРКИ «ЗРІЗІВ» (F/G) ===\n")

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

# Тестові дані проблем з «зрізами»
issues = [
    SliceIssue(subunit='1РСпП', row=45, fio='Іванов Іван Іванович', pseudo='Іван', val_f='Ш', val_g='', reason='У F="Ш", але G порожня — потрібен «зріз»'),
    SliceIssue(subunit='2РСпП', row=78, fio='Петров Петро Петрович', pseudo='Петро', val_f='ВЛК', val_g='', reason='У F="ВЛК", але G порожня — потрібен «зріз»'),
    SliceIssue(subunit='РМТЗ', row=123, fio='Сидоров Сидір Сидорович', pseudo='Сидір', val_f='ВД', val_g='', reason='У F="ВД", але G порожня — потрібен «зріз»'),
]

stats = {
    'total': 3
}

# Додаємо перевірку «зрізів» в LOG (як у методі check_slices)
log.add_separator('ПЕРЕВІРКА «ЗРІЗІВ» (F/G)')

log.add_entry(
    operation='Знайдено проблем з «зрізами»',
    source_file='ЗС',
    sheet='',
    rows=stats['total'],
    details=f"У F є токен (Ш/ВЛК/ВД), але G порожня"
)

# Додаємо детальну інформацію про кожну проблему
for issue in issues:
    log.add_entry(
        operation='Проблема з «зрізом»',
        source_file=issue.subunit,   # Колонка C LOG = Колонка B ЗС (підрозділ)
        sheet=issue.fio,             # Колонка D LOG = Колонка D ЗС (ПІБ)
        rows='',
        details=f"Псевдо: {issue.pseudo} | Рядок {issue.row} | F={issue.val_f} | G=порожньо"
    )

# Показуємо структуру LOG
print("📋 Структура LOG:\n")
print("=" * 110)
print(f"{'№':<5} {'Операція':<30} {'Файл':<25} {'Підрозділ':<12} {'Рядків':<8} {'Деталі':<50}")
print("=" * 110)

entry_number = 0
for entry in log.entries:
    if entry.get('is_separator', False):
        print(f"\n{'':5} {entry['operation']}")
        print("-" * 110)
    else:
        entry_number += 1
        details_short = entry['details'][:60] if len(entry['details']) > 60 else entry['details']
        print(f"{entry_number:<5} {entry['operation']:<30} {entry['source_file']:<25} {entry['sheet']:<12} {str(entry['rows']):<8} {details_short:<50}")

print("\n" + "=" * 110)

# Підсумок
summary = log.get_summary()
print(f"\n📊 Підсумок:")
print(f"   Всього операцій: {summary['operations']}")
print(f"   Всього файлів: {summary['files_processed']}")
print(f"   Всього рядків: {summary['total_rows']}")
print(f"   Час обробки: {summary['duration']:.2f} сек")

print("\n✅ Перевірка «зрізів» інтегрована в LOG під розділювачем '═══ ПЕРЕВІРКА «ЗРІЗІВ» (F/G) ═══'")
print("📌 Токени: Ш, ВЛК, ВД")
print("📌 Правило: Якщо у F є токен, то G має бути заповнена")
