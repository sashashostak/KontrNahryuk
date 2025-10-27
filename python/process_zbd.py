"""
Process ZBD (ЖБД) - Обробка CSV файлів та створення Word документів
Створює Word документ з таблицею 3 колонки × 5 рядків:
- Рядок 1: Заголовки ("Дата, час", "Завдання військ...", "Примітка")
- Рядок 2: Нумерація колонок (1, 2, 3)
- Рядки 3-5: Дані з CSV файлу (до 3 рядків)
"""

import sys
import json
import io
import csv
from pathlib import Path
from typing import List, Dict, Any
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Встановлюємо UTF-8 для stdout/stderr на Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


class ZBDProcessor:
    """Процесор для створення Word документів ЖБД з CSV файлів"""

    def __init__(self):
        self.csv_data = []

    def read_csv(self, csv_path: str) -> List[List[str]]:
        """
        Читання CSV файлу

        Args:
            csv_path: Шлях до CSV файлу

        Returns:
            Список рядків CSV файлу
        """
        print(f"📖 Читання CSV файлу: {csv_path}")

        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                data = list(reader)

            print(f"✅ Прочитано {len(data)} рядків")
            return data

        except Exception as e:
            print(f"❌ Помилка читання CSV: {str(e)}", file=sys.stderr)
            raise

    def create_word_document(self, output_path: str) -> None:
        """
        Створення Word документу з таблицею 3×5 (шаблон)

        Args:
            output_path: Шлях для збереження Word документу
        """
        print(f"\n📝 Створення Word документу...")

        try:
            # Створюємо новий документ
            doc = Document()

            # Налаштовуємо поля сторінки
            print("📐 Налаштування полів сторінки...")
            section = doc.sections[0]
            section.top_margin = Inches(1.5 / 2.54)      # 1.5 см
            section.bottom_margin = Inches(1.5 / 2.54)   # 1.5 см
            section.left_margin = Inches(2.5 / 2.54)     # 2.5 см
            section.right_margin = Inches(1.5 / 2.54)    # 1.5 см

            # Додаємо таблицю 3 колонки × 5 рядків
            print("📊 Створення таблиці 3×5...")
            table = doc.add_table(rows=5, cols=3)
            table.style = 'Table Grid'

            # Встановлюємо ширину колонок
            print("📏 Налаштування ширини колонок...")
            column_widths = [
                2.48,   # Колонка 1: 2.48 см
                12.32,  # Колонка 2: 12.32 см
                1.99    # Колонка 3: 1.99 см
            ]

            for row in table.rows:
                for idx, cell in enumerate(row.cells):
                    # Встановлюємо ширину в сантиметрах (конвертуємо в дюйми)
                    cell.width = Inches(column_widths[idx] / 2.54)

            # Рядок 1: Заголовки колонок
            print("📝 Заповнення заголовків...")
            headers = ['Дата,\nчас', 'Завдання військ та стисле висвітлення ходу бойових дій', 'Примітка']
            for col_idx, header in enumerate(headers):
                cell = table.rows[0].cells[col_idx]

                # Вертикальне вирівнювання по центру
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Очищаємо комірку та додаємо новий параграф
                cell.text = ''
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Додаємо текст з форматуванням
                run = paragraph.add_run(header)
                run.font.size = Pt(12)
                run.font.name = 'Times New Roman'

            # Рядок 2: Нумерація колонок
            print("📝 Заповнення нумерації...")
            numbers = ['1', '2', '3']
            for col_idx, number in enumerate(numbers):
                cell = table.rows[1].cells[col_idx]
                cell.text = number

                if cell.paragraphs:
                    paragraph = cell.paragraphs[0]
                    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run(number)
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Рядок 3: Шаблон з фіксованими даними
            print("📝 Заповнення рядка 3 шаблонними даними...")
            row3_data = [
                '01.10.2025',
                'РВП СпП 2 БСпП продовжує виконання бойових завдань.\nСтан та положення – без змін',
                ''  # Примітка порожня
            ]

            for col_idx, cell_data in enumerate(row3_data):
                cell = table.rows[2].cells[col_idx]  # Рядок 3 (індекс 2)
                cell.text = cell_data

                # Форматування тексту
                if cell.paragraphs:
                    paragraph = cell.paragraphs[0]
                    if paragraph.runs:
                        run = paragraph.runs[0]
                    else:
                        run = paragraph.add_run(cell_data)
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'

            # Рядок 4: Залишаємо порожнім
            print("📝 Рядок 4 залишається порожнім")

            # Рядок 5: Командир та підпис
            print("📝 Заповнення рядка 5 (командир та підпис)...")
            row5_data = [
                '',  # Колонка 1 порожня
                'Командир РВП СпП 2 бСпП\nстарший лейтенант                                               Артем БУЛАВІН',
                ''  # Колонка 3 порожня
            ]

            for col_idx, cell_data in enumerate(row5_data):
                cell = table.rows[4].cells[col_idx]  # Рядок 5 (індекс 4)
                cell.text = cell_data

                # Форматування тексту
                if cell.paragraphs:
                    paragraph = cell.paragraphs[0]
                    if paragraph.runs:
                        run = paragraph.runs[0]
                    else:
                        run = paragraph.add_run(cell_data)
                    run.font.bold = True  # Жирний шрифт для командира
                    run.font.size = Pt(12)
                    run.font.name = 'Times New Roman'

            # Додаємо кордони до таблиці
            self._set_table_borders(table)

            # Зберігаємо документ
            print(f"💾 Збереження документу: {output_path}")
            doc.save(output_path)
            print("✅ Word документ успішно створено!")

        except Exception as e:
            print(f"❌ Помилка створення Word документу: {str(e)}", file=sys.stderr)
            raise

    def _set_table_borders(self, table):
        """
        Встановлює кордони для таблиці

        Args:
            table: Таблиця docx
        """
        tbl = table._element
        tblPr = tbl.tblPr

        # Створюємо елемент borders якщо його немає
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)

        # Встановлюємо всі кордони
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')  # Товщина кордону
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')  # Чорний колір
            tblBorders.append(border)

    def process(self, csv_path: str, output_path: str) -> Dict[str, Any]:
        """
        Головна функція обробки

        Args:
            csv_path: Шлях до CSV файлу
            output_path: Шлях для збереження Word документу

        Returns:
            Словник з результатами обробки
        """
        try:
            print("🚀 === ОБРОБКА ЖБД ===\n")

            # TODO: В майбутньому тут буде читання та обробка CSV файлу
            # Наразі створюємо шаблон з фіксованими даними
            print("📝 Створення шаблону Word документу...")

            # Створюємо Word документ
            self.create_word_document(output_path)

            print("\n✅ === ОБРОБКА ЗАВЕРШЕНА ===")

            return {
                'success': True,
                'output_path': output_path,
                'rows_processed': 1,  # 1 рядок з шаблонними даними
                'message': f'Успішно створено шаблон Word документу ЖБД'
            }

        except Exception as e:
            error_msg = f"Помилка обробки: {str(e)}"
            print(f"\n❌ {error_msg}", file=sys.stderr)
            return {
                'success': False,
                'error': error_msg
            }


def main():
    """Головна функція - приймає JSON конфігурацію через stdin"""

    try:
        # Читаємо конфігурацію з stdin
        config_json = sys.stdin.read()
        config = json.loads(config_json)

        csv_path = config.get('csv_path')
        output_path = config.get('output_path')

        # Валідація вхідних параметрів
        if not csv_path:
            raise ValueError("Не вказано шлях до CSV файлу")

        if not output_path:
            raise ValueError("Не вказано шлях для збереження результату")

        # Перевірка існування CSV файлу
        if not Path(csv_path).exists():
            raise FileNotFoundError(f"CSV файл не знайдено: {csv_path}")

        # Створюємо директорію для результату якщо не існує
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)

        # Обробка
        processor = ZBDProcessor()
        result = processor.process(csv_path, output_path)

        # Повертаємо результат у JSON форматі
        print(f"\n__RESULT__{json.dumps(result, ensure_ascii=False)}__END__")

        sys.exit(0 if result['success'] else 1)

    except Exception as e:
        error_msg = str(e)
        print(f"\n❌ КРИТИЧНА ПОМИЛКА: {error_msg}", file=sys.stderr)

        result = {
            'success': False,
            'error': error_msg
        }
        print(f"\n__RESULT__{json.dumps(result, ensure_ascii=False)}__END__")
        sys.exit(1)


if __name__ == '__main__':
    main()
