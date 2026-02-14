# parse_specs_to_bom_many.py
# Основано на parse_spec_to_bom.py (актуальная версия пользователя)

from docx import Document
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


# ==========================
# НАСТРОЙКИ
# ==========================

OUTPUT_XLSX = "BOMs_parsed.xlsx"  # один общий выходной Excel

# Коллекция (упорядоченная) пар: (INPUT_DOCX, MODULE_CODE)
# Важно: порядок сохранится ровно как здесь.
SPECS = [
    ("КОР-01.00.000 Система связи модели DVS-21 Спецификация.docx", "КОР-01.00.000"),
    ("КОР-01.10.000 Кросс-плата. Спецификация.docx", "КОР-01.10.000"),
    ("КОР-01.20.000 Кросс-плата. Спецификация.docx", "КОР-01.20.000"),
    ("КОР-01.30.000 Кросс-плата. Спецификация.docx", "КОР-01.30.000"),
    ("КОР-01.31.000 Плата вспомогательная Спецификация.docx", "КОР-01.31.000"),
    ("КОР-01.40.000 Кросс-плата. Спецификация.docx", "КОР-01.40.000"),
    ("КОР-03.12.000 Модуль BusCPU. Спецификация.docx", "КОР-03.12.000"),
    ("КОР-03.13.000 Модуль связи Е1 Спецификация.docx", "КОР-03.13.000"),
    ("КОР-04.00.000 Пульт диспетчерский DTA-030 (спецификация).docx", "КОР-04.00.000"),
    ("КОР-04.10.000 Плата управления DTA-G Спецификация.docx", "КОР-04.10.000"),
    ("КОР-04.20.000 Плата кнопочная ТА-30Т-Т Спецификация.docx", "КОР-04.20.000"),
    ("КОР-05.00.000 Переговорное устройство WPS-04-25 Спецификация.docx", "КОР-05.00.000"),
    ("КОР-05.10.000 Плата управления WPS-04 Спецификация.docx", "КОР-05.10.000"),
    ("КОР-05.20.000 Плата усилителя. Спецификация.docx", "КОР-05.20.000")
   


    # добавляй дальше...
]


# ==========================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ==========================

def clean(text: str) -> str:
    """Убирает лишние пробелы"""
    return re.sub(r"\s+", " ", (text or "").strip())


quote_re = re.compile(r'["“”«»](.+?)["“”«»]')


def extract_manufacturer(text: str) -> str:
    """
    Производитель = последнее в кавычках
    """
    matches = quote_re.findall(text or "")
    return clean(matches[-1]) if matches else ""


# ==========================
# ОСНОВНОЙ ПАРСЕР
# (без изменений по сути)
# ==========================

def parse_spec(input_path, module_code):
    doc = Document(input_path)

    rows = []

    current_section = None

    # Буферы одной позиции
    buf_pos = None
    buf_qty = None

    buf_name = []
    buf_comment = []
    buf_designation = []

    def flush():
        """
        Сохраняет текущую позицию в результат
        """
        nonlocal buf_pos, buf_qty
        nonlocal buf_name, buf_comment, buf_designation

        if buf_pos is None:
            return

        name = clean(" ".join(buf_name))
        comment = clean(" ".join(buf_comment))
        designation = clean(" ".join(buf_designation))

        manufacturer = extract_manufacturer(name)

        rows.append([
            module_code,
            current_section,
            int(buf_pos),
            name,
            manufacturer,
            designation,
            int(buf_qty),
            comment
        ])

        # Сброс буферов
        buf_pos = None
        buf_qty = None
        buf_name = []
        buf_comment = []
        buf_designation = []

    # ==========================
    # ПРОХОД ПО ТАБЛИЦАМ
    # ==========================

    for table in doc.tables:
        for row in table.rows:
            cells = [clean(c.text) for c in row.cells]

            if not any(cells):
                continue

            row_text = " ".join(cells)

            # --------------------------
            # Определение разделов
            # --------------------------
            if "Стандартные изделия" in row_text:
                flush()
                current_section = "Стандартные"
                continue

            if "Прочие изделия" in row_text:
                flush()
                current_section = "Прочие"
                continue

            if any(x in row_text for x in ["Документация", "Детали"]):
                flush()
                current_section = None
                continue

            # --------------------------
            # Пропуск заголовков
            # --------------------------
            if "Формат" in row_text and "Поз." in row_text:
                continue

            if current_section not in ("Стандартные", "Прочие"):
                continue

            # --------------------------
            # Нормализация колонок
            # --------------------------
            # Ожидаемый формат:
            # Формат | Зона | Поз | Обозн | Наим | Кол | Прим
            while len(cells) < 7:
                cells.append("")

            fmt, zone, pos_c, desig_c, name_c, qty_c, comm_c = cells[:7]

            # --------------------------
            # Новая позиция?
            # --------------------------
            is_new = (
                re.fullmatch(r"\d{1,3}", pos_c)
                and
                re.fullmatch(r"\d{1,4}", qty_c)
            )

            if is_new:
                flush()
                buf_pos = pos_c
                buf_qty = qty_c

            # --------------------------
            # Накопление колонок
            # --------------------------
            if buf_pos is not None:
                if desig_c:
                    buf_designation.append(desig_c)
                if name_c:
                    buf_name.append(name_c)
                if comm_c:
                    buf_comment.append(comm_c)

    # Последняя позиция
    flush()
    return rows


# ==========================
# ВЫГРУЗКА В EXCEL (один файл)
# ==========================

def save_xlsx_many(specs, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "BOM"

    # Заголовок
    ws.append([
        "Module",
        "Section",
        "Pos",
        "Name",
        "Manufacturer",
        "PartNumber",
        "Qty",
        "Comment"
    ])

    total_rows = 0

    for i, (input_docx, module_code) in enumerate(specs, start=1):
        data = parse_spec(input_docx, module_code)

        # сортировка блока (как было)
        data.sort(key=lambda x: (0 if x[1] == "Стандартные" else 1, x[2]))

        for r in data:
            ws.append(r)
            total_rows += 1

        # Пустая строка после каждого файла (как ты попросил),
        # но не обязательно после последнего — можешь оставить/убрать по вкусу.
        ws.append([""] * 8)

    # Ширины колонок (как было)
    widths = [16, 14, 6, 60, 28, 28, 6, 60]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)
    return total_rows


# ==========================
# ТОЧКА ВХОДА
# ==========================

if __name__ == "__main__":
    n = save_xlsx_many(SPECS, OUTPUT_XLSX)
    print(f"Готово: {len(SPECS)} файлов, {n} строк BOM → {OUTPUT_XLSX}")
