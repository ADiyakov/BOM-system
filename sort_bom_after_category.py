import re
from pathlib import Path
from openpyxl import load_workbook, Workbook


# ========== НАСТРОЙКИ ==========
INPUT_XLSX = "BOM_with_category.xlsx"
SHEET_NAME = "BOM"
OUTPUT_XLSX = "BOM_with_category_sorted.xlsx"

CAT_FASTENERS = "Крепежные изделия"
SECTION_STANDARD = "Стандартные"
# ===============================


# --------- утилиты ----------
def s(x) -> str:
    return "" if x is None else str(x).strip()

def norm_space(x: str) -> str:
    return re.sub(r"\s+", " ", s(x)).strip()

def lower_ru_lat(x: str) -> str:
    return norm_space(x).lower()

def natural_key(text: str):
    """Естественная сортировка: A2 < A10."""
    t = lower_ru_lat(text)
    parts = re.split(r"(\d+)", t)
    out = []
    for p in parts:
        if p.isdigit():
            out.append(int(p))
        else:
            out.append(p)
    return tuple(out)

def to_float(num: str) -> float:
    return float(num.replace(",", "."))
# ----------------------------


# --------- крепёж ----------
DIN_RE = re.compile(r"(?iu)\bDIN\s*([0-9]{2,6})\b")
ISO_RE = re.compile(r"(?iu)\bISO\s*([0-9]{2,6})\b")
GOST_RE = re.compile(r"(?iu)\bГОСТ(?:\s*Р)?\s*([0-9]{2,6})\b")

M_SIZE_RE = re.compile(
    r"(?iu)\b[МM]\s*(\d+(?:[.,]\d+)?)\s*(?:[x×\*]\s*(\d+(?:[.,]\d+)?))?"
)

def fastener_key(name: str):
    """
    Сортировка для крепежа:
    1) DIN/ISO/ГОСТ номер
    2) Алфавит
    3) Размер резьбы (M)
    4) Длина
    """
    n = norm_space(name)

    din = DIN_RE.search(n)
    iso = ISO_RE.search(n)
    gost = GOST_RE.search(n)

    std_rank = 9
    std_num = 10**9

    if din:
        std_rank = 1
        std_num = int(din.group(1))
    elif iso:
        std_rank = 2
        std_num = int(iso.group(1))
    elif gost:
        std_rank = 3
        std_num = int(gost.group(1))

    # размеры M
    thread = 10**9
    length = 10**9
    m = M_SIZE_RE.search(n)
    if m:
        try:
            thread = to_float(m.group(1))
        except Exception:
            pass
        if m.group(2):
            try:
                length = to_float(m.group(2))
            except Exception:
                pass

    return (std_rank, std_num, lower_ru_lat(n), thread, length, natural_key(n))
# ----------------------------


def row_sort_key(row_dict: dict):
    cat = s(row_dict.get("Category"))
    name = s(row_dict.get("Name"))
    section = s(row_dict.get("Section"))

    cat_key = lower_ru_lat(cat)

    # Только крепеж — специальная логика
    if cat == CAT_FASTENERS or section == SECTION_STANDARD:
        return (cat_key, 0, fastener_key(name))

    # Всё остальное — обычная естественная сортировка
    return (cat_key, 9, natural_key(name))


def main():
    in_path = Path(INPUT_XLSX)
    if not in_path.exists():
        raise FileNotFoundError(f"Не найден {INPUT_XLSX}")

    wb = load_workbook(in_path)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    headers = [s(c.value) for c in ws[1]]
    if "Category" not in headers or "Name" not in headers:
        raise RuntimeError(f"Ожидались колонки Category и Name. Есть: {headers}")

    rows = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, len(headers) + 1)]
        row_dict = {headers[i]: values[i] for i in range(len(headers))}
        rows.append((row_dict, values))

    rows_sorted = sorted(rows, key=lambda rv: row_sort_key(rv[0]))

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = ws.title

    out_ws.append(headers)
    for _, values in rows_sorted:
        out_ws.append(values)

    out_wb.save(OUTPUT_XLSX)
    print(f"OK: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
