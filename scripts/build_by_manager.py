# -*- coding: utf-8 -*-
"""
Раскладывает позиции трёх готовых заявок по менеджерам отдела снабжения.
Категории/менеджер для каждой позиции определены вручную по таблице
docs/Отдел снабжения 17.06.2026.xlsx и содержимому заявок (см. отчёт агента).
"""
import re
import os
import openpyxl
from openpyxl.styles import Font, Alignment

BASE = r"C:\Users\PC-1\Desktop\Разное\Claude code\final"
SRC_DIR = os.path.join(BASE, "reports", "Заявки готовые")
OUT_DIR = os.path.join(BASE, "reports", "Заявки по менеджерам")

MANAGERS = {
    "M1": "Менеджер 1",
    "M2": "Менеджер 2",
    "M4": "Менеджер 4",
    "M3": "Менеджер 3",
    "M5": "Менеджер 5",
}

COLS = ["№ п/п", "Раздел", "Наименование", "Марка", "Ед. изм.", "Кол-во",
        "Область применения", "Примечание", "Категория"]

# ---- Заявка №83 Жилой дом Шатура ----
ZAYAVKA_83 = {
    1: ("M1", "Молниезащита — Электрика и слабые токи"),
    2: ("M1", "Молниезащита — Электрика и слабые токи"),
    3: ("M1", "Электрика и слабые токи"),
    4: ("M1", "Электрика и слабые токи"),
    5: ("M1", "Электрика и слабые токи"),
    6: ("M1", "Электрика и слабые токи"),
    7: ("M1", "Электрика и слабые токи"),
    8: ("M1", "Электрика и слабые токи"),
    9: ("M4", "Навигация — спорно: возможно Менеджер 1 (Электрика и слабые токи)"),
    10: ("M2", "Пожарное оборудование — спорно: возможно Менеджер 4 (Навигация)"),
    11: ("M1", "Электрика и слабые токи"),
    12: ("M1", "Утеплитель"),
    13: ("M3", "Плитка"),
    14: ("M3", "Плитка"),
    15: ("M3", "Плитка"),
    16: ("M3", "Плитка"),
    17: ("M3", "Плитка"),
}

# ---- Заявка №53 Жилые дома Подольск Климовск ----
ZAYAVKA_53 = {
    1: ("M1", "Электрика и слабые токи"),
    2: ("M1", "Электрика и слабые токи"),
    3: ("M1", "Электрика и слабые токи"),
    4: ("M2", "Электроинструмент и оборудование"),
    5: ("M2", "Электроинструмент и оборудование"),
    6: ("M2", "Спецодежда"),
    7: ("M2", "Спецодежда"),
    8: ("M2", "Спецодежда"),
    9: ("M2", "Электроинструмент и оборудование"),
    10: ("M2", "Электроинструмент и оборудование"),
    11: ("M2", "Электроинструмент и оборудование"),
    12: ("M2", "Крепёж"),
    13: ("M2", "Крепёж"),
    14: ("M2", "Крепёж"),
    15: ("M2", "Крепёж"),
    16: ("M2", "Крепёж"),
    17: ("M1", "Электрика и слабые токи"),
    18: ("M1", "Электрика и слабые токи"),
    19: ("M1", "Электрика и слабые токи"),
    20: ("M1", "Электрика и слабые токи"),
    21: ("M2", "Электроинструмент и оборудование"),
    22: ("M2", "Электроинструмент и оборудование — спорно: возможно Менеджер 4 (Бытовые нужды на объекте)"),
    23: ("M2", "Крепёж"),
    24: ("M1", "Электрика и слабые токи"),
    25: ("M1", "Электрика и слабые токи"),
    26: ("M4", "Расходный инструмент"),
    27: ("M4", "Расходный инструмент"),
    28: ("M4", "Расходный инструмент — спорно: возможно Менеджер 2 (Электроинструмент и оборудование)"),
    29: ("M4", "Расходный инструмент — спорно: возможно Менеджер 2 (Электроинструмент и оборудование)"),
    30: ("M2", "не распознано, предположительно Электроинструмент и оборудование"),
    31: ("M2", "Электроинструмент и оборудование"),
    32: ("M2", "Электроинструмент и оборудование"),
    33: ("M2", "не распознано, предположительно Бытовки"),
    34: ("M2", "Электроинструмент и оборудование — спорно: возможно Менеджер 1 (Электрика и слабые токи)"),
}

# ---- Заявка №44 ДОУ Видное ----
ZAYAVKA_44 = {
    1: ("M4", "Расходный инструмент"),
    2: ("M4", "Расходный инструмент"),
    3: ("M3", "Сухие смеси"),
    4: ("M3", "Сухие смеси"),
    5: ("M2", "Отделочные материалы"),
    6: ("M1", "ЖБИ"),
    7: ("M1", "ЖБИ"),
    8: ("M1", "ЖБИ — спорно: возможно Менеджер 3 (Стеновые материалы, газобетон)"),
    9: ("M2", "Сантехника"),
    10: ("M2", "Сантехника"),
    11: ("M2", "Сантехника"),
    12: ("M2", "Спецодежда"),
    13: ("M4", "Расходный инструмент — спорно: возможно Менеджер 2 (Электроинструмент и оборудование)"),
    14: ("M2", "Крепёж"),
    15: ("M1", "Электрика и слабые токи"),
    16: ("M1", "Электрика и слабые токи"),
    17: ("M1", "Электрика и слабые токи"),
    18: ("M1", "Электрика и слабые токи"),
    19: ("M1", "Электрика и слабые токи"),
    20: ("M1", "Электрика и слабые токи"),
    21: ("M2", "Электроинструмент и оборудование"),
    22: ("M2", "Крепёж"),
    23: ("M3", "Сухие смеси"),
}

JOBS = [
    {
        "src": "Заявка №83 от 24.09.2026г. Жилой дом Шатура.xlsx",
        "folder": "Заявка 83 Жилой дом Шатура",
        "map": ZAYAVKA_83,
    },
    {
        "src": "Заявка №53 от 24.09.2026г. Жилые дома Подольск Климовск.xlsx",
        "folder": "Заявка 53 Жилые дома Подольск Климовск",
        "map": ZAYAVKA_53,
    },
    {
        "src": "Заявка №44 от 23.09.2026г. ДОУ Видное.xlsx",
        "folder": "Заявка 44 ДОУ Видное",
        "map": ZAYAVKA_44,
    },
]


def parse_header(text):
    m = re.match(r"Заявка №(\d+) от (\d{2}\.\d{2}\.\d{4}г?\.) (.+)$", text)
    if not m:
        raise ValueError(f"Не удалось разобрать шапку: {text!r}")
    num, date, obj = m.groups()
    return num, date, obj


def load_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Заявка"]
    header_text = ws.cell(row=1, column=1).value
    num, date, obj = parse_header(header_text)
    rows = []
    for r in range(4, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 9)]
        if vals[0] is None:
            continue
        rows.append(vals)
    return header_text, num, date, obj, rows


def write_manager_file(out_path, header_full, manager_full_name, num, obj, items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Позиции"

    n_cols = 11
    title = f"{header_full} — {manager_full_name}"
    ws.cell(row=1, column=1, value=title)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left")

    header_row = 3
    headers = ["Заявка №", "Объект", "№ п/п", "Раздел", "Наименование", "Марка",
               "Ед. изм.", "Кол-во", "Область применения", "Примечание", "Категория"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = Font(bold=True)

    r = header_row + 1
    for item in items:
        npp, razdel, naim, marka, ed, kolvo, obl, prim, kategoria = item
        ws.cell(row=r, column=1, value=int(num))
        ws.cell(row=r, column=2, value=obj)
        ws.cell(row=r, column=3, value=npp)
        ws.cell(row=r, column=4, value=razdel)
        ws.cell(row=r, column=5, value=naim)
        ws.cell(row=r, column=6, value=marka)
        ws.cell(row=r, column=7, value=ed)
        ws.cell(row=r, column=8, value=kolvo)
        ws.cell(row=r, column=9, value=obl)
        ws.cell(row=r, column=10, value=prim)
        ws.cell(row=r, column=11, value=kategoria)
        r += 1

    widths = [9, 20, 7, 16, 45, 16, 9, 8, 20, 20, 45]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    assert out_path.replace("\\", "/").startswith(
        os.path.join(BASE, "reports", "Заявки по менеджерам").replace("\\", "/")
    ), f"Путь вне разрешённой папки: {out_path}"

    wb.save(out_path)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    report = []
    for job in JOBS:
        src_path = os.path.join(SRC_DIR, job["src"])
        header_full, num, date, obj, rows = load_rows(src_path)
        mapping = job["map"]

        by_manager = {}
        for row in rows:
            npp = row[0]
            if npp not in mapping:
                raise ValueError(f"{job['src']}: нет назначения для позиции №{npp}")
            mgr_key, category = mapping[npp]
            by_manager.setdefault(mgr_key, []).append(tuple(row) + (category,))

        total_in_zayavka = len(rows)
        total_assigned = sum(len(v) for v in by_manager.values())

        folder_path = os.path.join(OUT_DIR, job["folder"])
        os.makedirs(folder_path, exist_ok=True)

        job_report = {
            "folder": job["folder"],
            "total_in_zayavka": total_in_zayavka,
            "total_assigned": total_assigned,
            "managers": [],
        }

        for mgr_key, items in by_manager.items():
            manager_full_name = MANAGERS[mgr_key]
            fname = manager_full_name + ".xlsx"
            out_path = os.path.join(folder_path, fname)
            write_manager_file(out_path, header_full, manager_full_name, num, obj, items)
            job_report["managers"].append((manager_full_name, len(items)))

        report.append(job_report)

    with open(os.path.join(BASE, "tmp", "report.txt"), "w", encoding="utf-8") as f:
        for jr in report:
            f.write(f"== {jr['folder']} ==\n")
            f.write(f"всего в заявке: {jr['total_in_zayavka']}, разложено: {jr['total_assigned']}\n")
            for name, cnt in jr["managers"]:
                f.write(f"  {name} — {cnt}\n")
            f.write("\n")

    print("OK")


if __name__ == "__main__":
    main()
