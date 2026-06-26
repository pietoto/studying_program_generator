import argparse
import json
import os
from collections import defaultdict

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

# загрузка json-файла и проверка данных
def load_json(file_path):
    # проверяем существует ли файл
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    # читаем содержимое файла
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # проверяем что внутри действительно словарь и он не пустой
    if not isinstance(data, dict) or not data:
        raise ValueError("JSON-файл пустой или имеет неверную структуру")

    return data

# приводим текст к удобному виду
def normalize_text(value):
    # если значения нет, возвращаем пустую строку
    if value is None:
        return ""
    # переводим всё в строку
    text = str(value)
    # убираем переносы строк
    text = text.replace("\r", " ").replace("\n", " ")
    # убираем лишние пробелы
    while "  " in text:
        text = text.replace("  ", " ")

    return text.strip()

# получаем количество зачётных единиц
def get_credits(discipline):
    # проверяем оба возможных названия поля
    for key in ("credits", "total_credits"):
        try:
            value = discipline.get(key, 0)
            if value is None:
                continue
            return float(value)
        except (TypeError, ValueError):
            continue

    return 0.0

# подготавливаем семестры перед выводом
def prepare_semestrs(semestrs):
    prepared = {}

    for sem_key, disciplines in semestrs.items():
        filtered = []

        for disc in disciplines or []:
            # убираем факультативы
            if disc.get("is_facultative", False):
                continue

            # убираем дисциплины с нулевыми зачётными единицами
            credits = get_credits(disc)
            if credits <= 0:
                continue

            # нормализуем название
            name = normalize_text(disc.get("name", "")).lower()

            # если в названии есть русский как иностранный, оставляем базовый курс
            if "иностранный язык" in name and "русский язык как иностранный" in name:
                disc = disc.copy()
                disc["name"] = "Иностранный язык: Базовый курс"

            filtered.append(disc)

        # сортируем дисциплины по количеству ЗЕТ, потом по названию
        filtered.sort(
            key=lambda x: (
                -get_credits(x),
                normalize_text(x.get("name", "")).lower()
            )
        )

        prepared[str(sem_key)] = filtered

    return prepared


# группируем семестры по курсам
def group_by_course(semestrs):
    grouped = defaultdict(lambda: defaultdict(list))

    for sem_str, disciplines in semestrs.items():
        try:
            sem = int(sem_str)
        except (TypeError, ValueError):
            continue

        # определяем номер курса по номеру семестра
        course = (sem + 1) // 2

        for disc in disciplines:
            grouped[course][sem].append(disc)

    return dict(grouped)


# находит максимальное количество зачетных единиц
def find_max_credits(semestrs):
    max_credits = 0.0

    for disciplines in semestrs.values():
        for disc in disciplines or []:
            credits = get_credits(disc)
            if credits > max_credits:
                max_credits = credits

    return max_credits if max_credits > 0 else 1.0


# добавляет цвет ячейке
def add_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))

    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)

    shd.set(qn("w:fill"), fill)


# добавляет отступы в ячейке, чтобы текст не касался границ
def set_cell_margins(cell, top=20, start=20, bottom=20, end=20):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))

    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)

    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


# задаёт минимальную высоту строки таблицы
def set_row_height(row, height_pt):
    row.height = Pt(height_pt)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

# устанавливает поля страницы
def set_doc_margins(section):
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(0.4)
    section.right_margin = Cm(0.4)
    section.top_margin = Cm(0.4)
    section.bottom_margin = Cm(0.4)

# настраивает шрифт текста
def set_font(run, size=11, bold=False, italic=False, name="Times New Roman"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic

# добавляет обычный абзац
def add_text_paragraph(container, text, size=11, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, underline=False):
    p = container.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0

    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    run.font.underline = underline

    return p

# очищает ячейку
def clear_cell(cell):
    cell.text = ""

# заполняет ячейку текстом и оформляет её
def fill_cell_text(cell, text, size=10, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, color_fill=None):
    clear_cell(cell)

    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0

    # разделяем текст на строки
    lines = str(text).split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            p.add_run().add_break()
        run = p.add_run(line)
        set_font(run, size=size, bold=bold)

    # выравниваем текст по верхнему краю
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_margins(cell)

    if color_fill:
        add_shading(cell, color_fill)

# задаёт стиль таблицы
def add_table_border_style(table):
    # рамки у таблицы и выравнивание по центру
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

# подбирает правильную форму слова "год"
def plural_years(value):
    text = normalize_text(value)

    if not text:
        return "не указан"
    try:
        number = int(float(text))

        if number % 10 == 1 and number % 100 != 11:
            suffix = "год"
        elif number % 10 in (2, 3, 4) and number % 100 not in (12, 13, 14):
            suffix = "года"
        else:
            suffix = "лет"

        return f"{number} {suffix}"
    except ValueError:
        return text

# создает документ по данным учебного плана
def build_document(data, output_path):
    semestrs_raw = data.get("semestrs", {})
    if not semestrs_raw:
        raise ValueError("В JSON нет данных о семестрах")

    # подготавливаем данные для вывода
    semestrs = prepare_semestrs(semestrs_raw)
    grouped = group_by_course(semestrs)

    # определяем номера семестров
    all_semesters = []
    for s in semestrs.keys():
        if str(s).isdigit():
            all_semesters.append(int(s))
    all_semesters.sort()

    if not all_semesters:
        raise ValueError("Не удалось определить номера семестров")

    # считаем общие параметры
    total_semesters = len(all_semesters)
    max_course = max((sem + 1) // 2 for sem in all_semesters)
    max_credits = find_max_credits(semestrs)

    # ищем максимальное число строк
    max_rows = 0
    for disciplines in semestrs.values():
        if len(disciplines) > max_rows:
            max_rows = len(disciplines)

    if max_rows <= 0:
        max_rows = 1

    doc = Document()
    section = doc.sections[0]
    set_doc_margins(section)

    # базовый стиль документа
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal_style.font.size = Pt(10)

    # шапка документа
    add_text_paragraph(
        doc,
        "ВИЗУАЛИЗАЦИЯ УЧЕБНОГО ПЛАНА",
        size=12,
        bold=True,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    doc.add_paragraph()

    plan_info = data.get("plan_info", {})
    title = normalize_text(plan_info.get("title", "Не указано"))
    qualification = normalize_text(plan_info.get("qualification", "Не указана"))
    duration = plural_years(plan_info.get("duration", ""))
    year = normalize_text(plan_info.get("year", "Не указан"))

    add_text_paragraph(doc, f"Направление: {title or 'не указано'}", size=10, bold=True)
    add_text_paragraph(doc, f"Квалификация: {qualification or 'не указана'}", size=10, bold=True)
    add_text_paragraph(doc, f"Срок обучения: {duration}", size=10, bold=True)
    add_text_paragraph(doc, f"Учебный год: {year or 'не указан'}", size=10, bold=True)

    doc.add_paragraph()

    # создаём таблицу
    table = doc.add_table(rows=2 + max_rows, cols=total_semesters)
    add_table_border_style(table)

    usable_width_cm = 29.7 - 0.4 - 0.4
    col_width_cm = usable_width_cm / total_semesters

    for col in range(total_semesters):
        table.columns[col].width = Cm(col_width_cm)

    # заполняем заголовки курсов и семестров
    for col_index, sem in enumerate(all_semesters):
        course = (sem + 1) // 2

        # первый ряд - курс
        top_cell = table.cell(0, col_index)
        if sem % 2 == 1 and col_index + 1 < total_semesters:
            top_cell = top_cell.merge(table.cell(0, col_index + 1))

        fill_cell_text(
            top_cell,
            f"КУРС {course}",
            size=8,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            color_fill="D9D9D9",
        )

        # второй ряд - семестр
        bottom_cell = table.cell(1, col_index)
        fill_cell_text(
            bottom_cell,
            f"Сем. {sem}",
            size=8,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            color_fill="DDEBF7",
        )

    # заполняем таблицу дисциплинами
    for row_index in range(max_rows):
        row_credits = 0.0

        for col_index, sem in enumerate(all_semesters):
            disciplines = semestrs.get(str(sem), [])
            cell = table.cell(row_index + 2, col_index)

            if row_index < len(disciplines):
                disc = disciplines[row_index] or {}
                name = normalize_text(disc.get("name", ""))
                credits = get_credits(disc)

                if credits > row_credits:
                    row_credits = credits

                text = f"{name}\n ({credits:g} зет)"
                fill_cell_text(cell, text, size=8.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
            else:
                fill_cell_text(cell, "", size=8.5, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    # высота заголовков
    set_row_height(table.rows[0], 12)
    set_row_height(table.rows[1], 10)

    # сохраняем готовый документ
    doc.save(output_path)
    print(f"Готово: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Визуализация учебного плана в DOCX")
    parser.add_argument("-o", "--output", default="visualized_plan.docx", help="Путь к DOCX")
    args = parser.parse_args()

    # работаем с json-файлом
    input_file = "result_parsinig_plan.json"

    try:
        data = load_json(input_file)
        build_document(data, args.output)
    except FileNotFoundError as exc:
        print(f"Ошибка: {exc}")
    except ValueError as exc:
        print(f"Ошибка данных: {exc}")
    except Exception as exc:
        print(f"Неожиданная ошибка: {exc}")


if __name__ == "__main__":
    main()
