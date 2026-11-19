# -*- coding: utf-8 -*-
"""Лекция «Работа с физическими величинами» (1ЭО-24, ОП.08)."""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-24\op-08\5-fizicheskie-velichiny"
OUT = os.path.join(BASE, "docs", "Лекция — Работа с физическими величинами.docx")
FONT = "Times New Roman"


def _rfonts(rpr):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def style_font(style, size=None, bold=None, color=None):
    rpr = style.element.get_or_add_rPr()
    _rfonts(rpr)
    style.font.name = FONT
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
        col = rpr.find(qn("w:color"))
        if col is not None:
            for a in ("w:themeColor", "w:themeShade", "w:themeTint"):
                if col.get(qn(a)) is not None:
                    del col.attrib[qn(a)]


def rich(p, text, size=None, bold=None, italic=None):
    import re
    for tok in re.split(r"(\*\*[^*]+\*\*|\^[^\^]+\^|\*[^*]+\*)", text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2]); _rfonts(r._element.get_or_add_rPr()); r.bold = True
        elif tok.startswith("^") and tok.endswith("^") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.font.superscript = True
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.italic = True
        else:
            r = p.add_run(tok); _rfonts(r._element.get_or_add_rPr())
        r.font.name = FONT
        if size is not None:
            r.font.size = Pt(size)
        if bold is not None:
            r.bold = bold
        if italic is not None:
            r.italic = italic


def para(doc, text, indent=0.8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.07
    rich(p, text)


def heading(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    rich(p, text)


def shade(cell, color="E8F0EC"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def table(doc, caption, header, rows, widths):
    c = doc.add_paragraph()
    c.paragraph_format.space_before = Pt(4)
    c.paragraph_format.space_after = Pt(2)
    c.paragraph_format.keep_with_next = True
    rich(c, caption, size=9, italic=True)

    t = doc.add_table(rows=1, cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]
        cell.width = widths[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        rich(p, h, size=9, bold=True)
        shade(cell)

    for row in rows:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].width = widths[i]
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
            rich(p, value, size=9)

    after = doc.add_paragraph()
    after.paragraph_format.space_after = Pt(0)
    after.paragraph_format.line_spacing = 1.0
    rich(after, "", size=6)


def build():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2)

    normal = doc.styles["Normal"]
    style_font(normal, size=11)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.07

    h1 = doc.styles["Heading 1"]
    style_font(h1, size=11.5, bold=True, color=RGBColor(0, 0, 0))
    h1.paragraph_format.space_before = Pt(5)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Работа с физическими величинами", size=14, bold=True)

    para(doc, "Любая инженерная задача начинается с физических величин и их единиц. "
              "При работе с таблицами, базами данных и расчётными программами величина проходит "
              "несколько превращений: ввод, хранение, обработка и вывод результата. Ошибка "
              "в единицах на любом из этих этапов даёт неверный результат, причём заметить её "
              "труднее, чем арифметическую погрешность.")

    heading(doc, "Понятие физической величины")
    para(doc, "**Физическая величина** — свойство объекта или явления, которое можно измерить "
              "и выразить числом с указанием единицы. Величина записывается как произведение "
              "числового значения на единицу: например, сила тока 30 мА означает, что измеренная "
              "величина составляет тридцать тысячных ампера.")
    para(doc, "Числовое значение без единицы смысла не имеет: запись «ток равен 5» не позволяет "
              "понять, идёт речь об амперах или миллиамперах. Поэтому единица — обязательная часть "
              "значения, а не пояснение к нему.")
    para(doc, "Величины делят на основные и производные. В Международной системе единиц (СИ) "
              "семь основных величин: длина, масса, время, сила электрического тока, "
              "термодинамическая температура, количество вещества и сила света. Производные "
              "величины выражаются через основные: напряжение — через работу и заряд, "
              "сопротивление — через напряжение и силу тока.")

    heading(doc, "Единицы, применяемые в электротехнике")
    para(doc, "В профессиональных расчётах используют как единицы СИ, так и допущенные вместе "
              "с ними внесистемные единицы: минуту, час, градус, литр, а также киловатт-час "
              "для учёта электроэнергии. Обозначения единиц пишут латинскими буквами; "
              "в формулах их не склоняют и не заменяют словами.")

    table(
        doc,
        "Таблица 1 — Основные величины и их единицы",
        ["Величина", "Обозначение", "Единица", "Обозначение"],
        [
            ["Сила электрического тока", "I", "ампер", "А"],
            ["Электрическое напряжение", "U", "вольт", "В"],
            ["Электрическое сопротивление", "R", "ом", "Ом"],
            ["Электрическая мощность", "P", "ватт", "Вт"],
            ["Энергия", "W", "джоуль; киловатт-час", "Дж; кВт·ч"],
            ["Частота", "f", "герц", "Гц"],
            ["Электрическая ёмкость", "C", "фарад", "Ф"],
            ["Магнитная индукция", "B", "тесла", "Тл"],
        ],
        [Cm(6.2), Cm(3.0), Cm(4.6), Cm(3.4)],
    )

    heading(doc, "Кратные и дольные приставки")
    para(doc, "Для удобства записи используют десятичные приставки: кило (к) — тысяча, мега (М) — "
              "миллион, гига (Г) — миллиард; милли (м) — тысячная, микро (мк) — миллионная, "
              "нано (н) — миллиардная доли. Приставка входит в единицу: 1 мА — это 0,001 А, "
              "1 кОм — 1000 Ом.")
    para(doc, "Правило работы с приставками простое: **перед вычислениями величину приводят "
              "к основной единице**, а результат при необходимости снова переводят в удобную "
              "приставку. Приставка пишется слитно с обозначением единицы, без точки и пробела: "
              "мА, кВ, мкФ.")

    heading(doc, "Правила записи значений")
    para(doc, "**Пробел** между числом и обозначением единицы обязателен: 220 В, 30 мА, 50 Гц. "
              "Исключение — знак градуса, который ставится без пробела: 30°. **Десятичный "
              "разделитель** в русскоязычных документах — запятая: 0,5 А, 2,5 кОм.")
    para(doc, "**Диапазон** записывают с тире и повторяют единицу только при разных единицах: "
              "10–20 мА, но 5 В и 10 В. **Предельные значения** оформляют словами: «не более "
              "30 мА», «не менее 0,5 МОм».")
    para(doc, "В таблицах и программах единицу указывают в заголовке столбца или в отдельной "
              "графе. Это избавляет от повторения обозначения в каждой ячейке и исключает "
              "случайную смену единицы внутри столбца.")

    heading(doc, "Величины в таблицах и расчётных программах")
    para(doc, "Компьютерная обработка требует, чтобы в ячейке хранилось **число**, а не текст. "
              "Запись «30 мА» в ячейке считается текстом: с ней нельзя выполнить арифметическое "
              "действие. Поэтому в ячейку вводят числовое значение, а единицу выносят в заголовок "
              "или в отдельный столбец.")
    para(doc, "Перед обработкой все значения приводят к одной единице. Если часть данных записана "
              "в амперах, а часть — в миллиамперах, программа сложит числа без учёта приставок "
              "и выдаст результат, ошибочный в тысячу раз.")
    para(doc, "Отдельно различают **отображаемое значение** и **хранимое**. Формат ячейки может "
              "показывать 30 мкА, тогда как внутри хранится 0,00003 А. При копировании и экспорте "
              "данных важно понимать, какое из значений переносится.")
    para(doc, "Точность результата определяется точностью исходных данных. Промежуточные "
              "вычисления выполняют с запасом знаков, а в ответе оставляют столько значащих цифр, "
              "сколько позволяет точность измерений, и обязательно указывают единицу.")

    heading(doc, "Проверка размерности")
    para(doc, "Размерность — выражение величины через основные единицы. Проверка размерности "
              "служит простым способом обнаружить ошибку: если единицы в левой и правой частях "
              "формулы не совпадают, формула или подстановка неверны. Например, произведение "
              "напряжения на силу тока даёт ватты, а деление напряжения на сопротивление — амперы.")

    heading(doc, "Типичные ошибки")
    para(doc, "Чаще всего путают амперы и миллиамперы, ватты и киловатты, джоули и киловатт-часы. "
              "Опасны похожие обозначения: «МПа» и «мПа» различаются в миллиард раз, а «мм» и «м» — "
              "в тысячу. Ещё одна распространённая ошибка — округление до выполнения расчёта: "
              "округление допустимо только после получения результата.")

    heading(doc, "Краткий вывод")
    para(doc, "Физическая величина всегда состоит из числового значения и единицы. Работа "
              "с величинами в информационных системах требует единого подхода: единицы указывают "
              "в заголовках, значения приводят к одной единице до расчёта, в ячейках хранят числа, "
              "а результат проверяют по размерности. Эти правила не зависят от программы "
              "и одинаково важны для таблицы, базы данных и инженерного пакета.")

    heading(doc, "Вопросы для самопроверки")
    for i, q in enumerate([
        "Из каких частей состоит запись физической величины?",
        "Почему числовое значение без единицы не имеет смысла?",
        "Чем отличаются основные величины от производных?",
        "Что нужно сделать с величинами до вычислений?",
        "Почему запись «30 мА» в ячейке таблицы считается текстом?",
        "Как проверить правильность формулы по размерности?",
    ], 1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, f"{i}.\t{q}")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
