# -*- coding: utf-8 -*-
"""Практическая работа «Опрессовка гильзами и изоляция соединений» — 8 страниц по вариантам."""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\4-opressovka-gilzami"
OUT = os.path.join(BASE, "docs", "Практическая работа — Опрессовка гильзами и изоляция соединений.docx")
FONT = "Times New Roman"

TABLE_ROWS = [
    ("2,5", "ГМ-2,5", "—", "20", "1"),
    ("4", "ГМ-4", "—", "22", "1"),
    ("6", "ГМ-6", "—", "25", "1"),
    ("10", "ГМ-10", "ГА-10", "28", "1–2"),
    ("16", "ГМ-16", "ГА-16", "32", "2"),
    ("25", "ГМ-25", "ГА-25", "40", "2"),
    ("35", "ГМ-35", "ГА-35", "45", "2"),
]

VARIANTS = [
    {
        "num": 1,
        "cond": "Медная жила 2,5 мм², соединение в распределительной коробке в квартире.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Определите длину зачистки и число обжимов.",
            "Выберите способ изоляции для сухого помещения.",
            "Найдите ошибку: «обжать гильзу у края, заизолировать одним слоем изоленты».",
        ],
    },
    {
        "num": 2,
        "cond": "Медная жила 4 мм², соединение в подрозетнике.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Определите длину зачистки и число обжимов.",
            "Выберите изоляцию, удобную в ограниченном объёме подрозетника.",
            "Найдите ошибку: «зачистить жилы на 10 мм и ввести в гильзу неплотно».",
        ],
    },
    {
        "num": 3,
        "cond": "Медная жила 6 мм², соединение в коробке на кухне.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Определите длину зачистки и число обжимов.",
            "Выберите способ изоляции и объясните, почему он подходит.",
            "Найдите ошибку: «обжать гильзу, не зачистив жилы до металлического блеска».",
        ],
    },
    {
        "num": 4,
        "cond": "Медная жила 10 мм², соединение в распределительном щите.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Определите длину зачистки и число обжимов.",
            "Выберите изоляцию для щита.",
            "Найдите ошибку: «сделать один обжим и не проверять соединение протяжкой».",
        ],
    },
    {
        "num": 5,
        "cond": "Медная жила 16 мм², соединение в щите.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Сколько обжимов нужно и в каком порядке их выполнять?",
            "Выберите изоляцию и объясните выбор.",
            "Найдите ошибку: «жилу 16 мм² обжать одним обжимом посередине гильзы».",
        ],
    },
    {
        "num": 6,
        "cond": "Алюминиевая жила 10 мм², соединение на улице.",
        "tasks": [
            "Подберите гильзу и объясните, почему медная не подходит.",
            "Определите длину зачистки и число обжимов.",
            "Выберите изоляцию для улицы и обоснуйте.",
            "Найдите ошибку: «заизолировать одним слоем изоленты без герметизации».",
        ],
    },
    {
        "num": 7,
        "cond": "Алюминиевая жила 16 мм², соединение в сыром помещении.",
        "tasks": [
            "Подберите гильзу и обоснуйте выбор.",
            "Определите длину зачистки и число обжимов.",
            "Выберите изоляцию для сырого помещения.",
            "Найдите ошибку: «применить медную гильзу ГМ-16 для алюминиевой жилы».",
        ],
    },
    {
        "num": 8,
        "cond": "Нужно соединить медную жилу 25 мм² с алюминиевой 25 мм².",
        "tasks": [
            "Можно ли выполнить соединение в одной гильзе ГМ-25? Ответ обоснуйте.",
            "Что применяют для медно-алюминиевого соединения?",
            "Как проверить качество соединения после опрессовки?",
            "Найдите ошибку: «взять гильзу ГМ-25, ввести обе жилы и обжать».",
        ],
    },
]

QUESTIONS = [
    "Почему нельзя обжимать медную и алюминиевую жилы в одной гильзе?",
    "Чем гильза ГМ отличается от ГА и ГМЛ?",
    "От чего зависит длина зачистки жилы?",
    "Почему обжим выполняют по метке или от середины к краям?",
    "По каким признакам при осмотре бракуют соединение?",
    "Какую изоляцию выбирают для сырого помещения и почему?",
]


def _rfonts(rpr):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def rich(p, text, size=None, bold=None, italic=None):
    r = p.add_run(text)
    _rfonts(r._element.get_or_add_rPr())
    r.font.name = FONT
    if size is not None:
        r.font.size = Pt(size)
    if bold is not None:
        r.bold = bold
    if italic is not None:
        r.italic = italic


def para(doc, text, size=9.5, bold=False, after=1, align=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0
    rich(p, text, size=size, bold=bold)
    return p


def blank_line(doc, height=26):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = Pt(height)
    pPr = p._p.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "808080")
    bdr.append(bottom)
    pPr.insert_element_before(
        bdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
        "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
        "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
        "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
        "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
        "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
    )


def shade(cell, color="E8F0EC"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def sleeves_table(doc):
    header = ["Сечение, мм²", "Гильза медная", "Гильза алюминиевая", "Зачистка, мм", "Обжимов"]
    widths = [Cm(3.4), Cm(3.4), Cm(3.9), Cm(3.4), Cm(2.9)]
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
        rich(p, h, size=8.5, bold=True)
        shade(cell)
    for row in TABLE_ROWS:
        cells = t.add_row().cells
        for i, value in enumerate(row):
            cells[i].width = widths[i]
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            rich(p, value, size=8.5)


def build():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.top_margin = Cm(1.4)
    s.bottom_margin = Cm(1.2)
    s.left_margin = s.right_margin = Cm(1.7)

    normal = doc.styles["Normal"]
    _rfonts(normal.element.get_or_add_rPr())
    normal.font.size = Pt(9.5)
    normal.paragraph_format.space_after = Pt(1)
    normal.paragraph_format.line_spacing = 1.0

    for idx, v in enumerate(VARIANTS):
        if idx:
            doc.add_page_break()

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        rich(p, "Практическая работа. Опрессовка гильзами и изоляция соединений", size=11.5, bold=True)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        rich(p, f"ВАРИАНТ {v['num']}", size=12, bold=True)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        rich(p, "ФИО ______________________________  Группа ____________", size=9)

        para(doc, "Краткие сведения", size=9.5, bold=True)
        para(doc, "Гильза — трубка из меди или алюминия, в которую вводят жилы и обжимают прессом. "
                  "Медные гильзы ГМ применяют для медных жил, лужёные ГМЛ — для защиты от коррозии, "
                  "алюминиевые ГА — для алюминиевых жил. Медь с алюминием в одной гильзе не соединяют: "
                  "на границе металлов возникает гальваническая пара, соединение греется и разрушается.",
             size=8.5)
        para(doc, "Порядок: подобрать гильзу → зачистить жилы на длину гильзы плюс 2–3 мм → ввести "
                  "до упора, стык встык → обжать по метке или от середины к краям → осмотреть и "
                  "проверить протяжкой → заизолировать.", size=8.5)
        para(doc, "Годное соединение: нет заусенцев и трещин, гильза не раздута, проволоки не выходят "
                  "за торец, жила не проворачивается и не вытягивается. Изоляция: термоусадка, "
                  "изолента внахлёст в два слоя, изолирующие колпачки; в сырых помещениях — "
                  "с герметизацией.", size=8.5, after=3)

        para(doc, "Таблица гильз (учебные значения)", size=9.5, bold=True, after=1)
        sleeves_table(doc)

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        rich(p, "Задание", size=9.5, bold=True)

        para(doc, f"Условие: {v['cond']}", size=9, after=2)

        for i, task in enumerate(v["tasks"], 1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            p.paragraph_format.space_after = Pt(0)
            rich(p, f"{i}. {task}", size=9)
            blank_line(doc)

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(1)
        rich(p, "Контрольные вопросы", size=9.5, bold=True)

        for i, q in enumerate(QUESTIONS, 1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.5)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            p.paragraph_format.space_after = Pt(0)
            rich(p, f"{i}. {q}", size=9)
            blank_line(doc, height=22)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
