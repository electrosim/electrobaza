# -*- coding: utf-8 -*-
"""Раздаточный лист А4: правовой анализ ПУЭ (МДК 04.02 + ОП.12)."""

import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\5-pue-i-gost\docs"
OUT = os.path.join(BASE, "Задание — Правовой анализ ПУЭ (на печать).docx")
FONT = "Times New Roman"


def style_base(doc):
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.5)
    s.top_margin = Cm(1.2)
    s.bottom_margin = Cm(1.2)

    normal = doc.styles["Normal"]
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    normal.font.name = FONT
    normal.font.size = Pt(9.5)
    normal.paragraph_format.space_after = Pt(1)
    normal.paragraph_format.line_spacing = 1.0


def p(doc, text="", size=9.5, bold=False, italic=False, align=None, space_after=1, space_before=0, color=None):
    par = doc.add_paragraph()
    if align is not None:
        par.alignment = align
    par.paragraph_format.space_after = Pt(space_after)
    par.paragraph_format.space_before = Pt(space_before)
    par.paragraph_format.line_spacing = 1.0
    run = par.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    return par


def cell_text(cell, text, size=8.5, bold=False, italic=False, align=None):
    cell.paragraphs[0].text = ""
    par = cell.paragraphs[0]
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.line_spacing = 1.0
    if align is not None:
        par.alignment = align
    run = par.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def set_borders(table, sz=6, color="808080"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def no_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def build():
    doc = Document()
    style_base(doc)

    p(doc, "Правовой анализ Правил устройства электроустановок", size=13, bold=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=1)
    p(doc, "Практическая работа · МДК 04.02 + ОП.12 · 1ЭО-24", size=9, italic=True,
      align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)

    p(doc, "ФИО ______________________________________     Группа ____________     Дата ____________",
      size=9, space_after=6)

    p(doc, "1. Коротко о главном", size=10, bold=True, space_before=2, space_after=2)
    p(doc, "• ПУЭ — ведомственный нормативный акт, а не закон. Действует 6-е издание ПУЭ "
           "(с изменениями на 20.06.2003); 7-е издание вводилось отдельными главами, но не прошло "
           "регистрацию в Минюсте.")
    p(doc, "• Обязательность ПУЭ возникает через технические регламенты (ТР ТС 004/2011, 020/2011), "
           "проектную документацию, договор подряда, трудовые обязанности и надзор.")
    p(doc, "• За нарушение отвечают по-разному: административная (КоАП), уголовная (УК), "
           "гражданско-правовая (ГК), дисциплинарная (ТК).")
    p(doc, "• ГОСТ применяется добровольно, но становится обязательным, если на него ссылается "
           "техрегламент или договор.", space_after=5)

    p(doc, "2. Нормы для работы", size=10, bold=True, space_before=2, space_after=2)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_borders(table)
    widths = [Cm(9.0), Cm(9.0)]
    hdr = table.rows[0].cells
    cell_text(hdr[0], "ПУЭ (техническая норма)", 8.5, bold=True)
    cell_text(hdr[1], "Право (юридическая норма)", 8.5, bold=True)
    rows = [
        ("п. 2.1.21, гл. 2.1 (6-е изд.): соединение, ответвление и оконцевание жил проводов "
         "и кабелей должны производиться при помощи опрессовки, сварки, пайки или сжимов "
         "(винтовых, болтовых и т. п.).",
         "КоАП РФ, ст. 9.11 — нарушение правил эксплуатации электроустановок: штраф, "
         "приостановление деятельности."),
        ("гл. 1.7: заземление и защитные меры электробезопасности; металлические части, "
         "которые могут оказаться под напряжением, должны быть заземлены или занулены "
         "(изложение).",
         "УК РФ, ст. 143 — нарушение требований охраны труда, повлёкшее тяжкий вред "
         "или смерть."),
        ("гл. 3.1: защита электрических сетей до 1 кВ от сверхтоков автоматическими "
         "выключателями и предохранителями (изложение).",
         "ГК РФ, ст. 1064, 1068 — возмещение вреда; за вред, причинённый работником, "
         "отвечает работодатель."),
        ("гл. 2.1 действует в 6-м издании — проверяйте, какая редакция применяется "
         "к конкретному вопросу.",
         "ТК РФ, ст. 192 — дисциплинарные взыскания: замечание, выговор, увольнение."),
    ]
    for left, right in rows:
        cells = table.add_row().cells
        cell_text(cells[0], left, 8)
        cell_text(cells[1], right, 8)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]

    p(doc, "3. Ситуации", size=10, bold=True, space_before=6, space_after=2)
    p(doc, "Ситуация 1. В распределительной коробке жилы соединены скруткой и заизолированы "
           "изолентой. Через полгода место соединения нагрелось, произошло возгорание, "
           "повреждено имущество собственника квартиры. Экспертиза установила: причина — "
           "переходное сопротивление в скрутке.")
    p(doc, "Ситуация 2. При монтаже щита электромонтажник не установил УЗО, предусмотренное "
           "проектом и договором подряда. При прикосновении к повреждённому проводу жилец "
           "получил травму. Проект предусматривал УЗО с уставкой 30 мА.", space_after=4)

    p(doc, "4. Задание", size=10, bold=True, space_before=2, space_after=2)
    p(doc, "Заполните таблицу для каждой ситуации. Ссылки — на конкретный пункт ПУЭ и статью закона.",
      size=9, space_after=3)

    for label in ("Ситуация 1", "Ситуация 2"):
        p(doc, label, size=9.5, bold=True, space_before=3, space_after=2)
        t = doc.add_table(rows=0, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_borders(t, sz=4, color="A6A6A6")
        questions = [
            "1. Какой пункт (глава) ПУЭ нарушен и в чём именно?",
            "2. Почему эта норма обязательна для работника?",
            "3. Вид правонарушения и статья (КоАП, УК, ГК, ТК).",
            "4. Кто отвечает: работник, работодатель или организация?",
            "5. Вывод: что следовало сделать иначе?",
        ]
        for q in questions:
            row = t.add_row()
            row.cells[0].width = Cm(7.4)
            row.cells[1].width = Cm(10.6)
            cell_text(row.cells[0], q, 8.5)
            cell_text(row.cells[1], "", 8.5)
            row.height = Cm(0.92)

    p(doc, "5. Критерии", size=10, bold=True, space_before=6, space_after=2)
    p(doc, "• Точная ссылка на пункт ПУЭ и статью закона — 4 балла за каждую ситуацию.")
    p(doc, "• Различение видов ответственности и субъекта ответственности — 3 балла.")
    p(doc, "• Аргумент, почему норма обязательна (регламент, проект, договор, обязанности) — 2 балла.")
    p(doc, "• Вывод и предложенное исправление — 1 балл.", space_after=0)

    p(doc, "Работа выполняется без доступа к интернету: все необходимые нормы приведены на этом листе.", size=9, italic=True, space_before=4, space_after=0)

    os.makedirs(BASE, exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
