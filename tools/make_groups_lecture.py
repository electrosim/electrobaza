# -*- coding: utf-8 -*-
"""Лекция «Квалификационные группы по технике безопасности. Организация рабочего места»."""

import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\6-kvalifikacionnye-gruppy"
OUT = os.path.join(BASE, "docs", "Лекция — Квалификационные группы и организация рабочего места.docx")
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
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY
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
    p.paragraph_format.space_after = Pt(3)
    rich(p, "Квалификационные группы по технике безопасности.\n"
            "Организация рабочего места", size=13.5, bold=True)

    para(doc, "Безопасность электромонтажных работ определяется двумя условиями: допуском "
              "к работе подготовленного человека и правильно организованным рабочим местом. "
              "Первое условие закрепляется квалификационной группой по электробезопасности, "
              "второе — требованиями к подготовке места, инструменту, ограждениям и надзору.")

    heading(doc, "Квалификационные группы по электробезопасности")
    para(doc, "Квалификационная группа — это уровень знаний и умений работника в части "
              "электробезопасности, подтверждённый проверкой знаний. Группа определяет, какие "
              "работы работник может выполнять самостоятельно, а какие — только под руководством "
              "или в составе бригады. Всего установлено пять групп, обозначаемых римскими "
              "цифрами I–V.")
    para(doc, "**Группа I** присваивается неэлектротехническому персоналу, работа которого "
              "связана с возможной опасностью поражения электрическим током. Специального "
              "обучения не требуется: достаточно инструктажа с проверкой знаний в форме устного "
              "опроса и записи в журнале. Группа подтверждается ежегодно.")
    para(doc, "**Группы II–V** присваиваются электротехническому персоналу после обучения "
              "и проверки знаний. Чем выше группа, тем шире круг работ и ответственности: "
              "от обслуживания установок до 1000 В до руководства работами в установках "
              "выше 1000 В и права единоличного осмотра.")

    table(
        doc,
        "Таблица 1 — Квалификационные группы по электробезопасности",
        ["Группа", "Кому присваивается", "Что даёт"],
        [
            ["I", "Неэлектротехническому персоналу, связанному с работой вблизи "
                  "электроустановок", "Выполнение работ, не требующих обслуживания "
                  "электроустановок; инструктаж и запись в журнале"],
            ["II", "Начинающим электромонтажникам, персоналу, обслуживающему простые "
                   "установки", "Работы в электроустановках до 1000 В под надзором "
                   "или в составе бригады"],
            ["III", "Электромонтажникам, допущенным к самостоятельной работе",
             "Самостоятельное обслуживание установок до 1000 В, допуск и надзор "
             "за членами бригады"],
            ["IV", "Специалистам, руководящим работами в установках до 1000 В и "
                   "обслуживающим установки выше 1000 В",
             "Руководство бригадой, работы в установках выше 1000 В, проверка знаний "
             "у работников"],
            ["V", "Ответственным за электрохозяйство и особо сложные работы",
             "Руководство работами в установках выше 1000 В, единоличный осмотр, "
             "выдача нарядов и распоряжений"],
        ],
        [Cm(1.9), Cm(7.0), Cm(8.1)],
    )

    heading(doc, "Порядок присвоения и проверки знаний")
    para(doc, "К проверке знаний допускаются работники не моложе восемнадцати лет, прошедшие "
              "обучение по программе и медицинский осмотр. Проверка знаний бывает первичной, "
              "периодической и внеочередной. Периодическая проверка для электротехнического "
              "персонала проводится не реже одного раза в год, внеочередная — при переводе "
              "на другую работу, после нарушения правил или по требованию ответственного "
              "за электрохозяйство.")
    para(doc, "Результат проверки оформляется протоколом, а работнику выдаётся удостоверение "
              "с указанием группы и срока действия. Отсутствие удостоверения или просроченная "
              "проверка знаний означают, что работник к самостоятельной работе не допускается.")

    heading(doc, "Организация рабочего места")
    para(doc, "Рабочее место должно соответствовать характеру работы и требованиям безопасности. "
              "До начала работ проверяют: обесточены ли токоведущие части, установлены ли "
              "ограждения, вывешены ли плакаты безопасности, достаточно ли освещение, свободны ли "
              "проходы. Инструмент должен быть исправным и соответствовать виду работ.")
    para(doc, "На рабочем месте в действующей электроустановке указывают границы рабочей зоны "
              "и места, где запрещено находиться. При необходимости устанавливают переносные "
              "ограждения и заземления, а на приводах коммутационных аппаратов вывешивают "
              "запрещающие плакаты.")
    para(doc, "Работы в действующих электроустановках выполняют по наряду-допуску или "
              "распоряжению, с оформлением допуска и надзором. Члены бригады находятся "
              "на своих местах и не покидают рабочую зону без разрешения производителя работ.")
    para(doc, "Материалы и инструмент располагают так, чтобы они не загромождали проходы "
              "и не создавали опасность падения. После окончания работ рабочее место приводят "
              "в порядок, убирают инструмент и отходы, снимают временные ограждения "
              "и заземления.")

    heading(doc, "Плакаты и знаки безопасности")
    para(doc, "Плакаты безопасности делят на четыре группы. **Запрещающие** — «Не включать! "
              "Работают люди», «Не открывать! Работают люди» — вывешивают на приводах "
              "и ограждениях, чтобы исключить ошибочную подачу напряжения. **Предупреждающие** — "
              "«Осторожно! Электрическое напряжение» — обозначают опасность вблизи токоведущих "
              "частей.")
    para(doc, "**Предписывающие** плакаты («Работать здесь») указывают место, где разрешено "
              "выполнять работу, **указательные** («Заземлено») показывают, что участок заземлён. "
              "Плакаты вывешивают до начала работ и снимают только после их окончания, по указанию "
              "производителя работ.")

    heading(doc, "Средства защиты на рабочем месте")
    para(doc, "Средства защиты делят на основные и дополнительные. **Основные** выдерживают "
              "рабочее напряжение длительное время: изолирующие штанги и клещи, указатели "
              "напряжения, диэлектрические перчатки. **Дополнительные** сами по себе не защищают "
              "от напряжения, но усиливают действие основных: диэлектрические боты, галоши, "
              "коврики, изолирующие подставки.")
    para(doc, "Перед применением проверяют целостность средства защиты и срок его испытания. "
              "Просроченное или повреждённое средство защиты к работе не допускается.")

    heading(doc, "Документация и ответственность")
    para(doc, "Основные документы на рабочем месте: удостоверение о проверке знаний, "
              "наряд-допуск или распоряжение, журнал инструктажей и журнал учёта работ "
              "по нарядам. Ответственность распределена между выдающим наряд, допускающим, "
              "производителем работ и наблюдающим; каждый отвечает за свою часть допуска.")

    heading(doc, "Краткий вывод")
    para(doc, "Квалификационная группа подтверждает право работника выполнять определённый "
              "круг работ, а организация рабочего места обеспечивает безопасность этих работ. "
              "Допуск без группы, просроченное удостоверение или неубранное рабочее место "
              "сводят на нет любые знания: безопасность начинается с порядка и допуска.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
