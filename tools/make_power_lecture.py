# -*- coding: utf-8 -*-
"""Лекция «Органы власти Российской Федерации» (1ЭО-24, ОП.12)."""

import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-24\op-12\8-organy-vlasti"
OUT = os.path.join(BASE, "docs", "Лекция — Органы власти Российской Федерации.docx")
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
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Органы власти Российской Федерации", size=14, bold=True)

    para(doc, "Государственная власть в Российской Федерации строится на основе разделения "
              "на законодательную, исполнительную и судебную (статья 10 Конституции РФ). "
              "Органы каждой ветви самостоятельны, но действуют в единой системе: их решения "
              "опираются на Конституцию и законы, а споры между ними разрешаются правовыми "
              "средствами.")

    heading(doc, "Президент Российской Федерации")
    para(doc, "**Президент** — глава государства, гарант Конституции, прав и свобод человека "
              "и гражданина. Он определяет основные направления внутренней и внешней политики, "
              "представляет Россию внутри страны и в международных отношениях, является "
              "Верховным Главнокомандующим Вооружёнными Силами.")
    para(doc, "Президент избирается гражданами на шесть лет и не может занимать должность "
              "более двух сроков. Он назначает Председателя Правительства после утверждения "
              "кандидатуры Государственной Думой, формирует часть органов исполнительной власти, "
              "подписывает и обнародует федеральные законы, вправе распускать Государственную "
              "Думу в случаях, предусмотренных Конституцией. Свои решения Президент оформляет "
              "указами и распоряжениями.")
    para(doc, "Президент не входит ни в одну из трёх ветвей власти: его задача — обеспечивать "
              "согласованное функционирование и взаимодействие всех органов государственной "
              "власти.")

    heading(doc, "Федеральное Собрание")
    para(doc, "**Федеральное Собрание** — парламент Российской Федерации, представительный "
              "и законодательный орган. Он состоит из двух палат: Совета Федерации "
              "и Государственной Думы.")
    para(doc, "**Совет Федерации** представляет субъекты Федерации: в него входят по два "
              "представителя от каждого региона. Палата одобряет принятые Государственной Думой "
              "федеральные законы, утверждает важнейшие кадровые назначения и рассматривает "
              "вопросы войны и мира.")
    para(doc, "**Государственная Дума** состоит из 450 депутатов и избирается на пять лет. "
              "К её полномочиям относятся принятие федеральных законов, утверждение Председателя "
              "Правительства, решение вопроса о доверии Правительству, объявление амнистии. "
              "Закон проходит в Думе несколько чтений, затем одобряется Советом Федерации "
              "и подписывается Президентом.")

    heading(doc, "Правительство Российской Федерации")
    para(doc, "**Правительство** — высший исполнительный орган государства. В его состав входят "
              "Председатель Правительства, его заместители и федеральные министры. Правительство "
              "разрабатывает и исполняет федеральный бюджет, проводит единую финансовую, "
              "экономическую, социальную политику, управляет федеральной собственностью, "
              "обеспечивает оборону, безопасность и правопорядок.")
    para(doc, "Решения Правительства оформляются постановлениями и распоряжениями. "
              "Правительство подотчётно Президенту и ответственно перед Государственной Думой: "
              "Дума вправе выразить ему недоверие.")

    heading(doc, "Судебная власть и прокуратура")
    para(doc, "**Правосудие** в Российской Федерации осуществляется только судом. "
              "**Конституционный Суд** проверяет соответствие законов Конституции и даёт "
              "её официальное толкование. **Верховный Суд** — высший судебный орган "
              "по гражданским, уголовным, административным и экономическим делам. Судьи "
              "независимы, несменяемы и неприкосновенны.")
    para(doc, "**Прокуратура** — единая централизованная система органов, которая осуществляет "
              "надзор за соблюдением Конституции и исполнением законов. Прокуратура не входит "
              "ни в одну из трёх ветвей власти и действует самостоятельно.")

    heading(doc, "Субъекты Федерации и местное самоуправление")
    para(doc, "Россия — федеративное государство, в состав которого входят республики, края, "
              "области, города федерального значения, автономная область и автономные округа. "
              "Органы государственной власти субъектов действуют самостоятельно в пределах "
              "своих полномочий.")
    para(doc, "**Местное самоуправление** не входит в систему органов государственной власти "
              "(статья 12 Конституции). Оно решает вопросы местного значения: содержание "
              "жилищного фонда, благоустройство, организацию транспортного обслуживания "
              "и другие.")

    heading(doc, "Участие граждан")
    para(doc, "Граждане участвуют в формировании органов власти через выборы. Президент "
              "избирается прямым голосованием, Государственная Дума — по федеральному "
              "избирательному округу и одномандатным округам. Право избирать принадлежит "
              "гражданам, достигшим восемнадцати лет.")
    para(doc, "Помимо выборов, граждане вправе обращаться в органы государственной власти "
              "и местного самоуправления: обращение рассматривается в установленный срок, "
              "а ответ даётся по существу. Споры с органами власти разрешаются в суде.")

    table(
        doc,
        "Таблица 1 — Органы государственной власти Российской Федерации",
        ["Ветвь власти", "Органы", "Основное полномочие"],
        [
            ["Законодательная", "Федеральное Собрание: Совет Федерации "
                                "и Государственная Дума",
             "Принятие федеральных законов, представительство граждан и субъектов"],
            ["Исполнительная", "Правительство РФ, федеральные министерства и ведомства",
             "Исполнение законов, управление, разработка и исполнение бюджета"],
            ["Судебная", "Конституционный Суд, Верховный Суд, федеральные суды",
             "Правосудие и конституционный контроль"],
            ["Особый статус", "Президент РФ",
             "Глава государства, гарант Конституции, согласование работы органов власти"],
            ["Надзор", "Прокуратура РФ",
             "Надзор за соблюдением Конституции и исполнением законов"],
        ],
        [Cm(3.4), Cm(7.4), Cm(6.2)],
    )

    heading(doc, "Краткий вывод")
    para(doc, "Органы власти Российской Федерации образуют систему, в которой каждая ветвь "
              "отвечает за своё направление: парламент принимает законы, Правительство "
              "их исполняет, суды разрешают споры и проверяют законность, Президент "
              "обеспечивает согласованность работы. Такое устройство защищает от сосредоточения "
              "власти в одних руках и обеспечивает гражданам возможность отстаивать свои права.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
