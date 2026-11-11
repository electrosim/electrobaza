# -*- coding: utf-8 -*-
"""Лекция «Устройство защитного отключения» (1МЭ-25, МДК 05.02)."""

import os
import re
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\4-uzo"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Устройство защитного отключения.docx")
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
    for tok in re.split(r"(~[^~]+~|\^[^\^]+\^|\*[^*]+\*)", text):
        if not tok:
            continue
        if tok.startswith("~") and tok.endswith("~") and len(tok) > 2:
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.font.subscript = True
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


def figure(doc, name, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    sh = p.add_run().add_picture(os.path.join(IMG, name), width=Cm(width_cm))
    sp = sh._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement("a:ln"); ln.set("w", "0"); ln.append(OxmlElement("a:noFill")); sp.append(ln)
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_before = Pt(1)
    c.paragraph_format.space_after = Pt(3)
    rich(c, caption, size=9, italic=True)


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
    h1.paragraph_format.space_before = Pt(3)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(3)
    rich(p, "Устройство защитного отключения", size=14, bold=True)

    para(doc, "Устройство защитного отключения (УЗО) — коммутационный аппарат, который отключает цепь, "
              "когда часть тока уходит в землю. Оно защищает человека от поражения электрическим током "
              "при утечке и не заменяет автоматический выключатель: от сверхтока и короткого замыкания "
              "защищает именно автомат.")

    heading(doc, "Внешний вид и устройство")
    para(doc, "На корпусе УЗО есть рукоятка включения, кнопка «Тест», клеммы фазного и нулевого проводников "
              "и маркировка: номинальный ток, напряжение и уставка по дифференциальному току. Внутри "
              "находятся контактная группа, дифференциальный трансформатор, реле и механизм расцепления.")
    figure(doc, "rcd-photo.png", 5.2, "Рисунок 1 — Внешний вид устройства защитного отключения")

    heading(doc, "Принцип действия")
    para(doc, "Фазный и нулевой проводники образуют две первичные обмотки дифференциального трансформатора. "
              "В исправной цепи токи в них равны и направлены встречно, поэтому магнитные потоки "
              "компенсируются и во вторичной обмотке ЭДС нет. Когда часть тока утекает через повреждённую "
              "изоляцию или тело человека, равенство нарушается: появляется дифференциальный ток "
              "*I*~Δ~, магнитный поток в сердечнике не скомпенсирован, во вторичной обмотке наводится ЭДС, "
              "реле срабатывает и размыкает контакты.")
    figure(doc, "rcd-scheme.png", 8.6, "Рисунок 2 — Схема устройства и принцип действия УЗО")

    heading(doc, "Основные характеристики")
    para(doc, "Уставка УЗО — 10, 30, 100 или 300 мА. Для защиты людей применяют 30 мА, для защиты "
              "от возгорания — 100–300 мА. Время срабатывания не превышает 30 мс. Кнопка «Тест» создаёт "
              "искусственный дифференциальный ток и проверяет работоспособность — нажимать её нужно "
              "регулярно, обычно раз в месяц.")

    heading(doc, "Включение в цепь")
    para(doc, "УЗО ставят последовательно с нагрузкой, как правило после автоматического выключателя. "
              "Через его трансформатор должны проходить оба рабочих проводника. Нельзя соединять нулевой "
              "проводник после УЗО с нулём других цепей: часть тока пойдёт в обход, и устройство будет "
              "срабатывать ложно.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


def main():
    os.makedirs(IMG, exist_ok=True)
    src = r"C:\Users\user\.dsh\attachments\v1\objects"
    shutil.copyfile(os.path.join(src, "c5", "c5e34d68c0eb57c44a620280d09a0a831b8ed2d7fe86a4b18afc1e5ecdb86b15"),
                    os.path.join(IMG, "rcd-photo.png"))
    shutil.copyfile(os.path.join(src, "76", "76b4370fdd13f662c5f1b1067e2fb41d6a728d294eb93ed7f6d37cf592ad0cf5"),
                    os.path.join(IMG, "rcd-scheme.png"))
    build()


if __name__ == "__main__":
    main()
