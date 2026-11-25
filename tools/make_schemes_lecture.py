# -*- coding: utf-8 -*-
"""Лекция «Схемы подключения магнитного пускателя» (1МЭ-25, МДК 05.02)."""

import os
import re

from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\6-shemy-podklyucheniya"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Схемы подключения магнитного пускателя.docx")
SRC = r"C:\Users\user\.dsh\attachments\v1\objects"
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


def para(doc, text, indent=0.8, after=2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.07
    rich(p, text)


def heading(doc, text):
    p = doc.add_paragraph(style="Heading 1")
    rich(p, text)


def item(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.05
    rich(p, text, size=10.5)


def figure(doc, name, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    sh = p.add_run().add_picture(os.path.join(IMG, name), width=Cm(width_cm))
    sp = sh._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement("a:ln"); ln.set("w", "0"); ln.append(OxmlElement("a:noFill")); sp.append(ln)
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_before = Pt(1)
    c.paragraph_format.space_after = Pt(4)
    rich(c, caption, size=9, italic=True)


def prepare_images():
    os.makedirs(IMG, exist_ok=True)
    first = os.path.join(SRC, "1c", "1c114ffa45b2ab3eb431fa6f1bb6ccd7789a878dc17ed99e6e2f4d9167b2d5f3")
    second = os.path.join(SRC, "97", "97696986e48e566a6e850bab9498cbe55b45ebbd5d82ca748172795424212263")
    im = Image.open(first).convert("RGB")
    im = im.resize((im.width * 2, im.height * 2), Image.LANCZOS)
    im.save(os.path.join(IMG, "figure-1.png"), "PNG")
    im2 = Image.open(second).convert("RGB")
    im2.save(os.path.join(IMG, "figure-2.png"), "PNG")


def build():
    prepare_images()
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
    h1.paragraph_format.space_before = Pt(4)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Схемы подключения магнитного пускателя", size=14, bold=True)

    para(doc, "Схема подключения показывает, как соединить пускатель с питающей сетью, "
              "электродвигателем и аппаратами управления. На схемах различают силовую цепь, "
              "по которой ток идёт к двигателю, и цепь управления, питающую катушку пускателя "
              "и кнопки. Силовые провода изображают толстыми линиями, провода цепей управления — "
              "тонкими.")

    heading(doc, "Нереверсивная схема")
    para(doc, "В нереверсивной схеме двигатель вращается в одну сторону (рисунок 1). Напряжение "
              "от фаз *A*, *B*, *C* через трёхполюсный автоматический выключатель *SF1* подаётся "
              "на силовые контакты пускателя *KM1*, затем через нагревательные элементы тепловых "
              "реле *KK1* и *KK2* — на двигатель.")
    para(doc, "Цепь управления питается от фазы *A* и нулевого провода *N* через автоматический "
              "выключатель *SF2*. В неё последовательно включены кнопка «Стоп» *SB1* с размыкающим "
              "контактом, кнопка «Пуск» *SB2* с замыкающим контактом, контакты тепловых реле "
              "*KK1*, *KK2* и катушка пускателя *KM1*. Параллельно кнопке «Пуск» подключён "
              "блок-контакт *KM1* — он удерживает цепь после отпускания кнопки.")
    para(doc, "При нажатии «Пуск» катушка получает питание, силовые контакты замыкаются, "
              "двигатель запускается, а блок-контакт шунтирует кнопку. Нажатие «Стоп» разрывает "
              "цепь катушки — пускатель отключается. При перегрузке контакт теплового реле "
              "размыкается, и двигатель останавливается.")
    figure(doc, "figure-1.png", 9.5, "Рисунок 1 — Нереверсивная схема подключения пускателя")

    heading(doc, "Обозначения на схемах")
    for text in [
        "**A, B, C** — фазы питающей сети; **N** — нулевой рабочий провод.",
        "**SF1** — трёхполюсный автоматический выключатель силовой цепи; "
        "**SF2** — автоматический выключатель цепи управления.",
        "**KM1, KM2** — магнитные пускатели: силовые контакты в цепи двигателя и катушка "
        "в цепи управления.",
        "**SB1** — кнопка «Стоп» с размыкающим контактом; **SB2, SB3** — кнопки «Пуск» "
        "с замыкающими контактами.",
        "**KK1, KK2** — тепловые реле: нагревательные элементы в силовой цепи, размыкающие "
        "контакты — в цепи катушки.",
        "Блок-контакты **KM1** и **KM2** в цепи управления — самоблокировка и взаимная "
        "блокировка пускателей.",
        "**M** — трёхфазный асинхронный электродвигатель.",
    ]:
        item(doc, "— " + text)

    heading(doc, "Реверсивная схема")
    para(doc, "Реверсивная схема содержит два пускателя, *KM1* и *KM2* (рисунок 2). Их силовые "
              "контакты подключены так, что при включении второго пускателя две фазы на двигателе "
              "меняются местами — двигатель вращается в обратную сторону.")
    para(doc, "Кнопки *SB2* и *SB3* подают питание на катушки *KM1* и *KM2* соответственно. "
              "В цепь каждой катушки включён размыкающий блок-контакт другого пускателя: "
              "при работающем *KM1* его контакт разрывает цепь *KM2*, и наоборот. Это электрическая "
              "блокировка, исключающая одновременное включение. Дополнительно применяют "
              "механическую блокировку — коромысло между пускателями.")
    para(doc, "Для изменения направления вращения двигатель останавливают кнопкой «Стоп» *SB1*, "
              "а затем включают другой пускатель. Тепловые реле *KK1* и *KK2* защищают двигатель "
              "от перегрузки в обоих направлениях.")
    figure(doc, "figure-2.png", 12.5, "Рисунок 2 — Реверсивная схема подключения пускателя")

    heading(doc, "Краткий вывод")
    para(doc, "Нереверсивная схема обеспечивает пуск и останов двигателя с самоблокировкой, "
              "реверсивная добавляет возможность изменить направление вращения и защиту "
              "от одновременного включения двух пускателей. В обеих схемах тепловые реле "
              "разрывают цепь катушки при перегрузке, поэтому двигатель отключается автоматически.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
