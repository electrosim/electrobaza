# -*- coding: utf-8 -*-
"""Лекция «Однофазный переменный ток» (ОП.03) по материалам учебника."""

import os
import re
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1a-25-3a-26-11\op-03\3-peremenny-tok"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Однофазный переменный ток.docx")
SRC = os.path.join(os.environ["TEMP"], "op03")
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


def bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, it)


def numbered(doc, items):
    for i, it in enumerate(items, 1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, f"{i}.\t{it}")


def formula(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
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
    c.paragraph_format.space_after = Pt(4)
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
    h1.paragraph_format.space_before = Pt(4)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Однофазный переменный ток", size=14, bold=True)

    para(doc, "Переменным током называется такой электрический ток, который периодически изменяется "
              "по величине и направлению. Для его получения используют электромашинные генераторы, работа "
              "которых основана на явлении электромагнитной индукции. Почти вся электроэнергия вырабатывается "
              "в виде энергии переменного тока.")
    para(doc, "Переменный ток позволяет получать напряжение разного уровня: высокое — для передачи энергии "
              "на большие расстояния, низкое — для питания потребителей. Генераторы и двигатели переменного "
              "тока проще по устройству, надёжнее в работе и удобнее в эксплуатации, поэтому переменный ток "
              "получил преимущественное распространение. Наибольшее применение имеет синусоидальный ток: его "
              "изменение происходит плавно, без скачков и резких перепадов, что благоприятно сказывается "
              "на работе электрических машин и аппаратов.")

    heading(doc, "Основные параметры синусоидального тока")
    para(doc, "Мгновенное значение синусоидального тока описывается формулой:")
    formula(doc, "*i* = *I*~m~ sin(ω*t* + ψ)")
    para(doc, "где *I*~m~ — максимальное значение (амплитуда) тока; ω — угловая частота; ψ — начальная фаза, "
              "то есть значение аргумента в начальный момент времени при *t* = 0.")
    figure(doc, "fig-4-1.png", 9.5, "Рисунок 1 — Временная диаграмма синусоидального тока")

    heading(doc, "Период, частота и угловая частота")
    para(doc, "Период — время, в течение которого переменная ЭДС, напряжение или ток совершают одно полное "
              "изменение по величине и направлению (один цикл). Период обозначается буквой *T* и измеряется "
              "в секундах.")
    para(doc, "Частота — число полных изменений за одну секунду. Она обозначается буквой *f* и измеряется "
              "в герцах. Для больших частот применяют килогерцы и мегагерцы: 1 кГц = 1000 Гц, "
              "1 МГц = 1000 кГц = 1 000 000 Гц. Чем больше частота, тем короче период, поэтому частота — "
              "величина, обратная периоду:")
    formula(doc, "*f* = 1 / *T*")
    para(doc, "При вращении витка в магнитном поле один оборот соответствует 360°, или 2π радиан. Угловая "
              "скорость вращения выражается в радианах в секунду и называется угловой частотой:")
    formula(doc, "ω = 2π / *T* = 2π*f*")
    para(doc, "Угловая частота, выраженная в радианах в секунду, больше частоты, выраженной в герцах, в 2π раз.")

    heading(doc, "Мгновенные и максимальные значения")
    para(doc, "Значения переменной ЭДС, силы тока, напряжения и мощности в любой момент времени называют "
              "мгновенными и обозначают строчными буквами *e*, *i*, *u*, *p*:")
    formula(doc, "*e* = *E*~m~ sin(ω*t* + ψ)")
    formula(doc, "*i* = *I*~m~ sin(ω*t* + ψ)")
    formula(doc, "*u* = *U*~m~ sin(ω*t* + ψ)")
    formula(doc, "*p* = *P*~m~ sin(ω*t* + ψ)")
    para(doc, "Максимальным значением (амплитудой) называется наибольшая величина, которой переменная "
              "величина достигает за один период. Амплитуда ЭДС обозначается *E*~m~, напряжения — *U*~m~, "
              "тока — *I*~m~.")

    heading(doc, "Действующее значение")
    para(doc, "Действующим значением переменного тока называется такая сила постоянного тока, которая, "
              "протекая через равное сопротивление и за одно и то же время, выделяет столько же теплоты, "
              "сколько и переменный ток. Электроизмерительные приборы — амперметр и вольтметр — показывают "
              "именно действующие значения.")
    para(doc, "Для синусоидального тока действующее значение меньше максимального в √2 ≈ 1,41 раза:")
    formula(doc, "*I* = *I*~m~ / √2 = *I*~m~ / 1,41 = 0,707 *I*~m~")
    para(doc, "Аналогично для ЭДС и напряжения:")
    formula(doc, "*E* = *E*~m~ / √2 = 0,707 *E*~m~")
    formula(doc, "*U* = *U*~m~ / √2 = 0,707 *U*~m~")
    para(doc, "Пример. Вольтметр, подключённый к зажимам цепи, показывает действующее напряжение 36 В. "
              "Максимальное значение больше действующего в √2 раз:")
    formula(doc, "*U*~m~ = *U* · √2 = 36 · 1,41 ≈ 50,8 В")

    heading(doc, "Сдвиг фаз")
    para(doc, "Две синусоидальные величины могут изменяться неодинаково и достигать максимума в разные "
              "моменты времени. Сдвиг фаз между напряжением и током — алгебраическая величина, равная "
              "разности начальных фаз:")
    formula(doc, "φ = ψ~u~ − ψ~i~")
    figure(doc, "fig-4-2.png", 9.5,
           "Рисунок 2 — Диаграмма изменения величин переменного тока и напряжения, не совпадающих по фазе")
    para(doc, "Знак сдвига фаз показывает, какая величина опережает другую:")
    bullets(doc, [
        "если φ > 0, напряжение опережает ток по фазе;",
        "если φ < 0, напряжение отстаёт от тока по фазе;",
        "если φ = 0, ток и напряжение совпадают по фазе.",
    ])

    heading(doc, "Краткий вывод")
    para(doc, "Переменный синусоидальный ток характеризуется амплитудой, периодом, частотой, начальной фазой "
              "и действующим значением. Приборы показывают действующие значения, которые в √2 раза меньше "
              "амплитудных. Сдвиг фаз определяет, опережает напряжение ток или отстаёт от него.")

    heading(doc, "Контрольные вопросы")
    numbered(doc, [
        "Что называется переменным током?",
        "В каких единицах измеряется частота переменного тока?",
        "Какое значение переменного тока наносят на шкалах амперметров и вольтметров?",
        "Что такое сдвиг фаз между напряжением и током?",
    ])

    heading(doc, "Задания для самостоятельной работы")
    numbered(doc, [
        "ЭДС генератора определяется формулой *e* = 36 sin(314*t* + π/4). Определите начальную фазу "
        "и значение ЭДС при *t* = 0,15 с.",
        "ЭДС генератора имеет максимальное значение 110 В, частоту 400 Гц и начальную фазу 30°. "
        "Запишите выражение для мгновенного значения ЭДС.",
        "Вольтметр показывает 220 В. Определите максимальное значение напряжения.",
        "Ток в цепи *i* = 284 sin(314*t* + π/2). Найдите действующее значение, частоту и начальную фазу тока.",
    ])

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


def main():
    os.makedirs(IMG, exist_ok=True)
    shutil.copyfile(os.path.join(SRC, "fig4-1.png"), os.path.join(IMG, "fig-4-1.png"))
    shutil.copyfile(os.path.join(SRC, "fig4-2.png"), os.path.join(IMG, "fig-4-2.png"))
    build()


if __name__ == "__main__":
    main()
