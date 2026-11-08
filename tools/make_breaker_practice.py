# -*- coding: utf-8 -*-
"""Практическая работа «Принцип действия и установка автоматического выключателя»."""

import os
import re

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\circuit-breaker-selection"
IMG_DIR = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Практическая работа — Принцип действия и установка автоматического выключателя.docx")
FONT = "Times New Roman"

GREEN = (47, 107, 79)
SOFT = (238, 244, 239)
INK = (20, 35, 27)
GREY = (90, 107, 97)
ARIAL = r"C:\Windows\Fonts\arial.ttf"
ARIAL_BOLD = r"C:\Windows\Fonts\arialbd.ttf"


# ------------------------------------------------------------------ рисунки

def fnt(size, bold=False):
    return ImageFont.truetype(ARIAL_BOLD if bold else ARIAL, size)


def ctext(d, xy, text, f, fill=INK):
    box = d.textbbox((0, 0), text, font=f)
    d.text((xy[0] - (box[2] - box[0]) / 2, xy[1] - (box[3] - box[1]) / 2 - box[1]), text, font=f, fill=fill)


def box(d, x1, y1, x2, y2, lines, f, fill=SOFT, outline=GREEN, radius=14, text=INK):
    d.rounded_rectangle((x1, y1, x2, y2), radius=radius, fill=fill, outline=outline, width=3)
    step = 30
    start = (y1 + y2) / 2 - step * (len(lines) - 1) / 2
    for i, line in enumerate(lines):
        ctext(d, ((x1 + x2) / 2, start + i * step), line, f, text)


def arrow(d, x1, y1, x2, y2):
    d.line((x1, y1, x2, y2), fill=GREEN, width=3)
    dx, dy = x2 - x1, y2 - y1
    ln = max((dx * dx + dy * dy) ** 0.5, 1)
    ux, uy = dx / ln, dy / ln
    px, py = -uy, ux
    s = 12
    d.polygon([(x2, y2), (x2 - ux * s + px * s * 0.6, y2 - uy * s + py * s * 0.6),
               (x2 - ux * s - px * s * 0.6, y2 - uy * s - py * s * 0.6)], fill=GREEN)


def draw_scheme(path):
    W, H = 1600, 760
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    f20, f22, f24 = fnt(20), fnt(22), fnt(24, True)

    box(d, 40, 300, 230, 400, ["Ввод"], f22)
    box(d, 280, 300, 560, 400, ["Тепловой", "расцепитель"], f20)
    box(d, 610, 300, 960, 400, ["Электромагнитный", "расцепитель"], f20)
    box(d, 1010, 300, 1290, 400, ["Главные", "контакты"], f20)
    box(d, 1340, 300, 1560, 400, ["Вывод"], f22)
    for x in (230, 560, 960, 1290):
        arrow(d, x + 4, 350, x + 46, 350)

    box(d, 1010, 120, 1290, 220, ["Дугогасительная", "камера"], f20, fill="white", outline=(150, 170, 158))
    d.line((1150, 300, 1150, 222), fill=(150, 170, 158), width=3)

    box(d, 610, 520, 1290, 640, ["Механизм свободного", "расцепления", "и рукоятка"], f20)
    d.line((740, 520, 740, 402), fill=GREEN, width=3)
    d.line((900, 520, 900, 402), fill=GREEN, width=3)
    d.text((60, 660), "Ток проходит последовательно через тепловой и электромагнитный расцепители", font=f20, fill=GREY)
    im.save(path)


def draw_curve(path):
    W, H = 1600, 900
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    f20, f22 = fnt(20), fnt(22)

    left, right, top, bottom = 170, 1480, 70, 740
    d.line((left, bottom, right, bottom), fill=INK, width=3)
    d.line((left, top, left, bottom), fill=INK, width=3)

    # логарифмическая ось токов: 1..20 Iном
    import math

    def px(k):
        return left + (math.log10(k) / math.log10(20)) * (right - left)

    def py(t):
        return bottom - ((math.log10(t) + 2) / 6) * (bottom - top)  # 0,01..10000 с

    for k in (1, 2, 5, 10, 20):
        d.line((px(k), bottom, px(k), bottom + 10), fill=INK, width=2)
        ctext(d, (px(k), bottom + 34), str(k), f20)
    for t, label in ((0.01, "0,01"), (1, "1"), (100, "100"), (10000, "10 000")):
        d.line((left - 10, py(t), left, py(t)), fill=INK, width=2)
        d.text((20, py(t) - 12), label, font=f20, fill=INK)

    for k in (2, 5, 10):
        d.line((px(k), top, px(k), bottom), fill=(226, 232, 227), width=2)

    # тепловая зона: t = 60 / (k^2 - 1), пока не сработает электромагнитный
    pts = []
    k = 1.05
    while k < 10:
        t = 60 / (k * k - 1)
        if 0.01 <= t <= 10000:
            pts.append((px(k), py(t)))
        k += 0.05
    d.line(pts, fill=GREEN, width=6, joint="curve")

    # электромагнитная отсечка 10 Iном
    d.line((px(10), py(100), px(10), py(0.01)), fill=(190, 70, 50), width=6)
    d.text((px(10) + 20, py(0.03)), "отсечка 10·Iном", font=f22, fill=(190, 70, 50))
    d.text((px(1.5), py(200)), "тепловая зона", font=f22, fill=GREEN)

    ctext(d, ((left + right) / 2, 830), "Отношение тока к номинальному I / Iном", f22)
    d.text((10, top - 40), "Время срабатывания, с", font=f22, fill=INK)
    d.text((left + 20, top + 20), "характеристика типа C", font=f20, fill=GREY)
    im.save(path)


# ------------------------------------------------------------------ docx

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


def figure(doc, path, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    sh = p.add_run().add_picture(path, width=Cm(width_cm))
    sp = sh._inline.graphic.graphicData.pic.spPr
    ln = OxmlElement("a:ln"); ln.set("w", "0"); ln.append(OxmlElement("a:noFill")); sp.append(ln)
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_before = Pt(1)
    c.paragraph_format.space_after = Pt(4)
    rich(c, caption, size=9, italic=True)


VARIANTS = [
    (1, "3,0", "0,90", "2,5"),
    (2, "4,5", "0,85", "2,5"),
    (3, "2,2", "0,80", "1,5"),
    (4, "5,5", "0,90", "4,0"),
    (5, "1,5", "0,85", "1,5"),
    (6, "7,0", "0,88", "4,0"),
    (7, "3,5", "0,92", "2,5"),
    (8, "6,0", "0,86", "4,0"),
]


def table(doc):
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_before = Pt(4)
    cap.paragraph_format.space_after = Pt(2)
    rich(cap, "Таблица 1 — Варианты заданий", size=9, italic=True)

    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2.6), Cm(4.2), Cm(3.4), Cm(4.2)]
    head = ["Вариант", "Мощность P, кВт", "cos φ", "Сечение S, мм²"]
    for i, text in enumerate(head):
        cell = t.rows[0].cells[i]
        cell.width = widths[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rich(p, text, size=9.5, bold=True)
    for row in VARIANTS:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].width = widths[i]
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            rich(p, str(val), size=9.5)
    # компактные отступы в ячейках
    for r in t.rows:
        for c in r.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.line_spacing = 1.0


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
    rich(p, "Принцип действия и установка автоматического выключателя", size=14, bold=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Практическая работа", size=11, italic=True)

    heading(doc, "Цель работы")
    para(doc, "Научиться выбирать автоматический выключатель по расчётному току нагрузки, проверять его "
              "по допустимому току кабеля и условиям пуска, читать время-токовую характеристику.")

    heading(doc, "Краткие сведения")
    para(doc, "Автоматический выключатель защищает линию от перегрузки и короткого замыкания. Ток проходит "
              "последовательно через тепловой и электромагнитный расцепители. При перегрузке нагревается "
              "биметаллическая пластина и с выдержкой времени размыкает контакты; при коротком замыкании "
              "срабатывает электромагнитный расцепитель — почти мгновенно. Механизм свободного расцепления "
              "не даёт удерживать контакты включёнными при аварии, а дугогасительная камера гасит дугу.")
    figure(doc, os.path.join(IMG_DIR, "breaker-scheme.png"), 9.5, "Рисунок 1 — Устройство автоматического выключателя")

    heading(doc, "Исходные данные")
    para(doc, "Напряжение сети U = 220 В, кратность пускового тока k = 6. Длительно допустимый ток кабеля "
              "принят: 1,5 мм² — 19 А; 2,5 мм² — 27 А; 4 мм² — 38 А. Номинальный ряд выключателей: "
              "6, 10, 16, 20, 25, 32, 40 А. Вариант выбирается по последней цифре номера в журнале.")
    table(doc)

    heading(doc, "Расчётные формулы")
    para(doc, "Расчётный ток нагрузки:")
    formula(doc, "*I*~расч~ = *P* / (*U* · cos φ)")
    para(doc, "Условия выбора номинального тока выключателя:")
    formula(doc, "*I*~ном~ ≥ *I*~расч~")
    formula(doc, "*I*~ном~ ≤ *I*~доп~")
    para(doc, "Пусковой ток и проверка электромагнитного расцепителя:")
    formula(doc, "*I*~пуск~ = *k* · *I*~расч~")
    formula(doc, "*I*~отс~ = *K* · *I*~ном~ ≥ 1,2 · *I*~пуск~")
    para(doc, "где *K* — кратность отсечки: 5 для характеристики B, 10 для C, 20 для D.")

    heading(doc, "Пример расчёта (вариант 1)")
    para(doc, "Дано: P = 3,0 кВт; U = 220 В; cos φ = 0,90; S = 2,5 мм². Расчётный ток:")
    formula(doc, "*I*~расч~ = 3000 / (220 · 0,90) = 15,2 А")
    para(doc, "Принимаем выключатель 16 А: 16 ≥ 15,2 и 16 ≤ 27 — оба условия выполняются. Пусковой ток:")
    formula(doc, "*I*~пуск~ = 6 · 15,2 = 91 А")
    para(doc, "Для характеристики C отсечка равна 10 · 16 = 160 А, что больше 1,2 · 91 = 109 А. Условие "
              "выполняется, принимаем выключатель C16.")

    heading(doc, "График")
    para(doc, "Время-токовая характеристика типа C показывает зоны работы расцепителей: при токах до "
              "10·*I*~ном~ действует тепловая защита, при больших токах — электромагнитная отсечка.")
    figure(doc, os.path.join(IMG_DIR, "time-current-curve.png"), 9.5, "Рисунок 2 — Время-токовая характеристика типа C")

    heading(doc, "Задания")
    numbered(doc, [
        "По своему варианту вычислить расчётный ток и выбрать номинал выключателя из стандартного ряда.",
        "Проверить условие *I*~ном~ ≤ *I*~доп~; при невыполнении увеличить сечение и повторить выбор.",
        "Вычислить пусковой ток и проверить отсечку для характеристики C.",
        "Для варианта 6 подобрать сечение, при котором условие по допустимому току выполняется.",
        "Построить точку своего расчётного тока на графике и пояснить, какая защита сработает при перегрузке.",
    ])

    heading(doc, "Что включить в отчёт")
    bullets(doc, [
        "цель работы и исходные данные варианта;",
        "формулы и расчёт с подстановкой чисел;",
        "таблицу результатов: *I*~расч~, *I*~ном~, *I*~доп~, *I*~пуск~, *I*~отс~, вывод о выборе;",
        "график с отмеченными точками;",
        "вывод о правильности выбора выключателя.",
    ])

    heading(doc, "Контрольные вопросы")
    numbered(doc, [
        "Какой расцепитель срабатывает при перегрузке, а какой — при коротком замыкании?",
        "Почему номинальный ток выключателя не должен превышать допустимый ток кабеля?",
        "Что означает маркировка C16?",
        "Почему отсечку проверяют по пусковому току?",
        "Что делать, если расчётный ток больше допустимого тока выбранного сечения?",
    ])

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


def main():
    os.makedirs(IMG_DIR, exist_ok=True)
    draw_scheme(os.path.join(IMG_DIR, "breaker-scheme.png"))
    draw_curve(os.path.join(IMG_DIR, "time-current-curve.png"))
    build()


if __name__ == "__main__":
    main()
