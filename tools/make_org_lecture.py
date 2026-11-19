# -*- coding: utf-8 -*-
"""Лекция «Организация труда при выполнении электромонтажных работ» (1ЭО-25, МДК 04.01)."""

import os
import re

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1eo-25\mdk-04-01\3-organizaciya-truda"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Организация труда при выполнении электромонтажных работ.docx")
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
            r = p.add_run(tok[1:-1]); _rfonts(r._element.get_or_add_rPr()); r.font.italic = True
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


def formula(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    rich(p, text, size=11.5, bold=True)


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


def shade(cell, color="E8F0EC"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def table(doc, caption, header, rows):
    c = doc.add_paragraph()
    c.paragraph_format.space_before = Pt(4)
    c.paragraph_format.space_after = Pt(2)
    c.paragraph_format.keep_with_next = True
    rich(c, caption, size=9, italic=True)

    t = doc.add_table(rows=1, cols=len(header))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    widths = [Cm(6.4), Cm(3.2), Cm(3.4), Cm(4.0)]
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
            if i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
    h1.paragraph_format.space_before = Pt(4)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Организация труда при выполнении электромонтажных работ", size=14, bold=True)

    para(doc, "Организация труда — это заранее продуманный порядок выполнения работы: кто что делает, "
              "какими инструментами, в какие сроки и по какой документации. На электромонтаже от неё "
              "зависят три вещи одновременно: производительность, качество и безопасность.")

    heading(doc, "Задачи и элементы организации труда")
    para(doc, "К задачам относят рациональное разделение работы между исполнителями, оснащение рабочих "
              "мест, нормирование времени и контроль результата. Элементы организации труда: разделение "
              "и кооперация, организация рабочего места, нормирование, оплата, дисциплина и охрана труда.")
    para(doc, "Разделение труда строится по квалификации: электромонтажник 3-го разряда выполняет "
              "подготовительные операции, 4–5-го — монтаж щитов и пусконаладку. Кооперация связывает "
              "исполнителей в звено или бригаду, чтобы операции шли последовательно и без простоев.")

    heading(doc, "Формы организации бригад")
    para(doc, "Специализированная бригада выполняет один вид работ, например прокладку кабельных трасс. "
              "Комплексная бригада закрывает весь цикл — от разметки до сдачи щита. Звено из двух-трёх "
              "человек удобно на небольших объектах, а индивидуальная работа допускается только там, "
              "где не требуется постоянное взаимодействие.")
    para(doc, "Совмещение профессий сокращает численность: электромонтажник с допуском стропальщика "
              "обходится без отдельного рабочего на подъёме материалов.")

    heading(doc, "Рабочее место электромонтажника")
    para(doc, "Рабочее место организуют так, чтобы инструмент и материалы лежали в зоне досягаемости, "
              "а проходы оставались свободными. Инструмент делят на основной, вспомогательный и "
              "измерительный; его хранят в сумке или ящике с отдельными отделениями, чтобы не искать "
              "мелочь на дне.")
    figure(doc, "toolbox.jpg", 9.0, "Рисунок 1 — Рабочее место и инструмент электромонтажника")

    heading(doc, "Нормирование труда")
    para(doc, "Норма времени Н~вр~ — время на единицу работы, норма выработки Н~выр~ — объём работы "
              "за смену. Их связывает длительность смены Т~см~:")
    formula(doc, "Н~выр~ = Т~см~ / Н~вр~")
    para(doc, "Сдельная расценка Р~сд~ показывает оплату за единицу работы при часовой ставке С~час~:")
    formula(doc, "Р~сд~ = С~час~ × Н~вр~")
    para(doc, "Срок выполнения объёма Q бригадой из Ч человек с коэффициентом выполнения норм К~вн~:")
    formula(doc, "Д = Q × Н~вр~ / (Ч × Т~см~ × К~вн~)")
    para(doc, "Производительность труда П~т~ — объём работы за единицу времени, а её рост считают "
              "в процентах к базовому значению:")
    formula(doc, "П~т~ = Q / Т;   рост = (П~т2~ / П~т1~ − 1) × 100 %")

    heading(doc, "Примеры расчёта")
    para(doc, "Пример 1. Норма времени на установку розетки 0,4 чел.-ч, смена 8 ч. Н~выр~ = 8 / 0,4 = "
              "20 розеток за смену.")
    para(doc, "Пример 2. Бригаде из 4 человек нужно установить 400 розеток при той же норме и "
              "К~вн~ = 1,1. Д = 400 × 0,4 / (4 × 8 × 1,1) = 160 / 35,2 ≈ 4,5 смены.")
    para(doc, "Пример 3. При часовой ставке 350 руб. расценка за розетку Р~сд~ = 350 × 0,4 = 140 руб., "
              "а за 400 штук — 56 000 руб.")
    para(doc, "Пример 4. Если выработка выросла с 20 до 25 розеток за смену, рост производительности "
              "составил (25 / 20 − 1) × 100 = 25 %.")

    heading(doc, "Оплата труда в бригаде")
    para(doc, "При сдельной оплате заработок зависит от объёма: З~сд~ = Р~сд~ × Q. Если работа "
              "выполнена бригадой, общий фонд распределяют по коэффициенту трудового участия КТУ "
              "с учётом отработанного времени Т:")
    formula(doc, "З~i~ = Ф × (К~i~ × Т~i~) / Σ(К × Т)")
    para(doc, "Пример 5. Фонд бригады 120 000 руб. Трое рабочих отработали по 40 ч с КТУ 1,2; 1,0 и 0,9. "
              "Сумма К × Т = 48 + 40 + 36 = 124. Первый получит 120 000 × 48 / 124 ≈ 46 452 руб., "
              "второй — 38 710 руб., третий — 34 839 руб.")
    para(doc, "Повременная оплата применяется на пусконаладке и работах с ненормируемым объёмом. "
              "В обоих случаях премия привязывается к качеству и сроку, а не только к количеству.")

    heading(doc, "Контроль и приёмка работ")
    para(doc, "Качество проверяют по этапам: скрытые работы оформляют актом до закрытия конструкции, "
              "затем выполняют приёмо-сдаточные испытания. Исполнительная документация — схемы, "
              "протоколы измерений, акты — подтверждает, что монтаж выполнен по проекту.")
    para(doc, "Мастер контролирует соблюдение технологии и ведёт учёт выполненных объёмов. Такой учёт "
              "связывает нормирование с оплатой: по нему начисляют заработок и оценивают "
              "производительность бригады.")

    heading(doc, "Нормы времени на типовые операции")
    para(doc, "В таблице приведены учебные нормы: по ним удобно сравнивать операции между собой.")
    table(
        doc,
        "Таблица 1 — Нормы времени на типовые электромонтажные операции",
        ["Операция", "Единица", "Н~вр~, чел.-ч", "Разряд"],
        [
            ["Установка розетки", "1 шт.", "0,4", "3"],
            ["Установка выключателя", "1 шт.", "0,35", "3"],
            ["Монтаж светильника", "1 шт.", "0,6", "4"],
            ["Прокладка кабеля в гофре", "1 м", "0,08", "3"],
            ["Монтаж распределительного щита", "1 шт.", "4,0", "5"],
            ["Подключение электродвигателя", "1 шт.", "1,2", "5"],
        ],
    )

    heading(doc, "Документация и допуск к работе")
    para(doc, "Работы ведут по проекту производства работ и технологической карте: в них указаны "
              "последовательность операций, инструмент и требования безопасности. В действующих "
              "электроустановках монтаж выполняют по наряду-допуску или распоряжению, а бригаду "
              "инструктируют на рабочем месте под подпись.")
    figure(doc, "brigade.jpg", 10.0, "Рисунок 2 — Бригада на объекте: подготовка материалов и кабеля")

    heading(doc, "Безопасность при организации работ")
    para(doc, "Монтаж в действующих электроустановках выполняют по наряду-допуску, в остальных случаях — "
              "по распоряжению. К работе допускают лиц с группой по электробезопасности не ниже "
              "требуемой для данного вида работ; бригаду инструктируют на рабочем месте и обеспечивают "
              "средствами защиты.")
    para(doc, "Перед сменой проверяют исправность инструмента, наличие ограждений и заземлений. Работы "
              "в зоне действия другого оборудования согласуют с его владельцем, чтобы не создавать "
              "опасность для соседних бригад.")

    heading(doc, "Вопросы для самопроверки")
    qs = [
        "Что такое организация труда и какие элементы она включает?",
        "Чем отличается специализированная бригада от комплексной?",
        "Как связаны норма времени и норма выработки?",
        "Как рассчитать срок выполнения работ бригадой?",
        "Как распределяется бригадный заработок по КТУ?",
        "Какие документы оформляют перед началом монтажа?",
    ]
    for i, q in enumerate(qs, 1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, f"{i}.\t{q}")

    heading(doc, "Краткий вывод")
    para(doc, "Организация труда связывает людей, инструмент, нормы и документы в один порядок работы. "
              "Расчёт норм и численности делает этот порядок проверяемым: видно, сколько времени и "
              "средств требует объект и где теряется производительность.")

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print("saved:", OUT)


if __name__ == "__main__":
    build()
