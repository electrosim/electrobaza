# -*- coding: utf-8 -*-
"""Лекция «Устройство и принцип действия магнитного пускателя» (1МЭ-25, МДК 05.02)."""

import os
import re
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\5-magnitnyy-puskatel"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Устройство и принцип действия магнитного пускателя.docx")
FONT = "Times New Roman"
SRC = r"C:\Users\user\.dsh\attachments\v1\objects"


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
    h1.paragraph_format.space_before = Pt(4)
    h1.paragraph_format.space_after = Pt(2)
    h1.paragraph_format.line_spacing = 1.05
    h1.paragraph_format.keep_with_next = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, "Устройство и принцип действия магнитного пускателя", size=14, bold=True)

    para(doc, "Магнитный пускатель — аппарат дистанционного управления электродвигателем: он замыкает "
              "и размыкает силовую цепь по команде из цепи управления. Основу пускателя составляет "
              "электромагнитный контактор, дополненный тепловым реле и кнопками.")

    heading(doc, "Общее устройство")
    para(doc, "Пускатель собран в корпусе из изоляционного материала. На нём выведены зажимы: *A1* и *A2* — "
              "выводы катушки, *1L1*, *3L2*, *5L3* — вход силовых контактов, *2T1*, *4T2*, *6T3* — выход "
              "к нагрузке, *13* и *14* — вспомогательный (блок-)контакт.")
    para(doc, "Основные части аппарата: электромагнитная система (катушка, неподвижный сердечник, "
              "подвижный якорь), контактная система (подвижные и неподвижные силовые контакты, "
              "вспомогательные контакты), возвратные пружины, дугогасительные камеры и корпус.")
    figure(doc, "puskatel-photo.png", 7.0, "Рисунок 1 — Контактор магнитный: общий вид")

    heading(doc, "Устройство магнитной системы")
    para(doc, "Сердечник и якорь набраны из тонких листов электротехнической стали: шихтовка снижает "
              "вихревые токи и нагрев магнитопровода. Катушка надета на сердечник, якорь связан "
              "с траверсой, на которой закреплены подвижные контакты.")
    para(doc, "На переменном токе магнитный поток пульсирует с частотой сети, поэтому якорь стремится "
              "вибрировать и гудит. Чтобы сила притяжения не падала до нуля, на часть полюса сердечника "
              "укладывают короткозамкнутый виток — медное кольцо, которое сдвигает часть потока по фазе.")
    figure(doc, "magnitnaya-sistema.png", 13.0,
           "Рисунок 2 — Магнитная система: слева выключено, справа включено")
    para(doc, "На рисунке 2: *1* — катушка; *2* — подвижная часть магнитопровода (якорь); *3* — неподвижные "
              "силовые контакты; *4* — подвижные силовые контакты.", indent=0)

    heading(doc, "Принцип действия")
    para(doc, "При подаче напряжения на катушку по её обмотке идёт ток, в магнитопроводе возникает "
              "магнитный поток, и якорь притягивается к сердечнику, преодолевая усилие возвратной пружины. "
              "Подвижные контакты прижимаются к неподвижным — силовая цепь замыкается, одновременно "
              "переключаются вспомогательные контакты.")
    para(doc, "При снятии напряжения с катушки магнитный поток исчезает, пружина отводит якорь, "
              "и контакты размыкаются. Возникающая при разрыве цепи дуга гасится в дугогасительной камере.")
    para(doc, "Пока по катушке не идёт ток, контакты замкнутыми не остаются: при исчезновении напряжения "
              "пускатель сам отключает двигатель. Это свойство называют нулевой защитой — после "
              "восстановления напряжения самопроизвольного пуска не происходит.")

    heading(doc, "Управление и самоблокировка")
    para(doc, "Кнопка «Пуск» подаёт напряжение на катушку. Параллельно кнопке включают блок-контакт "
              "*13–14*: после срабатывания он шунтирует кнопку, и цепь катушки остаётся замкнутой, "
              "когда кнопку отпускают. Кнопка «Стоп» разрывает цепь катушки — пускатель отключается.")

    heading(doc, "Тепловое реле")
    para(doc, "В пускателе тепловое реле выполнено на биметаллических пластинах, по которым протекает "
              "ток двигателя. При перегрузке пластина изгибается и размыкает контакт в цепи катушки — "
              "двигатель отключается. От короткого замыкания реле не защищает: для этого служит "
              "автоматический выключатель.")

    heading(doc, "Краткий вывод")
    para(doc, "Магнитный пускатель работает за счёт электромагнита: катушка притягивает якорь и замыкает "
              "силовые контакты, а пружина размыкает их при снятии напряжения. Устройство простое, "
              "но именно оно обеспечивает дистанционное управление, нулевую защиту и защиту двигателя "
              "от перегрузки.")

    heading(doc, "Вопросы для самопроверки")
    questions = [
        "Из каких частей состоит магнитный пускатель?",
        "Что происходит при подаче напряжения на катушку?",
        "Зачем на сердечнике устанавливают короткозамкнутый виток?",
        "Как работает самоблокировка кнопки «Пуск»?",
        "Почему при исчезновении напряжения пускатель отключается сам?",
        "От чего защищает тепловое реле и чего оно не защищает?",
    ]
    for i, q in enumerate(questions, 1):
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


def main():
    os.makedirs(IMG, exist_ok=True)
    shutil.copyfile(os.path.join(SRC, "a5", "a5c04254862bb74d71c92458e31a77a6e82cfac1e994399f06c673d87f5157b9"),
                    os.path.join(IMG, "puskatel-photo.png"))
    shutil.copyfile(os.path.join(SRC, "88", "887fc08479dc250b824537932e5626754e6f3499b55f7c2d8340dc20ddb7133f"),
                    os.path.join(IMG, "magnitnaya-sistema.png"))
    build()


if __name__ == "__main__":
    main()
