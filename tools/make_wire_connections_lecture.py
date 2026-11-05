# -*- coding: utf-8 -*-
"""Сборка Word-лекции «Виды соединений жил проводов и кабелей» (1МЭ-25, МДК 05.01)."""

import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BASE = r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\wire-connections"
IMG = os.path.join(BASE, "images")
OUT = os.path.join(BASE, "docs", "Лекция — Виды соединений жил проводов и кабелей.docx")

FONT = "Times New Roman"


def _set_rfonts(rpr):
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)


def style_font(style, size=None, bold=None, color=None):
    rpr = style.element.get_or_add_rPr()
    _set_rfonts(rpr)
    style.font.name = FONT
    if size is not None:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color is not None:
        style.font.color.rgb = color
        col = rpr.find(qn("w:color"))
        if col is not None:
            for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
                if col.get(qn(attr)) is not None:
                    del col.attrib[qn(attr)]


def run_font(run, size=None, bold=None, italic=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    _set_rfonts(rpr)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def rich(paragraph, text, size=None, bold=None, italic=None):
    """Добавляет текст с разметкой: ~нижний индекс~, ^верхний индекс^, *курсив*."""
    for token in re.split(r"(~[^~]+~|\^[^\^]+\^|\*[^*]+\*)", text):
        if not token:
            continue
        if token.startswith("~") and token.endswith("~") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, italic)
            run.font.subscript = True
        elif token.startswith("^") and token.endswith("^") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, italic)
            run.font.superscript = True
        elif token.startswith("*") and token.endswith("*") and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run_font(run, size, bold, True)
        else:
            run = paragraph.add_run(token)
            run_font(run, size, bold, italic)


doc = Document()

section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.left_margin = Cm(2)
section.right_margin = Cm(2)
section.top_margin = Cm(2)
section.bottom_margin = Cm(2)

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

for style_name in ("List Bullet", "List Number"):
    st = doc.styles[style_name]
    style_font(st, size=11)
    st.paragraph_format.space_after = Pt(1)
    st.paragraph_format.line_spacing = 1.05


def title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    rich(p, text, size=14, bold=True)


def heading(text):
    p = doc.add_paragraph(style="Heading 1")
    rich(p, text)


def para(text, indent=0.8, space_after=2):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent:
        p.paragraph_format.first_line_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.07
    rich(p, text)


def bullets(items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, item)


def numbered(items):
    """Нумерованный список с ручной нумерацией: каждый список начинается с 1."""
    for index, item in enumerate(items, start=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.05
        rich(p, f"{index}.\t{item}")


def formula(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.07
    rich(p, text)


def figure(filename, width_cm, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    shape = p.add_run().add_picture(os.path.join(IMG, filename), width=Cm(width_cm))
    sp_pr = shape._inline.graphic.graphicData.pic.spPr
    line = OxmlElement("a:ln")
    line.set("w", "0")
    line.append(OxmlElement("a:noFill"))
    sp_pr.append(line)
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_before = Pt(1)
    c.paragraph_format.space_after = Pt(4)
    rich(c, caption, size=9, italic=True)


title("Виды соединений жил проводов и кабелей")

para(
    "Соединение жил — одна из самых ответственных операций электромонтажа. Именно в местах "
    "соединений чаще всего появляются повышенное сопротивление, нагрев и потеря контакта, "
    "поэтому к ним предъявляют такие же требования, как к целым жилам. Ниже рассмотрены способы "
    "соединения медных и алюминиевых жил, их возможности и ограничения, порядок выполнения "
    "и контроль качества.",
    indent=0,
)

heading("Требования к соединениям")
para(
    "Правила устройства электроустановок (ПУЭ, п. 2.1.21) разрешают соединять, ответвлять "
    "и оконцовывать жилы проводов и кабелей только с помощью опрессовки, сварки, пайки или "
    "сжимов — винтовых, болтовых и других. Скрутки в этом перечне нет, поэтому как "
    "самостоятельное соединение она не допускается."
)
para(
    "Требования к контактным соединениям устанавливает ГОСТ 10434-82. Главный показатель — "
    "сопротивление контакта: его сравнивают с сопротивлением целого участка жилы такой же длины, "
    "и соединение не должно быть заметно хуже целой жилы."
)
para("К любому соединению предъявляют четыре требования:")
bullets(
    [
        "электрическая надёжность — сопротивление контакта близко к сопротивлению целой жилы;",
        "механическая прочность — соединение выдерживает натяжение провода и вибрацию;",
        "стойкость к коррозии — оксиды меди и алюминия повышают сопротивление, а их прямой "
        "контакт во влажной среде разрушается;",
        "безопасность — место соединения изолируют не хуже самой жилы и размещают в коробке "
        "или зажиме, а не в толще стены.",
    ]
)

heading("Скрутка")
para(
    "Скрутка — самый простой способ: зачищенные жилы складывают и свивают вручную. Сначала она "
    "работает нормально, но со временем ослабевает, окисляется и начинает греться. Нагрев "
    "ускоряет окисление, сопротивление растёт ещё быстрее, и процесс развивается сам. Поэтому "
    "в современном монтаже скрутку применяют только как подготовительную операцию перед пайкой "
    "или сваркой."
)
figure("wire-twist.png", 7.5, "Рисунок 1 — Скрутка проводов")

heading("Пайка")
para(
    "При пайке жилы нагревают и соединяют расплавленным припоем. Медные жилы паяют мягкими "
    "припоями (ПОС-61, ПОС-40) с канифолью или паяльным флюсом: жилы зачищают до блеска, "
    "облуживают, скручивают и пропаивают по всей длине. Припой должен затечь между проволоками, "
    "а не лежать сверху."
)
para(
    "Достоинства пайки — простота и стабильное сопротивление контакта. Недостатки — меньшая "
    "прочность при вибрации, необходимость навыка и опасность перегрева, от которого портится "
    "изоляция. Алюминий паяют специальными припоями и флюсами, поэтому в бытовой проводке его "
    "чаще соединяют другими способами."
)

heading("Сварка")
para(
    "Сварка соединяет жилы в монолит без припоя. Медные жилы сваривают контактным способом: "
    "концы скручивают и оплавляют угольным электродом сварочного аппарата, пока не образуется "
    "шарик расплавленного металла. Остывший шарик зачищают, покрывают лаком и изолируют."
)
para(
    "Сварное соединение самое надёжное: в нём нет припоя и флюса, оно не окисляется "
    "и не ослабевает. Но нужны аппарат, электроды и навык, а работа с открытым пламенем "
    "и раскалённым металлом требует строгого соблюдения правил пожарной безопасности."
)

heading("Опрессовка")
para(
    "При опрессовке жилу вставляют в гильзу или наконечник и обжимают пресс-клещами. Гильзу "
    "подбирают по материалу и сечению жилы: медную — для меди, алюминиевую — для алюминия, "
    "медно-алюминиевую — для перехода с одного металла на другой. Обжатие выполняют в двух "
    "местах по длине гильзы, чтобы металл заполнил все пустоты."
)
para(
    "Правильно опрессованное соединение не греется и выдерживает натяжение. Проверяют его "
    "осмотром и пробным натяжением: жила не должна проворачиваться и вытягиваться из гильзы. "
    "Многопроволочную жилу вставляют в гильзу целиком, не скручивая в жгут."
)
figure("crimp-terminals.jpg", 8.5, "Рисунок 2 — Наконечники, закрепляемые опрессовкой")

heading("Винтовые, болтовые и самозажимные зажимы")
para(
    "Винтовой зажим прижимает жилу винтом. Для надёжности под винт ставят пружинную шайбу, "
    "а многопроволочную жилу предварительно опрессовывают наконечником. Соединение можно "
    "разобрать, но винты нужно периодически протягивать: под давлением металл постепенно "
    "деформируется, и контакт ослабевает. Болтовые соединения применяют для больших сечений и шин."
)
para(
    "Самозажимные (пружинные) клеммы фиксируют жилу пружиной и не требуют протяжки. Монтаж "
    "занимает секунды, но клемма должна соответствовать сечению и материалу жилы, а зачистка — "
    "быть ровной, без надрезов. Медную и алюминиевую жилы нельзя сводить в одном винтовом "
    "зажиме напрямую: образуется гальваническая пара, и контакт разрушается."
)
figure("wire-nuts.jpg", 8.5, "Рисунок 3 — Колпачковые зажимы для соединения жил")

heading("Подготовка жил к соединению")
para(
    "Перед любым соединением с жилы снимают изоляцию. Длину зачистки выбирают по способу: "
    "для клеммного зажима — 8–12 мм, для гильзы — на глубину гильзы, для пайки — по длине "
    "скрутки. Изоляцию снимают стриппером или ножом, стараясь не надрезать металл: надрез "
    "снижает прочность жилы, и провод ломается при изгибе. Зачищенную жилу сразу протирают "
    "и соединяют — на воздухе медь быстро покрывается оксидной плёнкой, и контакт ухудшается."
)

heading("Как выбрать способ соединения")
para("Способ выбирают по материалу и сечению жил, условиям среды и возможности обслуживания.")
bullets(
    [
        "в квартирных коробках чаще применяют самозажимные клеммы или опрессовку;",
        "для ответственных соединений, где нужна максимальная надёжность, — сварку;",
        "там, где соединение должно разбираться, — винтовые и болтовые зажимы;",
        "во влажных помещениях и на улице — герметичные зажимы и гильзы с изоляцией.",
    ]
)

heading("Расчёт контактного соединения")
para(
    "Качество контакта оценивают по сопротивлению. Сопротивление целого участка жилы длиной *l* "
    "и сечением *S*:"
)
formula("*R*~уч~ = *ρ* · *l* / *S*")
para(
    "где *ρ* — удельное сопротивление материала жилы (для меди 0,0175 Ом·мм^2^/м), *l* — длина "
    "участка в метрах, *S* — площадь сечения в квадратных миллиметрах."
)
para(
    "Соединение считается нормальным, если его сопротивление не превышает сопротивление такого "
    "же участка целой жилы:"
)
formula("*R*~кс~ ≤ *R*~уч~")
para("Потери на контакте находят по закону Ома и формуле мощности:")
formula("Δ*U* = *I* · *R*~кс~")
formula("*P* = *I*^2^ · *R*~кс~")
para("Пример. Медная жила сечением 2,5 мм^2^, длина участка 0,05 м:")
formula("*R*~уч~ = 0,0175 · 0,05 / 2,5 = 0,00035 Ом = 0,35 мОм")
para("При токе 16 А и сопротивлении контакта 0,5 мОм:")
formula("Δ*U* = 16 · 0,0005 = 0,008 В = 8 мВ")
formula("*P* = 16^2^ · 0,0005 = 0,128 Вт")
para(
    "Нагрев незаметен. Если контакт окислен и его сопротивление выросло до 0,1 Ом, то при том же "
    "токе падение напряжения составит 1,6 В, а потери — 25,6 Вт: соединение нагревается, "
    "изоляция плавится, возможен пожар."
)

heading("Порядок выполнения и контроль")
numbered(
    [
        "Снять изоляцию на нужную длину, не надрезая жилу.",
        "Зачистить жилу до металлического блеска и сразу приступить к соединению.",
        "Выполнить соединение выбранным способом.",
        "Изолировать место соединения и уложить его в коробку или зажим.",
        "Проверить качество: осмотреть, измерить сопротивление, протянуть винтовые зажимы после "
        "первого нагрева.",
    ]
)

heading("Типичные ошибки")
bullets(
    [
        "скрутка без пайки или сварки;",
        "прямое соединение меди с алюминием;",
        "короткая зачистка: изоляция попадает в гильзу или под винт;",
        "перегрев при пайке и оплавление изоляции;",
        "соединение вне коробки — в толще стены или под штукатуркой.",
    ]
)

heading("Краткий вывод")
para(
    "Надёжность проводки зависит не только от сечения жил, но и от качества соединений. "
    "Допустимые способы — опрессовка, сварка, пайка и зажимы; скрутка как самостоятельное "
    "соединение запрещена. Любое соединение должно иметь малое сопротивление, механическую "
    "прочность, стойкость к коррозии и полноценную изоляцию."
)

heading("Вопросы для самопроверки")
numbered(
    [
        "Какие способы соединения жил допускает ПУЭ и почему скрутка в них не входит?",
        "Чем опрессовка отличается от пайки и сварки?",
        "Как подбирают гильзу для опрессовки?",
        "Почему винтовые зажимы нужно периодически протягивать?",
        "Какое соединение считается нормальным по сопротивлению?",
        "Почему окисленный контакт опаснее при большом токе?",
    ]
)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
