# -*- coding: utf-8 -*-
"""Контрольная работа «Синусоидальный переменный ток»: 4 варианта, 3 вопроса + 2 задачи."""

import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

TOPIC = r"D:\Electrobaza\content\groups\1a-25-3a-26-11\op-03\4-kontrolnaya-sinusoidalnyy-tok"
DOCX = os.path.join(TOPIC, "docs", "Контрольная работа — Синусоидальный переменный ток.docx")
HTML = os.path.join(TOPIC, "Контрольная работа.html")
REL = "content/groups/1a-25-3a-26-11/op-03/4-kontrolnaya-sinusoidalnyy-tok"
FONT = "Times New Roman"

Q = {
    1: "Какой электрический ток называют переменным?",
    2: ("На каком явлении основана работа генераторов переменного тока и почему переменный ток "
        "получил преимущественное распространение?"),
    3: ("Какой ток называют синусоидальным и каким выражением записывают его мгновенное значение?"),
    4: "Что называют амплитудой (максимальным значением)?",
    5: "Что называют периодом переменного тока и в каких единицах он измеряется?",
    6: "Что называют частотой переменного тока и в каких единицах измеряется?",
    7: ("Какие значения переменных величин называют мгновенными и как они обозначаются "
        "для ЭДС, напряжения, тока и мощности?"),
    8: ("Как связаны действующее и амплитудное значения синусоидального тока?"),
    9: "Что такое сдвиг фаз между напряжением и током?",
}

VARIANTS = [
    {
        "num": 1,
        "questions": [1, 4, 8],
        "tasks": [
            ("Развиваемая генератором ЭДС определяется выражением:",
             "e = 40 sin(100πt + π/6)",
             "Определите начальную фазу ЭДС и её мгновенное значение при t = 0,005 с."),
            ("Определите максимальное значение переменного напряжения, измеряемого вольтметром, "
             "если вольтметр показывает 110 В.", None, None),
        ],
    },
    {
        "num": 2,
        "questions": [2, 5, 9],
        "tasks": [
            ("На зажимах генератора ЭДС имеет синусоидальную форму. Максимальное значение — "
             "120 В, частота — 50 Гц, начальная фаза — −30°. Запишите выражение для мгновенного "
             "значения ЭДС.", None, None),
            ("В электрической цепи протекает переменный ток:",
             "i = 140 sin(100πt + π/3)",
             "Найдите действующее значение, частоту и начальную фазу тока."),
        ],
    },
    {
        "num": 3,
        "questions": [3, 6, 7],
        "tasks": [
            ("Развиваемая генератором ЭДС определяется выражением:",
             "e = 80 sin(200πt + π/4)",
             "Определите начальную фазу ЭДС и её мгновенное значение при t = 0,0025 с."),
            ("В электрической цепи протекает переменный ток:",
             "i = 70,7 sin(120πt + π/2)",
             "Найдите действующее значение, частоту и начальную фазу тока."),
        ],
    },
    {
        "num": 4,
        "questions": [1, 6, 8],
        "tasks": [
            ("На зажимах генератора ЭДС имеет синусоидальную форму. Максимальное значение — "
             "80 В, частота — 100 Гц, начальная фаза — +90°. Запишите выражение для мгновенного "
             "значения ЭДС.", None, None),
            ("Определите максимальное значение переменного напряжения, измеряемого вольтметром, "
             "если вольтметр показывает 36 В.", None, None),
        ],
    },
]


def _rfonts(rpr):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)


def rich(p, text, size=None, bold=None, italic=None):
    for tok in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text):
        if not tok:
            continue
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            r = p.add_run(tok[2:-2]); _rfonts(r._element.get_or_add_rPr()); r.bold = True
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


def build_docx():
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2)

    normal = doc.styles["Normal"]
    _rfonts(normal.element.get_or_add_rPr())
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.05

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    rich(p, "КОНТРОЛЬНАЯ РАБОТА", size=14, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    rich(p, "Тема: «Синусоидальный переменный ток»", size=11, italic=True)

    for v in VARIANTS:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        rich(p, f"ВАРИАНТ {v['num']}", size=12, bold=True)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        rich(p, "**Теоретические вопросы**", size=11)

        for i, qn in enumerate(v["questions"], 1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.first_line_indent = Cm(-0.7)
            p.paragraph_format.space_after = Pt(1)
            rich(p, f"{i}. {Q[qn]}")

        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        rich(p, "**Задачи**", size=11)

        for i, (head, equation, tail) in enumerate(v["tasks"], 1):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.first_line_indent = Cm(-0.7)
            p.paragraph_format.space_after = Pt(1)
            rich(p, f"{i}. {head}")
            if equation:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(1)
                rich(p, equation, size=11, italic=True)
            if tail:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.left_indent = Cm(0.7)
                p.paragraph_format.space_after = Pt(3)
                rich(p, tail)

    os.makedirs(os.path.dirname(DOCX), exist_ok=True)
    doc.save(DOCX)
    print("saved:", DOCX)


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_html():
    lines = []
    for v in VARIANTS:
        lines.append("<section>")
        lines.append(f"  <h2>Вариант {v['num']}</h2>")
        lines.append("  <h3>Теоретические вопросы</h3>\n  <ol>")
        for qn in v["questions"]:
            lines.append(f"    <li>{esc(Q[qn])}</li>")
        lines.append("  </ol>")
        lines.append("  <h3>Задачи</h3>")
        for i, (head, equation, tail) in enumerate(v["tasks"], 1):
            lines.append(f"  <p><strong>{i}.</strong> {esc(head)}</p>")
            if equation:
                lines.append(f"  <p><em>{esc(equation)}</em></p>")
            if tail:
                lines.append(f"  <p>{esc(tail)}</p>")
        lines.append("</section>")

    link = (f'<p><a href="{REL}/docs/Контрольная работа — Синусоидальный переменный ток.docx">'
            "Скачать контрольную работу в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML)


if __name__ == "__main__":
    build_docx()
    build_html()
