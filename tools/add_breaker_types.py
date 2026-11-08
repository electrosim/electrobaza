# -*- coding: utf-8 -*-
"""Типы B/C/D по вариантам в практической работе."""

import os

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

DOCX = (
    r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\circuit-breaker-selection\docs"
    r"\Практическая работа — Принцип действия и установка автоматического выключателя.docx"
)
TYPES = ["C", "B", "C", "D", "B", "C", "D", "B"]


def set_text(par, text):
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def patch(out):
    d = docx.Document(DOCX)

    # 1. Колонка «Тип» в таблице вариантов
    table = d.tables[0]
    col = table.add_column(Cm(2.2))
    header = table.rows[0].cells[-1]
    header.width = Cm(2.2)
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Тип")
    run.bold = True
    run.font.size = Pt(9.5)
    run.font.name = "Times New Roman"
    for row, typ in zip(table.rows[1:], TYPES):
        cell = row.cells[-1]
        cell.width = Cm(2.2)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(typ)
        run.font.size = Pt(9.5)
        run.font.name = "Times New Roman"
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0

    # 2. Исходные данные: кратность пуска зависит от типа
    for par in d.paragraphs:
        if par.text.strip().startswith("Напряжение сети U = 220 В"):
            set_text(
                par,
                "Напряжение сети U = 220 В. Кратность пускового тока зависит от типа выключателя: "
                "B — 3, C — 6, D — 10. Длительно допустимый ток кабеля принят: 1,5 мм² — 19 А; "
                "2,5 мм² — 27 А; 4 мм² — 38 А. Номинальный ряд выключателей: 6, 10, 16, 20, 25, 32, 40 А. "
                "Тип выключателя и сечение — по таблице вариантов.",
            )
        if par.text.strip().startswith("Вычислить пусковой ток и проверить отсечку"):
            set_text(
                par,
                "Вычислить пусковой ток и проверить отсечку для типа своего варианта (B, C или D).",
            )

    d.save(out)
    print("patched:", out)


if __name__ == "__main__":
    patch(os.environ.get("OUT_DOCX", DOCX))
