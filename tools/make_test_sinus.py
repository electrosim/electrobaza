# -*- coding: utf-8 -*-
"""Публикация контрольной работы «Синусоидальный переменный ток» (1А-25 / 3А-26, ОП.03)."""

import os
import shutil

import docx
from docx.oxml.ns import qn

TOPIC = r"D:\Electrobaza\content\groups\1a-25-3a-26-11\op-03\4-kontrolnaya-sinusoidalnyy-tok"
SRC = r"C:\Users\user\Desktop\Контрольная_работа_Синусоидальный_ток.docx"
DOCX = os.path.join(TOPIC, "docs", "Контрольная работа — Синусоидальный переменный ток.docx")
HTML = os.path.join(TOPIC, "Контрольная работа.html")
REL = "content/groups/1a-25-3a-26-11/op-03/4-kontrolnaya-sinusoidalnyy-tok"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def strip_answers(path):
    """Удаляет из копии раздел с ответами (начиная с «Задача 1.»)."""
    d = docx.Document(path)
    body = d.element.body
    started = False
    for child in list(body.iterchildren()):
        tag = child.tag.split("}")[-1]
        if tag == "sectPr":
            continue
        if not started:
            if tag == "p":
                text = "".join(t.text or "" for t in child.iter(f"{W}t")).strip()
                if text.startswith("Задача 1."):
                    started = True
            if not started:
                continue
        body.remove(child)
    d.save(path)


def build():
    os.makedirs(os.path.dirname(DOCX), exist_ok=True)
    shutil.copyfile(SRC, DOCX)
    strip_answers(DOCX)
    print("docx saved, answers removed")

    d = docx.Document(DOCX)
    lines = []
    section = False
    for par in d.paragraphs:
        text = par.text.strip()
        if not text:
            continue
        style = par.style.name
        if style == "ExamTitle":
            continue
        if style == "ExamSubtitle":
            continue
        if text.startswith("ВАРИАНТ"):
            if section:
                lines.append("</section>")
            num = text.split()[-1]
            lines.append(f"<section>\n  <h2>Вариант {num}</h2>")
            section = True
            continue
        if style == "Equation":
            lines.append(f"  <p><em>{esc(text)}</em></p>")
            continue
        lines.append(f"  <p>{esc(text)}</p>")
    if section:
        lines.append("</section>")

    link = (f'<p><a href="{REL}/docs/Контрольная работа — Синусоидальный переменный ток.docx">'
            "Скачать контрольную работу в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("html saved:", HTML)


if __name__ == "__main__":
    build()
