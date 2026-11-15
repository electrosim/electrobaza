# -*- coding: utf-8 -*-
"""Публикация практической работы по ПУЭ: перенос docx и HTML-версия."""

import os
import re
import shutil

import docx

SRC_DOCX = (r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\5-pue-i-gost\docs"
            r"\Задание — Правовой анализ ПУЭ (2 листа).docx")
TOPIC = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\6-pravovoy-analiz-pue"
DOCX = os.path.join(TOPIC, "docs", "Задание — Правовой анализ ПУЭ (2 листа).docx")
HTML = os.path.join(TOPIC, "Практическая работа.html")
REL = "content/groups/1eo-24/mdk-04-02/semester-3/6-pravovoy-analiz-pue"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def cell_items(cell):
    items = []
    for p in cell.paragraphs:
        t = p.text.strip()
        if t:
            items.append(t)
    return items


def build():
    os.makedirs(os.path.dirname(DOCX), exist_ok=True)
    if not os.path.exists(DOCX):
        try:
            shutil.move(SRC_DOCX, DOCX)
            print("moved docx")
        except Exception as e:
            shutil.copyfile(SRC_DOCX, DOCX)
            print("copied docx (move failed):", e)

    d = docx.Document(DOCX)
    P = [p.text.strip() for p in d.paragraphs]

    tech = cell_items(d.tables[0].rows[0].cells[0])
    law = cell_items(d.tables[0].rows[0].cells[1])
    q1 = [r.cells[0].text.strip() for r in d.tables[1].rows]
    q2 = [r.cells[0].text.strip() for r in d.tables[2].rows]

    parts = []
    parts.append(f"<section>\n  <h2>{esc(P[1])}</h2>")
    parts.append(f"  <p>{esc(P[2])}</p>")

    parts.append("  <h3>Технические нормы</h3>\n  <ul>")
    for item in tech[1:]:
        parts.append(f"    <li>{esc(item)}</li>")
    parts.append("  </ul>")

    parts.append("  <h3>Правовые нормы</h3>\n  <ul>")
    for item in law[1:]:
        parts.append(f"    <li>{esc(item)}</li>")
    parts.append("  </ul>")
    parts.append("</section>")

    parts.append("<section>\n  <h2>Ключевые понятия</h2>")
    parts.append(f"  <p>{esc(P[4])}</p>")
    parts.append(f"  <p>{esc(P[5])}</p>")
    parts.append("</section>")

    steps = re.split(r"\s+(?=\d\.\s)", P[7])
    parts.append("<section>\n  <h2>Как отвечать (алгоритм)</h2>\n  <ol>")
    for s in steps:
        parts.append(f"    <li>{esc(s.strip())}</li>")
    parts.append("  </ol>\n</section>")

    parts.append("<section>\n  <h2>Задание</h2>")
    for title, situation, questions in (("Вариант 1", P[12], q1), ("Вариант 2", P[17], q2)):
        parts.append(f"  <h3>{esc(title)}</h3>")
        parts.append(f"  <p>{esc(situation)}</p>")
        parts.append("  <ol>")
        for q in questions:
            q = re.sub(r"^\d+\.\s*", "", q)
            parts.append(f"    <li>{esc(q)}</li>")
        parts.append("  </ol>")
    parts.append("</section>")

    link = (f'<p><a href="{REL}/docs/Задание — Правовой анализ ПУЭ (2 листа).docx">'
            "Скачать задание в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(parts) + "\n\n" + link + "\n")
    print("saved:", HTML)


if __name__ == "__main__":
    build()
