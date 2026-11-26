# -*- coding: utf-8 -*-
"""HTML-версия практической работы по опрессовке гильзами (8 вариантов)."""

import os
import re

import docx

TOPIC = r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\4-opressovka-gilzami"
DOCX = os.path.join(TOPIC, "docs", "Практическая работа — Опрессовка гильзами и изоляция соединений.docx")
HTML = os.path.join(TOPIC, "Практическая работа.html")
REL = "content/groups/1me-25/mdk-05-01/4-opressovka-gilzami"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def render_table(table, caption):
    head = "".join(f"<th>{esc(c.text.strip())}</th>" for c in table.rows[0].cells)
    rows = []
    for row in table.rows[1:]:
        cells = "".join(f"<td>{esc(c.text.strip())}</td>" for c in row.cells)
        rows.append(f"      <tr>{cells}</tr>")
    cap = f"\n    <caption>{esc(caption)}</caption>" if caption else ""
    return (
        '  <div class="table-scroll">\n    <table>' + cap +
        f"\n      <thead><tr>{head}</tr></thead>\n      <tbody>\n" +
        "\n".join(rows) + "\n      </tbody>\n    </table>\n  </div>"
    )


def build():
    d = docx.Document(DOCX)

    intro = []
    variants = []
    current = None
    mode = None

    for par in d.paragraphs:
        text = par.text.strip()
        if not text:
            continue
        if text.startswith("Практическая работа."):
            continue
        if text.startswith("ФИО"):
            continue
        if text == "Краткие сведения":
            mode = "intro" if not intro else "ignore"
            continue
        if text.startswith("Таблица гильз"):
            mode = "table"
            continue
        if text == "Задание":
            mode = "tasks"
            continue
        if text == "Контрольные вопросы":
            mode = "questions"
            continue
        if text.startswith("ВАРИАНТ"):
            num = re.search(r"\d+", text).group(0)
            current = {"num": num, "condition": "", "tasks": [], "questions": []}
            variants.append(current)
            mode = "variant_head"
            continue
        if text.startswith("Условие:"):
            current["condition"] = text[len("Условие:"):].strip()
            mode = "tasks"
            continue
        m = re.match(r"^\d+\.\s+(.*)$", text)
        if m:
            if mode == "questions":
                current["questions"].append(m.group(1))
            elif mode == "tasks":
                current["tasks"].append(m.group(1))
            continue
        if mode == "intro":
            intro.append(text)

    table = d.tables[0] if d.tables else None

    lines = ["<section>", "  <h2>Краткие сведения</h2>"]
    for t in intro:
        lines.append(f"  <p>{esc(t)}</p>")
    lines.append("</section>")

    if table is not None:
        lines.append(render_table(
            table, "Таблица — Гильзы, длина зачистки и число обжимов (учебные значения)"
        ))

    for v in variants:
        lines.append("<section>")
        lines.append(f"  <h2>Вариант {v['num']}</h2>")
        lines.append(f"  <p><em>Условие: {esc(v['condition'])}</em></p>")
        lines.append("  <h3>Задания</h3>\n  <ol>")
        for t in v["tasks"]:
            lines.append(f"    <li>{esc(t)}</li>")
        lines.append("  </ol>")
        lines.append("  <h3>Контрольные вопросы</h3>\n  <ol>")
        for q in v["questions"]:
            lines.append(f"    <li>{esc(q)}</li>")
        lines.append("  </ol>")
        lines.append("</section>")

    link = (f'<p><a href="{REL}/docs/Практическая работа — Опрессовка гильзами и изоляция соединений.docx">'
            "Скачать практическую работу в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "variants:", len(variants), "intro:", len(intro))


if __name__ == "__main__":
    build()
