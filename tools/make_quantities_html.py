# -*- coding: utf-8 -*-
"""HTML-версия лекции «Работа с физическими величинами»."""

import os
import re

import docx

TOPIC = r"D:\Electrobaza\content\groups\1eo-24\op-08\5-fizicheskie-velichiny"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Работа с физическими величинами.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
REL = "content/groups/1eo-24/op-08/5-fizicheskie-velichiny"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inline(par):
    out = []
    for run in par.runs:
        text = esc(run.text)
        if not text:
            continue
        if run.font.superscript:
            text = f"<sup>{text}</sup>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.bold:
            text = f"<strong>{text}</strong>"
        out.append(text)
    return "".join(out)


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

    blocks = []
    pi = 0
    ti = 0
    pending_caption = None
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            blocks.append(("table", (d.tables[ti], pending_caption)))
            ti += 1
            pending_caption = None
            continue
        if tag != "p":
            continue
        par = d.paragraphs[pi]
        pi += 1
        text = par.text.strip()
        if pi == 1:
            continue
        if par.style.name.startswith("Heading"):
            blocks.append(("h2", text))
            continue
        if text.startswith("Таблица"):
            pending_caption = text
            continue
        if re.match(r"^\d+\.\t", text):
            blocks.append(("ol", re.sub(r"^\d+\.\t", "", inline(par))))
            continue
        blocks.append(("p", inline(par)))

    lines = []
    section = False
    ol_buf = []

    def flush_ol():
        nonlocal ol_buf
        if ol_buf:
            lines.append("  <ol>")
            for item in ol_buf:
                lines.append(f"    <li>{item}</li>")
            lines.append("  </ol>")
            ol_buf = []

    def close_section():
        nonlocal section
        flush_ol()
        if section:
            lines.append("</section>")
            section = False

    for kind, payload in blocks:
        if kind == "h2":
            close_section()
            lines.append(f"<section>\n  <h2>{esc(payload)}</h2>")
            section = True
        elif kind == "ol":
            ol_buf.append(payload)
        elif kind == "p":
            flush_ol()
            lines.append(f"  <p>{payload}</p>")
        elif kind == "table":
            flush_ol()
            table, caption = payload
            lines.append(render_table(table, caption))
    close_section()

    link = (f'<p><a href="{REL}/docs/Лекция — Работа с физическими величинами.docx">'
            "Скачать лекцию в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "blocks:", len(blocks))


if __name__ == "__main__":
    build()
