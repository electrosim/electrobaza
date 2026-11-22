# -*- coding: utf-8 -*-
"""HTML-версия лекции о квалификационных группах и организации рабочего места."""

import os

import docx

TOPIC = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\7-kvalifikacionnye-gruppy"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Квалификационные группы и организация рабочего места.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
REL = "content/groups/1eo-24/mdk-04-02/semester-3/7-kvalifikacionnye-gruppy"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inline(par):
    out = []
    for run in par.runs:
        text = esc(run.text)
        if not text:
            continue
        if run.italic:
            text = f"<em>{text}</em>"
        if run.bold:
            text = f"<strong>{text}</strong>"
        out.append(text.replace("\n", " "))
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
        blocks.append(("p", inline(par)))

    lines = []
    section = False

    def close_section():
        nonlocal section
        if section:
            lines.append("</section>")
            section = False

    for kind, payload in blocks:
        if kind == "h2":
            close_section()
            lines.append(f"<section>\n  <h2>{esc(payload)}</h2>")
            section = True
        elif kind == "p":
            lines.append(f"  <p>{payload}</p>")
        elif kind == "table":
            table, caption = payload
            lines.append(render_table(table, caption))
    close_section()

    link = (f'<p><a href="{REL}/docs/Лекция — Квалификационные группы и организация рабочего места.docx">'
            "Скачать лекцию в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "blocks:", len(blocks))


if __name__ == "__main__":
    build()
