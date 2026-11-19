# -*- coding: utf-8 -*-
"""HTML-версия лекции об организации труда из Word-файла."""

import os
import re

import docx

TOPIC = r"D:\Electrobaza\content\groups\1eo-25\mdk-04-01\3-organizaciya-truda"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Организация труда при выполнении электромонтажных работ.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
IMG = os.path.join(TOPIC, "images")
REL = "content/groups/1eo-25/mdk-04-01/3-organizaciya-truda"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inline(par):
    out = []
    for run in par.runs:
        text = esc(run.text)
        if not text:
            continue
        if run.font.subscript:
            text = f"<sub>{text}</sub>"
        elif run.font.superscript:
            text = f"<sup>{text}</sup>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.bold:
            text = f"<strong>{text}</strong>"
        out.append(text)
    return "".join(out)


def is_note(par):
    pPr = par._p.pPr
    return pPr is not None and pPr.find(f"{W}shd") is not None


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

    images = []
    for i, shape in enumerate(d.inline_shapes, 1):
        part = d.part.related_parts[shape._inline.graphic.graphicData.pic.blipFill.blip.embed]
        ext = "png" if "png" in part.content_type else "jpg"
        name = f"figure-{i}.{ext}"
        with open(os.path.join(IMG, name), "wb") as fh:
            fh.write(part.blob)
        images.append(name)

    blocks = []
    img_i = 0
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
        if par._element.findall(f".//{A}blip"):
            blocks.append(("fig", (images[img_i], None)))
            img_i += 1
            continue
        style = par.style.name
        if style.startswith("Heading"):
            blocks.append(("h2", text))
            continue
        if text.startswith("Рисунок"):
            for i in range(len(blocks) - 1, -1, -1):
                if blocks[i][0] == "fig" and blocks[i][1][1] is None:
                    blocks[i] = ("fig", (blocks[i][1][0], text))
                    break
            continue
        if text.startswith("Таблица"):
            pending_caption = text
            continue
        if is_note(par):
            if text.startswith("Главное"):
                blocks.append(("note", [text.replace("**", "")]))
            elif text.startswith("—"):
                if blocks and blocks[-1][0] == "note":
                    blocks[-1][1].append(re.sub(r"^—\s*", "", text))
                else:
                    blocks.append(("note", [text]))
            continue
        if re.match(r"^\d+\.\s", text):
            blocks.append(("ol", re.sub(r"^\d+\.\s*", "", inline(par))))
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
        elif kind == "note":
            flush_ol()
            title = esc(payload[0])
            items = "\n".join(f"    <li>{esc(x)}</li>" for x in payload[1:])
            block = ['  <div class="note">', f"    <p><strong>{title}</strong></p>"]
            if items:
                block += ["    <ul>", items, "    </ul>"]
            block.append("  </div>")
            lines.append("\n".join(block))
        elif kind == "fig":
            flush_ol()
            name, caption = payload
            cap = esc(caption or "")
            lines.append(
                "  <figure>\n"
                f'    <img src="{REL}/images/{name}" alt="{cap}" loading="lazy" />\n'
                f"    <figcaption>{cap}</figcaption>\n"
                "  </figure>"
            )
        elif kind == "table":
            flush_ol()
            table, caption = payload
            lines.append(render_table(table, caption))
    close_section()

    link = (f'<p><a href="{REL}/docs/Лекция — Организация труда при выполнении электромонтажных работ.docx">'
            "Скачать лекцию в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "images:", len(images), "blocks:", len(blocks))


if __name__ == "__main__":
    build()
