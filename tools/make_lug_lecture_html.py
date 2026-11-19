# -*- coding: utf-8 -*-
"""Публикация лекции «Оконцевание жил наконечниками»: перенос docx и HTML."""

import os
import re
import shutil

import docx

SRC = r"C:\Users\user\Downloads\Лекция_Оконцевание_жил_наконечниками.docx"
TOPIC = r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\3-okoncevanie-zhil-nakonechnikami"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Оконцевание жил наконечниками.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
IMG = os.path.join(TOPIC, "images")
REL = "content/groups/1me-25/mdk-05-01/3-okoncevanie-zhil-nakonechnikami"
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


def render_table(table, caption):
    head_cells = "".join(f"<th>{esc(c.text.strip())}</th>" for c in table.rows[0].cells)
    body = []
    for row in table.rows[1:]:
        cells = "".join(f"<td>{esc(c.text.strip())}</td>" for c in row.cells)
        body.append(f"      <tr>{cells}</tr>")
    cap = f"\n    <caption>{esc(caption)}</caption>" if caption else ""
    return (
        '  <div class="table-scroll">\n    <table>' + cap +
        f"\n      <thead><tr>{head_cells}</tr></thead>\n      <tbody>\n" +
        "\n".join(body) +
        "\n      </tbody>\n    </table>\n  </div>"
    )


def build():
    os.makedirs(os.path.dirname(DOCX), exist_ok=True)
    os.makedirs(IMG, exist_ok=True)
    if not os.path.exists(DOCX):
        shutil.copyfile(SRC, DOCX)
        print("copied docx")

    d = docx.Document(DOCX)

    # все картинки по порядку
    all_images = []
    for i, shape in enumerate(d.inline_shapes, 1):
        part = d.part.related_parts[shape._inline.graphic.graphicData.pic.blipFill.blip.embed]
        ext = "png" if "png" in part.content_type else "jpg"
        name = f"figure-{i}.{ext}"
        with open(os.path.join(IMG, name), "wb") as fh:
            fh.write(part.blob)
        all_images.append(name)

    blocks = []
    img_i = 0
    pi = 0
    ti = 0
    pending_table_caption = None
    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            table = d.tables[ti]
            ti += 1
            blocks.append(("table", (table, pending_table_caption)))
            pending_table_caption = None
            continue
        if tag != "p":
            continue
        par = d.paragraphs[pi]
        pi += 1
        text = par.text.strip()
        if pi == 1:
            continue
        blips = par._element.findall(f".//{A}blip")
        if blips:
            names = []
            for _ in blips:
                names.append(all_images[img_i])
                img_i += 1
            blocks.append(("fig", (names, None)))
            continue
        style = par.style.name
        if style in ("Heading 1", "Heading 2"):
            blocks.append(("h2" if style == "Heading 1" else "h3", text))
            continue
        if style == "Caption":
            if text.startswith("Рисунок"):
                for i in range(len(blocks) - 1, -1, -1):
                    if blocks[i][0] == "fig" and blocks[i][1][1] is None:
                        blocks[i] = ("fig", (blocks[i][1][0], text))
                        break
            elif text.startswith("Таблица"):
                pending_table_caption = text
            else:
                blocks.append(("p", inline(par)))
            continue
        if style == "Numbered item" or re.match(r"^\d+\.\t", text):
            plain = par.text
            item = plain.split("\t", 1)[1] if "\t" in plain else re.sub(r"^\d+\.\s*", "", plain)
            blocks.append(("ol", esc(item)))
            continue
        if style == "Source line":
            blocks.append(("source", esc(text)))
            continue
        blocks.append(("p", inline(par)))

    lines = []
    section = False
    buf = []
    src_buf = []

    def flush_list():
        nonlocal buf
        if not buf:
            return
        lines.append("  <ol>")
        for _, html in buf:
            lines.append(f"    <li>{html}</li>")
        lines.append("  </ol>")
        buf = []

    def flush_source():
        nonlocal src_buf
        if not src_buf:
            return
        for html in src_buf:
            lines.append(f'  <p class="source">{html}</p>')
        src_buf = []

    def close_section():
        nonlocal section
        flush_list()
        flush_source()
        if section:
            lines.append("</section>")
            section = False

    for kind, payload in blocks:
        if kind == "h2":
            close_section()
            lines.append(f"<section>\n  <h2>{esc(payload)}</h2>")
            section = True
        elif kind == "h3":
            flush_list()
            flush_source()
            lines.append(f"  <h3>{esc(payload)}</h3>")
        elif kind == "p":
            flush_list()
            flush_source()
            lines.append(f"  <p>{payload}</p>")
        elif kind == "ol":
            flush_source()
            buf.append(("ol", payload))
        elif kind == "source":
            flush_list()
            src_buf.append(payload)
        elif kind == "fig":
            flush_list()
            flush_source()
            names, caption = payload
            cap = esc(caption or "")
            if len(names) == 1:
                inner = f'    <img src="{REL}/images/{names[0]}" alt="{cap}" loading="lazy" />'
            else:
                imgs = "\n".join(
                    f'      <img src="{REL}/images/{n}" alt="{cap}" loading="lazy" />' for n in names
                )
                inner = f'    <div class="figure-pair">\n{imgs}\n    </div>'
            lines.append(
                "  <figure>\n" + inner + "\n"
                f"    <figcaption>{cap}</figcaption>\n"
                "  </figure>"
            )
        elif kind == "table":
            flush_list()
            flush_source()
            table, caption = payload
            lines.append(render_table(table, caption))
    close_section()

    link = (f'<p><a href="{REL}/docs/Лекция — Оконцевание жил наконечниками.docx">'
            "Скачать лекцию в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "images:", len(all_images))


if __name__ == "__main__":
    build()
