# -*- coding: utf-8 -*-
"""HTML-версия лекции «Однофазный переменный ток» из Word-файла."""

import os
import re

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

TOPIC = r"D:\Electrobaza\content\groups\1a-25-3a-26-11\op-03\alternating-current"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Однофазный переменный ток.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
IMG = os.path.join(TOPIC, "images")
REL = "content/groups/1a-25-3a-26-11/op-03/alternating-current"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


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


def save_images(document):
    names = []
    for i, shape in enumerate(document.inline_shapes, 1):
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        part = document.part.related_parts[blip.embed]
        ext = "png" if "png" in part.content_type else "jpg"
        name = f"figure-{i}.{ext}"
        with open(os.path.join(IMG, name), "wb") as fh:
            fh.write(part.blob)
        names.append(name)
    return names


def convert():
    d = docx.Document(DOCX)
    images = save_images(d)
    img_i = 0
    blocks = []
    pi = 0
    for child in d.element.body.iterchildren():
        if child.tag.split("}")[-1] != "p":
            continue
        par = d.paragraphs[pi]
        pi += 1
        text = par.text.strip()
        if pi == 1:
            continue
        if par._element.findall(f".//{W}drawing"):
            blocks.append(("fig", (images[img_i], None)))
            img_i += 1
            continue
        if par.style.name.startswith("Heading"):
            blocks.append(("h2", text))
            continue
        if text.startswith("Рисунок "):
            for i in range(len(blocks) - 1, -1, -1):
                if blocks[i][0] == "fig" and blocks[i][1][1] is None:
                    blocks[i] = ("fig", (blocks[i][1][0], text))
                    break
            continue
        if par.style.name == "List Bullet":
            blocks.append(("ul", inline(par)))
            continue
        plain = re.sub(r"<[^>]+>", "", inline(par))
        if re.match(r"^\d+\.\t", plain):
            blocks.append(("ol", plain.split("\t", 1)[1]))
            continue
        if par.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) < 90 and any(c in text for c in "=≥≤≈"):
            blocks.append(("formula", inline(par)))
            continue
        blocks.append(("p", inline(par)))

    lines = []
    section = False
    buf = []

    def flush():
        nonlocal buf
        if not buf:
            return
        tag = buf[0][0]
        lines.append(f"  <{tag}>")
        for _, html in buf:
            lines.append(f"    <li>{html}</li>")
        lines.append(f"  </{tag}>")
        buf = []

    def close_section():
        nonlocal section
        flush()
        if section:
            lines.append("</section>")
            section = False

    for kind, payload in blocks:
        if kind == "h2":
            close_section()
            lines.append(f"<section>\n  <h2>{esc(payload)}</h2>")
            section = True
        elif kind in ("ul", "ol"):
            if buf and buf[0][0] != kind:
                flush()
            buf.append((kind, payload))
        elif kind == "p":
            flush()
            lines.append(f"  <p>{payload}</p>")
        elif kind == "formula":
            flush()
            lines.append(f'  <ul class="formulas">\n    <li>{payload}</li>\n  </ul>')
        elif kind == "fig":
            flush()
            name, caption = payload
            cap = esc(caption or "")
            lines.append(
                "  <figure>\n"
                f'    <img src="{REL}/images/{name}" alt="{cap}" loading="lazy" />\n'
                f"    <figcaption>{cap}</figcaption>\n"
                "  </figure>"
            )
    close_section()

    link = (
        f'<p><a href="{REL}/docs/Лекция — Однофазный переменный ток.docx">'
        "Скачать лекцию в Word (.docx)</a></p>"
    )
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "images:", images)


if __name__ == "__main__":
    convert()
