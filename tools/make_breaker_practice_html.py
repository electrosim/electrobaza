# -*- coding: utf-8 -*-
"""HTML-версия практической работы по автоматическому выключателю из Word-файла."""

import os
import re

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

TOPIC = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\circuit-breaker-selection"
DOCX = os.path.join(
    TOPIC, "docs",
    "Практическая работа — Принцип действия и установка автоматического выключателя.docx",
)
HTML = os.path.join(TOPIC, "Практическая работа.html")
IMG_DIR = os.path.join(TOPIC, "images")
REL = "content/groups/1me-25/mdk-05-02/circuit-breaker-selection"

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inline(par):
    """Текст абзаца с сохранением индексов, курсива и жирного."""
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
    """Достаёт картинки из docx в папку images и возвращает имена по порядку."""
    names = []
    index = 0
    for shape in document.inline_shapes:
        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
        r_id = blip.embed
        part = document.part.related_parts[r_id]
        ext = "png" if "png" in part.content_type else "jpg"
        index += 1
        name = f"figure-{index}.{ext}"
        with open(os.path.join(IMG_DIR, name), "wb") as fh:
            fh.write(part.blob)
        names.append(name)
    return names


def convert():
    document = docx.Document(DOCX)
    images = save_images(document)
    img_index = 0

    blocks = []  # (kind, payload)
    pending_caption = None
    pending_table_caption = None
    skip_next = False

    body = document.element.body
    pi = 0
    ti = 0
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            table = document.tables[ti]
            ti += 1
            blocks.append(("table", (table, pending_table_caption)))
            pending_table_caption = None
            continue
        if tag != "p":
            continue
        par = document.paragraphs[pi]
        pi += 1
        text = par.text.strip()
        has_img = bool(par._element.findall(f".//{W}drawing"))

        if pi <= 2:  # заголовок и подзаголовок «Практическая работа»
            continue

        if has_img:
            name = images[img_index]
            img_index += 1
            blocks.append(("fig", (name, None)))
            continue

        if par.style.name.startswith("Heading"):
            blocks.append(("h2", text))
            continue

        if pending_caption is not None:
            blocks.append(("cap", pending_caption))
            pending_caption = None

        if text.startswith("Рисунок "):
            # подпись к предыдущему рисунку
            for i in range(len(blocks) - 1, -1, -1):
                if blocks[i][0] == "fig" and blocks[i][1][1] is None:
                    blocks[i] = ("fig", (blocks[i][1][0], text))
                    break
            continue

        if text.startswith("Таблица "):
            pending_table_caption = text
            continue

        if par.style.name == "List Bullet":
            blocks.append(("ul", inline(par)))
            continue

        plain = re.sub(r"<[^>]+>", "", inline(par))
        if re.match(r"^\d+\.\t", plain):
            blocks.append(("ol", plain.split("\t", 1)[1]))
            continue

        if par.alignment == WD_ALIGN_PARAGRAPH.CENTER and len(text) < 80 and any(ch in text for ch in "=≥≤"):
            blocks.append(("formula", inline(par)))
            continue

        blocks.append(("p", inline(par)))

    # рендер
    lines = []
    section = False
    list_buf = []

    def flush_list():
        nonlocal list_buf
        if not list_buf:
            return
        tag = list_buf[0][0]
        lines.append(f"  <{tag}>")
        for _, html in list_buf:
            lines.append(f"    <li>{html}</li>")
        lines.append(f"  </{tag}>")
        list_buf = []

    def close_section():
        nonlocal section
        flush_list()
        if section:
            lines.append("</section>")
            section = False

    for kind, payload in blocks:
        if kind == "h2":
            close_section()
            lines.append(f"<section>\n  <h2>{esc(payload)}</h2>")
            section = True
        elif kind in ("ul", "ol"):
            if list_buf and list_buf[0][0] != kind:
                flush_list()
            if not list_buf:
                list_buf.append((kind, payload))
            else:
                list_buf.append((kind, payload))
        elif kind == "p":
            flush_list()
            lines.append(f"  <p>{payload}</p>")
        elif kind == "fig":
            flush_list()
            name, caption = payload
            cap_html = esc(caption or "")
            lines.append(
                "  <figure>\n"
                f'    <img src="{REL}/images/{name}" alt="{cap_html}" loading="lazy" />\n'
                f"    <figcaption>{cap_html}</figcaption>\n"
                "  </figure>"
            )
        elif kind == "formula":
            flush_list()
            lines.append(f'  <ul class="formulas">\n    <li>{payload}</li>\n  </ul>')
        elif kind == "table":
            flush_list()
            table, caption = payload
            rows = []
            header = "".join(f"<th>{esc(c.text.strip())}</th>" for c in table.rows[0].cells)
            for row in table.rows[1:]:
                cells = "".join(f"<td>{esc(c.text.strip())}</td>" for c in row.cells)
                rows.append(f"      <tr>{cells}</tr>")
            cap = f"\n    <caption>{esc(caption)}</caption>" if caption else ""
            lines.append(
                '  <div class="table-scroll">\n    <table>' + cap +
                f"\n      <thead><tr>{header}</tr></thead>\n      <tbody>\n" +
                "\n".join(rows) +
                "\n      </tbody>\n    </table>\n  </div>"
            )
    close_section()

    link = (
        f'<p><a href="{REL}/docs/Практическая работа — Принцип действия и установка '
        'автоматического выключателя.docx">Скачать работу в Word (.docx)</a></p>'
    )
    html = "\n".join(lines) + "\n\n" + link + "\n"
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(html)
    print("saved:", HTML)
    print("images:", images)


if __name__ == "__main__":
    convert()
