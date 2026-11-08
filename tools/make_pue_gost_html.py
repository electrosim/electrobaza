# -*- coding: utf-8 -*-
"""HTML-версия лекции по ПУЭ из отредактированного Word-файла."""

import os
import re

import docx

DOCX = (
    r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\pue-and-gost\docs"
    r"\Лекция — Правила устройства электроустановок. Государственный стандарт.docx"
)
HTML = r"D:\Electrobaza\content\groups\1eo-24\mdk-04-02\semester-3\pue-and-gost\Лекция.html"
LINK = (
    '<p><a href="content/groups/1eo-24/mdk-04-02/semester-3/pue-and-gost/docs/'
    'Лекция — Правила устройства электроустановок. Государственный стандарт.docx">'
    "Скачать лекцию в Word (.docx)</a></p>"
)


def esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def blocks_from_docx(path):
    document = docx.Document(path)
    items = []
    for index, par in enumerate(document.paragraphs):
        text = par.text.strip()
        if not text or index == 0:
            continue
        style = par.style.name
        if style.startswith("Heading"):
            items.append(("h2", text))
        elif style == "List Bullet":
            items.append(("ul", text))
        elif re.match(r"^\d+\.\s", text):
            items.append(("ol", re.sub(r"^\d+\.\s*", "", text)))
        else:
            items.append(("p", text))
    return items


def render(items):
    lines = []
    section = False
    buffer = []

    def flush():
        nonlocal buffer
        if not buffer:
            return
        tag = buffer[0][0]
        lines.append(f"  <{tag}>")
        for _, text in buffer:
            lines.append(f"    <li>{esc(text)}</li>")
        lines.append(f"  </{tag}>")
        buffer = []

    for kind, text in items:
        if kind == "h2":
            flush()
            if section:
                lines.append("</section>")
            lines.append(f"<section>\n  <h2>{esc(text)}</h2>")
            section = True
        elif kind in ("ul", "ol"):
            if buffer and buffer[0][0] != kind:
                flush()
            buffer.append((kind, text))
        else:
            flush()
            lines.append(f"  <p>{esc(text)}</p>")
    flush()
    if section:
        lines.append("</section>")
    return "\n".join(lines)


def main():
    html = render(blocks_from_docx(DOCX)) + "\n\n" + LINK + "\n"
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(html)
    print("saved:", HTML)


if __name__ == "__main__":
    main()
