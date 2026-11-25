# -*- coding: utf-8 -*-
"""HTML-версия лекции «Схемы подключения магнитного пускателя»."""

import os

import docx

TOPIC = r"D:\Electrobaza\content\groups\1me-25\mdk-05-02\6-shemy-podklyucheniya"
DOCX = os.path.join(TOPIC, "docs", "Лекция — Схемы подключения магнитного пускателя.docx")
HTML = os.path.join(TOPIC, "Лекция.html")
IMG = os.path.join(TOPIC, "images")
REL = "content/groups/1me-25/mdk-05-02/6-shemy-podklyucheniya"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


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
        out.append(text)
    return "".join(out)


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
    for child in d.element.body.iterchildren():
        if child.tag.split("}")[-1] != "p":
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
        if par.style.name.startswith("Heading"):
            blocks.append(("h2", text))
            continue
        if text.startswith("Рисунок"):
            for i in range(len(blocks) - 1, -1, -1):
                if blocks[i][0] == "fig" and blocks[i][1][1] is None:
                    blocks[i] = ("fig", (blocks[i][1][0], text))
                    break
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
        elif kind == "fig":
            name, caption = payload
            cap = esc(caption or "")
            lines.append(
                "  <figure>\n"
                f'    <img src="{REL}/images/{name}" alt="{cap}" loading="lazy" />\n'
                f"    <figcaption>{cap}</figcaption>\n"
                "  </figure>"
            )
    close_section()

    link = (f'<p><a href="{REL}/docs/Лекция — Схемы подключения магнитного пускателя.docx">'
            "Скачать лекцию в Word (.docx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n\n" + link + "\n")
    print("saved:", HTML, "images:", len(images), "blocks:", len(blocks))


if __name__ == "__main__":
    build()
