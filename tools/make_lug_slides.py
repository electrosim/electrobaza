# -*- coding: utf-8 -*-
"""Презентация к лекции об оконцевании: слайды в PNG + HTML для сайта."""

import os
import shutil

TOPIC = r"D:\Electrobaza\content\groups\1me-25\mdk-05-01\3-okoncevanie-zhil-nakonechnikami"
SLIDES_SRC = r"C:\Users\user\AppData\Local\Temp\slides-render-ba68a0aed7c3468c9621d22328108ff9\render"
PPTX_SRC = r"C:\Users\user\Downloads\Оконцевание_жил_наконечниками.pptx"
IMG = os.path.join(TOPIC, "images", "slides")
DOCX = os.path.join(TOPIC, "docs", "Презентация — Оконцевание жил наконечниками.pptx")
HTML = os.path.join(TOPIC, "Презентация.html")
REL = "content/groups/1me-25/mdk-05-01/3-okoncevanie-zhil-nakonechnikami"

TITLES = [
    "Оконцевание жил наконечниками",
    "Как работает опрессовка",
    "Основные виды наконечников",
    "Подбор по жиле и клемме",
    "Инструмент и матрицы",
    "Опрессовка НШВИ",
    "Опрессовка силового наконечника",
    "Алюминиевая жила и медная шина",
    "Контроль качества опрессовки",
    "Задания для закрепления",
]


def build():
    os.makedirs(IMG, exist_ok=True)
    os.makedirs(os.path.dirname(DOCX), exist_ok=True)

    pages = sorted(f for f in os.listdir(SLIDES_SRC) if f.startswith("page-") and f.endswith(".png"))
    assert len(pages) == len(TITLES), (len(pages), len(TITLES))
    names = []
    for i, page in enumerate(pages, 1):
        name = f"slide-{i:02d}.png"
        shutil.copyfile(os.path.join(SLIDES_SRC, page), os.path.join(IMG, name))
        names.append(name)

    if not os.path.exists(DOCX):
        shutil.copyfile(PPTX_SRC, DOCX)

    parts = []
    for i, (name, title) in enumerate(zip(names, TITLES), 1):
        alt = f"Слайд {i} — {title}"
        parts.append(
            f'<section class="slide slide-image" data-slide="{i}">\n'
            f'  <img src="{REL}/images/slides/{name}" alt="{alt}" loading="lazy" />\n'
            "</section>"
        )
    link = (f'<p><a href="{REL}/docs/Презентация — Оконцевание жил наконечниками.pptx">'
            "Скачать презентацию (.pptx)</a></p>")
    with open(HTML, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(parts) + "\n\n" + link + "\n")
    print("saved:", HTML, "slides:", len(names))


if __name__ == "__main__":
    build()
